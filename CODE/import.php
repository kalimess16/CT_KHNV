<?php

function khnv_parse_workbook(string $path): array
{
    if (!class_exists('ZipArchive')) {
        throw new RuntimeException('PHP extension ZipArchive chưa được bật.');
    }
    if (!is_file($path)) {
        throw new RuntimeException("Không tìm thấy file: $path");
    }

    $zip = new ZipArchive();
    $status = $zip->open($path);
    if ($status !== true) {
        throw new RuntimeException('Không thể mở file Excel.');
    }

    $sharedStrings = khnv_load_shared_strings($zip);

    $sheetXml = new DOMDocument();
    $sheetXml->preserveWhiteSpace = true;
    $sheetXml->formatOutput = false;
    $sheetXml->loadXML(khnv_read_zip_xml($zip, 'xl/worksheets/sheet1.xml'), LIBXML_NONET);

    $xp = new DOMXPath($sheetXml);
    $xp->registerNamespace('x', KHNV_MAIN_NS);
    $rows = [];
    $cellsByRef = [];
    $formulaRefs = [];

    foreach ($xp->query('//x:sheetData/x:row') as $rowNode) {
        /** @var DOMElement $rowNode */
        $rowNum = (int) $rowNode->getAttribute('r');
        $rows[$rowNum] = [
            'r' => $rowNum,
            'cells' => [],
        ];

        foreach ($xp->query('./x:c', $rowNode) as $cellNode) {
            /** @var DOMElement $cellNode */
            $ref = $cellNode->getAttribute('r');
            preg_match('/^([A-Z]+)(\d+)$/', $ref, $m);
            $col = $m[1] ?? '';
            $type = $cellNode->getAttribute('t');
            $style = $cellNode->getAttribute('s');
            $formulaNode = $xp->query('./x:f', $cellNode)->item(0);
            $formula = $formulaNode ? trim($formulaNode->textContent) : '';

            $value = '';
            if ($type === 's') {
                $vNode = $xp->query('./x:v', $cellNode)->item(0);
                $index = $vNode ? (int) trim($vNode->textContent) : -1;
                $value = ($index >= 0 && array_key_exists($index, $sharedStrings)) ? $sharedStrings[$index] : '';
            } elseif ($type === 'inlineStr' || $type === 'str') {
                $value = khnv_node_text($cellNode);
            } else {
                $vNode = $xp->query('./x:v', $cellNode)->item(0);
                $value = $vNode ? trim($vNode->textContent) : '';
            }

            $isTextCell = ($rowNum <= 2) || in_array($col, ['A', 'B', 'C'], true) || in_array($type, ['inlineStr', 'str'], true);
            $normalizedValue = $isTextCell ? khnv_normalize_cell_text($value) : khnv_clean_number_string($value);

            $cell = [
                'ref' => $ref,
                'row' => $rowNum,
                'col' => $col,
                'type' => $type,
                'style' => $style,
                'formula' => $formula,
                'has_formula' => $formula !== '',
                'value' => $normalizedValue,
                'is_text' => $isTextCell,
            ];

            $rows[$rowNum]['cells'][$col] = $cell;
            $cellsByRef[$ref] = $cell;
            if ($formula !== '') {
                $formulaRefs[] = $ref;
            }
        }
    }

    $zip->close();

    khnv_recalculate_rows($rows);
    foreach ($rows as $rowNum => $rowData) {
        foreach (($rowData['cells'] ?? []) as $col => $cell) {
            $cellsByRef[$cell['ref']]['value'] = $cell['value'];
        }
    }

    $groups = khnv_build_groups($rows);

    return [
        'path' => $path,
        'rows' => $rows,
        'cellsByRef' => $cellsByRef,
        'groups' => $groups,
        'formulaRefs' => $formulaRefs,
        'sharedStrings' => $sharedStrings,
    ];
}

function khnv_recalculate_rows(array &$rows): void
{
    ksort($rows);
    foreach ($rows as $rowNum => &$row) {
        if (!isset($row['cells'])) {
            continue;
        }
        foreach (['F', 'I', 'L', 'O', 'R', 'U', 'X', 'AA', 'AD'] as $col) {
            if (!isset($row['cells'][$col])) {
                continue;
            }
            $left1 = khnv_row_numeric_value($row['cells'], khnv_index_to_col(khnv_col_to_index($col) - 2));
            $left2 = khnv_row_numeric_value($row['cells'], khnv_index_to_col(khnv_col_to_index($col) - 1));
            $row['cells'][$col]['value'] = khnv_clean_number_string($left1 + $left2);
        }
    }
    unset($row);
}

function khnv_row_numeric_value(array $cells, string $col): float
{
    if (!isset($cells[$col])) {
        return 0.0;
    }
    $raw = (string) ($cells[$col]['value'] ?? '');
    if ($raw === '') {
        return 0.0;
    }
    $num = (float) $raw;
    return is_finite($num) ? $num : 0.0;
}

function khnv_build_groups(array $rows): array
{
    ksort($rows);
    $groups = [];
    $current = null;

    foreach ($rows as $rowNum => $row) {
        if ($rowNum < 3) {
            continue;
        }
        $cells = $row['cells'] ?? [];
        $pgd = $cells['B']['value'] ?? '';
        $commune = $cells['C']['value'] ?? '';
        if ($pgd === '' || $commune === '') {
            continue;
        }
        if ($current === null || $current['pgd'] !== $pgd) {
            if ($current !== null) {
                $groups[] = $current;
            }
            $current = [
                'pgd' => $pgd,
                'rows' => [],
                'start' => $rowNum,
                'end' => $rowNum,
            ];
        }
        $current['rows'][] = $rowNum;
        $current['end'] = $rowNum;
    }

    if ($current !== null) {
        $groups[] = $current;
    }

    foreach ($groups as &$group) {
        $group['short'] = khnv_group_short_name($group['pgd']);
        $group['heading'] = khnv_group_heading($group['pgd']);
    }
    unset($group);

    return $groups;
}

function khnv_detect_loan_groups(array $rows): array
{
    $row1 = $rows[1]['cells'] ?? [];
    $row2 = $rows[2]['cells'] ?? [];
    $maxIndex = 0;

    foreach ([$row1, $row2] as $rowCells) {
        foreach ($rowCells as $cell) {
            $col = (string) ($cell['col'] ?? '');
            if ($col === '') {
                continue;
            }
            $index = khnv_col_to_index($col);
            if ($index > $maxIndex) {
                $maxIndex = $index;
            }
        }
    }

    $groups = [];
    for ($start = 4; $start <= $maxIndex; $start += 3) {
        $startCol = khnv_index_to_col($start);
        $adjustCol = khnv_index_to_col($start + 1);
        $targetCol = khnv_index_to_col($start + 2);

        $label = khnv_normalize_cell_text((string) ($row1[$startCol]['value'] ?? ''));
        $sub0 = khnv_normalize_cell_text((string) ($row2[$startCol]['value'] ?? ''));
        $sub1 = khnv_normalize_cell_text((string) ($row2[$adjustCol]['value'] ?? ''));
        $sub2 = khnv_normalize_cell_text((string) ($row2[$targetCol]['value'] ?? ''));

        if ($label === '' && $sub0 === '' && $sub1 === '' && $sub2 === '') {
            continue;
        }

        $groups[] = [
            'start_index' => $start,
            'start' => $startCol,
            'adjust' => $adjustCol,
            'target' => $targetCol,
            'label' => $label,
            'subtitles' => [$sub0, $sub1, $sub2],
        ];
    }

    return $groups;
}

function khnv_group_short_name(string $pgd): string
{
    $pgd = trim($pgd);
    if ($pgd === '') {
        return '';
    }
    if (preg_match('/^Phòng giao dịch NHCSXH\s+(.+)$/u', $pgd, $m)) {
        return khnv_safe_upper(khnv_normalize_text_key($m[1]));
    }
    if (preg_match('/^Hội sở Chi nhánh(?:\s+.+)?$/u', $pgd)) {
        return 'HỘI SỞ CHI NHÁNH';
    }
    return khnv_safe_upper(khnv_normalize_text_key($pgd));
}

function khnv_group_heading(string $pgd): string
{
    $short = khnv_group_short_name($pgd);
    if ($short === 'HỘI SỞ CHI NHÁNH') {
        return 'ĐỐI VỚI CÁC PHƯỜNG CỦA HỘI SỞ CHI NHÁNH';
    }
    if ($short === 'QUẢNG NAM' || $short === 'NÚI THÀNH' || $short === 'HỘI AN' || $short === 'CẨM LỆ' || $short === 'THANH KHÊ' || $short === 'LIÊN CHIỂU' || $short === 'SƠN TRÀ' || $short === 'NGŨ HÀNH SƠN' || $short === 'HÒA VANG') {
        return 'ĐỐI VỚI CÁC XÃ, PHƯỜNG CỦA PGD NHCSXH ' . $short;
    }
    return 'ĐỐI VỚI CÁC XÃ CỦA PGD NHCSXH ' . $short;
}

function khnv_loan_row_has_exportable_value(array $rowData, array $loan): bool
{
    $adjCol = (string) ($loan['adjust'] ?? '');
    $adjust = $adjCol !== '' ? khnv_row_numeric_value($rowData, $adjCol) : 0.0;
    return abs($adjust) > 0.00000001;
}

function khnv_group_has_exportable_content(array $group, array $state): bool
{
    $loanGroups = array_values(array_filter(
        khnv_detect_loan_groups($state['rows']),
        static fn(array $loan): bool => trim((string) ($loan['label'] ?? '')) !== ''
    ));

    foreach (($group['rows'] ?? []) as $rowNum) {
        if (!isset($state['rows'][$rowNum]['cells'])) {
            continue;
        }
        $rowData = $state['rows'][$rowNum]['cells'];
        foreach ($loanGroups as $loan) {
            if (khnv_loan_row_has_exportable_value($rowData, $loan)) {
                return true;
            }
        }
    }

    return false;
}

function khnv_apply_changes(array &$state, array $changes): void
{
    foreach ($changes as $ref => $value) {
        if (!preg_match('/^[A-Z]+\d+$/', (string) $ref)) {
            continue;
        }
        if (!isset($state['cellsByRef'][$ref])) {
            continue;
        }
        preg_match('/^([A-Z]+)(\d+)$/', $ref, $m);
        $row = (int) $m[2];
        $col = $m[1];
        $clean = is_string($value) ? trim($value) : (string) $value;
        $isTextCell = ($row <= 2) || in_array($col, ['A', 'B', 'C'], true);
        $normalized = $isTextCell ? khnv_normalize_cell_text($clean) : khnv_clean_number_string($clean);
        $state['rows'][$row]['cells'][$col]['value'] = $normalized;
        $state['cellsByRef'][$ref]['value'] = $normalized;
    }
    khnv_recalculate_rows($state['rows']);
    foreach ($state['rows'] as $rowNum => $rowData) {
        foreach (($rowData['cells'] ?? []) as $col => $cell) {
            $state['cellsByRef'][$cell['ref']]['value'] = $cell['value'];
        }
    }
}

function khnv_build_cell_map_from_dom(DOMDocument $xml): array
{
    $xp = new DOMXPath($xml);
    $xp->registerNamespace('x', KHNV_MAIN_NS);
    $map = [];
    foreach ($xp->query('//x:sheetData/x:row/x:c') as $cellNode) {
        /** @var DOMElement $cellNode */
        $map[$cellNode->getAttribute('r')] = $cellNode;
    }
    return $map;
}

function khnv_set_cell_text(DOMDocument $dom, DOMElement $cell, string $value): void
{
    try {
        while ($cell->firstChild) {
            $cell->removeChild($cell->firstChild);
        }
        $cell->removeAttribute('t');
        $cell->setAttribute('t', 'inlineStr');
        $is = $dom->createElementNS(KHNV_MAIN_NS, 'is');
        $t = $dom->createElementNS(KHNV_MAIN_NS, 't');
        if ($value !== '' && preg_match('/^\s|\s$/u', $value)) {
            $t->setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
        }
        $t->appendChild($dom->createTextNode($value));
        $is->appendChild($t);
        $cell->appendChild($is);
    } catch (Exception $e) {
        throw new RuntimeException('Lỗi khi đặt giá trị cell text: $e->getMessage()');
    }
}

function khnv_set_cell_number(DOMDocument $dom, DOMElement $cell, string $value, bool $preserveFormula = false): void
{
    try {
        $formulaNode = null;
        if ($preserveFormula) {
            foreach ($cell->childNodes as $child) {
                if ($child instanceof DOMElement && $child->localName === 'f' && $child->namespaceURI === KHNV_MAIN_NS) {
                    $formulaNode = $child->cloneNode(true);
                    break;
                }
            }
        }
        while ($cell->firstChild) {
            $cell->removeChild($cell->firstChild);
        }
        $cell->removeAttribute('t');
        if ($formulaNode instanceof DOMNode) {
            $cell->appendChild($formulaNode);
        }
        if ($value !== '') {
            $v = $dom->createElementNS(KHNV_MAIN_NS, 'v', $value);
            $cell->appendChild($v);
        }
    } catch (Exception $e) {
        throw new RuntimeException('Lỗi khi đặt giá trị cell number: $e->getMessage()');
    }
}

function khnv_set_word_text(DOMElement $container, string $text): void
{
    $nodes = [];
    foreach ($container->getElementsByTagNameNS(KHNV_WORD_NS, 't') as $t) {
        $nodes[] = $t;
    }
    if (!$nodes) {
        return;
    }
    $nodes[0]->textContent = $text;
    for ($i = 1; $i < count($nodes); $i++) {
        $nodes[$i]->textContent = '';
    }
}

function khnv_strip_docx_page_breaks(DOMElement $root): void
{
    $xp = new DOMXPath($root->ownerDocument);
    $xp->registerNamespace('w', KHNV_WORD_NS);
    foreach ($xp->query('.//w:br[@w:type="page"]', $root) as $brNode) {
        if ($brNode instanceof DOMNode && $brNode->parentNode instanceof DOMNode) {
            $brNode->parentNode->removeChild($brNode);
        }
    }
}

function khnv_replace_docx_placeholders(DOMElement $root, array $replacements): void
{
    $xp = new DOMXPath($root->ownerDocument);
    $xp->registerNamespace('w', KHNV_WORD_NS);
    foreach ($xp->query('.//w:p', $root) as $pNode) {
        if (!$pNode instanceof DOMElement) {
            continue;
        }
        $textNodes = [];
        foreach ($xp->query('.//w:t', $pNode) as $tNode) {
            if ($tNode instanceof DOMElement) {
                $textNodes[] = $tNode;
            }
        }
        if (!$textNodes) {
            continue;
        }
        $text = '';
        foreach ($textNodes as $tNode) {
            $text .= $tNode->textContent;
        }
        $updated = false;
        foreach ($replacements as $placeholder => $replacement) {
            $placeholder = (string) $placeholder;
            if (strpos($text, $placeholder) === false) {
                continue;
            }
            $text = str_replace($placeholder, (string) $replacement, $text);
            $updated = true;
        }
        if ($updated) {
            $textNodes[0]->textContent = $text;
            for ($i = 1; $i < count($textNodes); $i++) {
                $textNodes[$i]->textContent = '';
            }
        }
    }
}

function khnv_update_sheet_xml(string $sourcePath, array $state): string
{
    $zip = new ZipArchive();
    if ($zip->open($sourcePath) !== true) {
        throw new RuntimeException('Không thể mở file Excel để cập nhật.');
    }

    $sheetXml = new DOMDocument();
    $sheetXml->preserveWhiteSpace = true;
    $sheetXml->formatOutput = false;
    $sheetXml->loadXML(khnv_read_zip_xml($zip, 'xl/worksheets/sheet1.xml'), LIBXML_NONET);
    $cellMap = khnv_build_cell_map_from_dom($sheetXml);

    foreach ($state['rows'] as $rowNum => $rowData) {
        foreach (($rowData['cells'] ?? []) as $col => $cell) {
            $ref = $cell['ref'];
            if (!isset($cellMap[$ref])) {
                continue;
            }
            $node = $cellMap[$ref];
            $value = $cell['value'] ?? '';
            if ($cell['has_formula']) {
                khnv_set_cell_number($sheetXml, $node, $value, true);
                continue;
            }
            if (in_array($col, ['A', 'B', 'C'], true) || $cell['type'] === 's' || $cell['type'] === 'inlineStr' || $cell['type'] === 'str') {
                khnv_set_cell_text($sheetXml, $node, (string) $value);
            } else {
                khnv_set_cell_number($sheetXml, $node, (string) $value);
            }
        }
    }

    $workbookXml = new DOMDocument();
    $workbookXml->preserveWhiteSpace = true;
    $workbookXml->formatOutput = false;
    $workbookXml->loadXML(khnv_read_zip_xml($zip, 'xl/workbook.xml'), LIBXML_NONET);
    $xp = new DOMXPath($workbookXml);
    $xp->registerNamespace('x', KHNV_MAIN_NS);
    $calcPr = $xp->query('//x:calcPr')->item(0);
    if ($calcPr instanceof DOMElement) {
        $calcPr->setAttribute('fullCalcOnLoad', '1');
        $calcPr->setAttribute('forceFullCalc', '1');
    }

    $tempPath = tempnam(sys_get_temp_dir(), 'khnv_xlsx_');
    if ($tempPath === false) {
        $zip->close();
        throw new RuntimeException('Không thể tạo file tạm.');
    }
    $tempPath .= '.xlsx';
    @unlink($tempPath);

    $out = new ZipArchive();
    if ($out->open($tempPath, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
        $zip->close();
        throw new RuntimeException('Không thể tạo file Excel tạm.');
    }

    for ($i = 0; $i < $zip->numFiles; $i++) {
        $stat = $zip->statIndex($i);
        if (!$stat || !isset($stat['name'])) {
            continue;
        }
        $name = $stat['name'];
        if ($name === 'xl/worksheets/sheet1.xml') {
            $out->addFromString($name, $sheetXml->saveXML());
        } elseif ($name === 'xl/workbook.xml') {
            $out->addFromString($name, $workbookXml->saveXML());
        } else {
            $content = $zip->getFromIndex($i);
            if ($content !== false) {
                $out->addFromString($name, $content);
            }
        }
    }

    $out->close();
    $zip->close();

    return $tempPath;
}

function khnv_save_workbook(string $path, array $changes): void
{
    $state = khnv_parse_workbook($path);
    khnv_apply_changes($state, $changes);
    $tempPath = khnv_update_sheet_xml($path, $state);

    $backup = $path . '.bak';
    @copy($path, $backup);
    try {
        if (file_exists($path)) {
            @unlink($path);
        }
        if (!@rename($tempPath, $path)) {
            if (!@copy($tempPath, $path)) {
                throw new RuntimeException('Không thể ghi đè file Excel.');
            }
            @unlink($tempPath);
        }
    } catch (Throwable $e) {
        if (is_file($backup)) {
            @copy($backup, $path);
        }
        if (is_file($tempPath)) {
            @unlink($tempPath);
        }
        throw $e;
    }
}






