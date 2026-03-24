<?php

function khnv_docx_table_rows(DOMXPath $xp, DOMElement $table): array
{
    $rows = [];
    foreach ($xp->query('./w:tr', $table) as $rowNode) {
        if ($rowNode instanceof DOMElement) {
            $rows[] = $rowNode;
        }
    }
    return $rows;
}

function khnv_docx_row_cells(DOMXPath $xp, DOMElement $row): array
{
    $cells = [];
    foreach ($xp->query('./w:tc', $row) as $cellNode) {
        if ($cellNode instanceof DOMElement) {
            $cells[] = $cellNode;
        }
    }
    return $cells;
}

function khnv_docx_row_texts(DOMXPath $xp, DOMElement $row): array
{
    $texts = [];
    foreach (khnv_docx_row_cells($xp, $row) as $cellNode) {
        $texts[] = trim(preg_replace('/\s+/u', ' ', khnv_node_text($cellNode)) ?? '');
    }
    return $texts;
}

function khnv_docx_find_table_templates(array $rows, DOMXPath $xp): array
{
    $communeIndex = null;
    $loanIndex = null;

    foreach ($rows as $index => $row) {
        if (!$row instanceof DOMElement) {
            continue;
        }
        $texts = khnv_docx_row_texts($xp, $row);
        if (count($texts) < 4) {
            continue;
        }

        $cell0 = $texts[0] ?? '';
        $cell1 = $texts[1] ?? '';
        $cell2 = $texts[2] ?? '';
        $cell3 = $texts[3] ?? '';

        if ($communeIndex === null) {
            $looksLikeCommune = $cell0 === '' && $cell1 !== '';
            if ($looksLikeCommune) {
                $hasCommunePlaceholder = stripos($cell1, '{{tenxa}}') !== false;
                $hasEmptyValueColumns = $cell2 === '' && $cell3 === '';
                $hasPlaceholderValueColumns = strpos($cell2, '{{') !== false && strpos($cell3, '{{') !== false;
                $looksLikeCommune = $hasCommunePlaceholder || $hasEmptyValueColumns || $hasPlaceholderValueColumns;
            }
            if ($looksLikeCommune) {
                $communeIndex = $index;
                continue;
            }
        }

        if ($communeIndex !== null && $loanIndex === null && preg_match('/^\d+$/', $cell0)) {
            $loanIndex = $index;
            break;
        }
    }

    if ($communeIndex === null) {
        $communeIndex = min(2, max(0, count($rows) - 1));
    }
    if ($loanIndex === null) {
        $loanIndex = min($communeIndex + 1, max(0, count($rows) - 1));
    }

    return [
        'commune' => $communeIndex,
        'loan' => $loanIndex,
    ];
}

function khnv_docx_title_paragraphs(DOMElement $body): array
{
    $paragraphs = [];
    foreach ($body->childNodes as $child) {
        if (!$child instanceof DOMElement) {
            continue;
        }
        if ($child->localName === 'tbl') {
            break;
        }
        if ($child->localName === 'p') {
            $paragraphs[] = $child;
        }
    }
    return $paragraphs;
}

function khnv_docx_replace_first_year(string $text, int $reportYear): string
{
    if (!preg_match('/(?<!\d)(?:19|20)\d{2}(?!\d)/', $text)) {
        return $text;
    }
    return preg_replace('/(?<!\d)(?:19|20)\d{2}(?!\d)/', (string) $reportYear, $text, 1) ?? $text;
}

function khnv_docx_update_title_lines_by_position(DOMElement $body, int $reportYear): void
{
    $paragraphs = khnv_docx_title_paragraphs($body);
    if (count($paragraphs) < 5) {
        return;
    }

    $titleParagraph = $paragraphs[0];
    $noteParagraph = $paragraphs[2];

    if (strpos(khnv_node_text($titleParagraph), '{{') === false) {
        khnv_set_word_text($titleParagraph, khnv_docx_replace_first_year(khnv_node_text($titleParagraph), $reportYear));
    }
    if (strpos(khnv_node_text($noteParagraph), '{{') === false) {
        khnv_set_word_text($noteParagraph, khnv_docx_replace_first_year(khnv_node_text($noteParagraph), $reportYear));
    }
}

function khnv_docx_update_table_headers_by_position(DOMElement $table, string $headerAdjust, string $headerTarget): void
{
    $xp = new DOMXPath($table->ownerDocument);
    $xp->registerNamespace('w', KHNV_WORD_NS);

    $rows = khnv_docx_table_rows($xp, $table);
    if (!$rows) {
        return;
    }

    $headerCells = khnv_docx_row_cells($xp, $rows[0]);
    if (count($headerCells) >= 4) {
        khnv_set_word_text($headerCells[2], $headerAdjust);
        khnv_set_word_text($headerCells[3], $headerTarget);
    }
}

function khnv_render_docx_table_from_state(DOMElement $table, array $group, array $state): void
{
    $xp = new DOMXPath($table->ownerDocument);
    $xp->registerNamespace('w', KHNV_WORD_NS);
    khnv_strip_docx_page_breaks($table);

    $rows = khnv_docx_table_rows($xp, $table);
    if (count($rows) < 4) {
        return;
    }

    $loanGroups = array_values(array_filter(
        khnv_detect_loan_groups($state['rows']),
        static fn(array $loan): bool => trim((string) ($loan['label'] ?? '')) !== ''
    ));
    if (!$loanGroups) {
        return;
    }

    $templateIndexes = khnv_docx_find_table_templates($rows, $xp);
    $communeTemplate = $rows[$templateIndexes['commune']] ?? $rows[2];
    $loanTemplate = $rows[$templateIndexes['loan']] ?? $rows[min(3, count($rows) - 1)];

    for ($i = count($rows) - 1; $i >= $templateIndexes['commune']; $i--) {
        $row = $rows[$i];
        if ($row instanceof DOMElement && $row->parentNode instanceof DOMNode) {
            $row->parentNode->removeChild($row);
        }
    }

    $appendClone = static function (DOMElement $templateRow) use ($table): DOMElement {
        $clone = $templateRow->cloneNode(true);
        $table->appendChild($clone);
        return $clone;
    };

    $setCellText = static function (?DOMNode $cell, string $value): void {
        if ($cell instanceof DOMElement) {
            khnv_set_word_text($cell, $value);
        }
    };

    foreach (($group['rows'] ?? []) as $rowNum) {
        if (!isset($state['rows'][$rowNum])) {
            continue;
        }

        $rowData = $state['rows'][$rowNum]['cells'] ?? [];
        $communeName = trim((string) ($rowData['C']['value'] ?? ''));
        if ($communeName === '') {
            continue;
        }

        $eligibleLoans = [];
        foreach ($loanGroups as $loan) {
            if (khnv_loan_row_has_exportable_value($rowData, $loan)) {
                $eligibleLoans[] = $loan;
            }
        }
        if (!$eligibleLoans) {
            continue;
        }

        $communeRow = $appendClone($communeTemplate);
        khnv_replace_docx_placeholders($communeRow, [
            '{{tenxa}}' => $communeName,
        ]);
        $communeCells = $xp->query('./w:tc', $communeRow);
        if ($communeCells && $communeCells->length >= 4) {
            $setCellText($communeCells->item(0), '');
            $setCellText($communeCells->item(1), $communeName);
            $setCellText($communeCells->item(2), '');
            $setCellText($communeCells->item(3), '');
        }

        $displayStt = 1;
        foreach ($eligibleLoans as $loanIndex => $loan) {
            $adjCol = (string) ($loan['adjust'] ?? '');

            $targetCol = (string) ($loan['target'] ?? '');
            $loanRow = $appendClone($loanTemplate);
            $loanCells = $xp->query('./w:tc', $loanRow);
            if (!$loanCells || $loanCells->length < 4) {
                continue;
            }

            $setCellText($loanCells->item(0), (string) $displayStt);
            $setCellText($loanCells->item(1), (string) ($loan['label'] ?? ''));
            $setCellText($loanCells->item(2), khnv_format_report_number($rowData[$adjCol]['value'] ?? ''));
            $setCellText($loanCells->item(3), khnv_format_report_number($rowData[$targetCol]['value'] ?? ''));
            $displayStt++;
        }
    }
}

function khnv_update_docx_title_lines(DOMElement $body, int $reportYear): void
{
    khnv_replace_docx_placeholders($body, [
        '{{TIEU_DE_1}}' => 'DANH MỤC ĐIỀU CHỈNH CHỈ TIÊU KẾ HOẠCH TÍN DỤNG NĂM ' . $reportYear,
        '{{GHI_CHU_1}}' => '(Kèm theo Quyết định số /QĐ-BĐD ngày tháng năm ' . $reportYear,
        '{{GHI_CHU_2}}' => 'của Trưởng Ban đại diện HĐQT NHCSXH thành phố Đà Nẵng)',
        '{{DON_VI}}' => 'Đơn vị: triệu đồng',
    ]);
    khnv_docx_update_title_lines_by_position($body, $reportYear);
}

function khnv_strip_docx_placeholders(DOMElement $root): void
{
    $xp = new DOMXPath($root->ownerDocument);
    $xp->registerNamespace('w', KHNV_WORD_NS);
    foreach ($xp->query('.//w:t', $root) as $tNode) {
        if (!$tNode instanceof DOMElement) {
            continue;
        }
        $text = $tNode->textContent;
        if (strpos($text, '{{') === false) {
            continue;
        }
        $clean = preg_replace('/\{\{[^}]+\}\}/u', '', $text) ?? $text;
        $tNode->textContent = $clean;
    }
}

function khnv_create_page_break_paragraph(DOMDocument $dom): DOMElement
{
    $p = $dom->createElementNS(KHNV_WORD_NS, 'w:p');
    $r = $dom->createElementNS(KHNV_WORD_NS, 'w:r');
    $br = $dom->createElementNS(KHNV_WORD_NS, 'w:br');
    $br->setAttributeNS(KHNV_WORD_NS, 'w:type', 'page');
    $r->appendChild($br);
    $p->appendChild($r);
    return $p;
}



function khnv_validate_import_filename(string $filename): void
{
    $filename = basename($filename);
    if (!preg_match('/^CTKHNV.*\.xlsx$/i', $filename)) {
        throw new RuntimeException('File import phải bắt đầu bằng CTKHNV và có đuôi .xlsx');
    }
}

function khnv_import_uploaded_workbook(array $file, string $targetPath = KHNV_INPUT_XLSX): void
{
    $error = (int) ($file['error'] ?? UPLOAD_ERR_NO_FILE);
    if ($error !== UPLOAD_ERR_OK) {
        throw new RuntimeException('Chưa chọn file Excel để import hoặc file tải lên bị lỗi.');
    }

    $name = (string) ($file['name'] ?? '');
    $tmpName = (string) ($file['tmp_name'] ?? '');
    if ($name === '' || $tmpName === '') {
        throw new RuntimeException('File import không hợp lệ.');
    }

    khnv_validate_import_filename($name);

    if (!is_uploaded_file($tmpName)) {
        throw new RuntimeException('Không xác nhận được file upload.');
    }

    khnv_parse_workbook($tmpName);

    $backup = $targetPath . '.bak';
    if (is_file($targetPath)) {
        @copy($targetPath, $backup);
    }

    try {
        if (file_exists($targetPath)) {
            @unlink($targetPath);
        }

        if (!@move_uploaded_file($tmpName, $targetPath)) {
            if (!@copy($tmpName, $targetPath)) {
                if (is_file($backup)) {
                    @copy($backup, $targetPath);
                }
                throw new RuntimeException('Không thể lưu file import vào INPUT/TEST.xlsx.');
            }
        }
    } catch (Throwable $e) {
        if (is_file($backup) && !is_file($targetPath)) {
            @copy($backup, $targetPath);
        }
        throw $e;
    }
}

function khnv_export_docx_from_state(array $state, string $templatePath, ?string $outputPath = null): string
{
    if (!class_exists('ZipArchive')) {
        throw new RuntimeException('PHP extension ZipArchive chưa được bật.');
    }
    if (!is_file($templatePath)) {
        throw new RuntimeException('Không tìm thấy template DOCX: $templatePath');
    }

    $zip = new ZipArchive();
    if ($zip->open($templatePath) !== true) {
        throw new RuntimeException('Không thể mở template DOCX.');
    }

    $docXml = new DOMDocument();
    $docXml->preserveWhiteSpace = true;
    $docXml->formatOutput = false;
    $docXml->loadXML(khnv_read_zip_xml($zip, 'word/document.xml'), LIBXML_NONET);
    $xp = new DOMXPath($docXml);
    $xp->registerNamespace('w', KHNV_WORD_NS);
    $body = $xp->query('/w:document/w:body')->item(0);
    if (!$body instanceof DOMElement) {
        $zip->close();
        throw new RuntimeException('DOCX template không hợp lệ.');
    }
    khnv_strip_docx_page_breaks($body);

    $reportYear = khnv_detect_report_year($state['rows']);
    khnv_update_docx_title_lines($body, $reportYear);
    $loanGroups = array_values(array_filter(
        khnv_detect_loan_groups($state['rows']),
        static fn(array $loan): bool => trim((string) ($loan['label'] ?? '')) !== ''
    ));
    $headerAdjust = '';
    $headerTarget = '';
    if ($loanGroups) {
        $headerAdjust = (string) ($loanGroups[0]['subtitles'][1] ?? '');
        $headerTarget = (string) ($loanGroups[0]['subtitles'][2] ?? '');
    }
    if (!khnv_is_meaningful_header_text($headerAdjust)) {
        $headerAdjust = 'Điều chỉnh tăng trưởng';
    }
    if (!khnv_is_meaningful_header_text($headerTarget)) {
        $headerTarget = 'Chỉ tiêu kế hoạch năm ' . $reportYear;
    }
    khnv_replace_docx_placeholders($body, [
        '{{DIEU_CHINH_TANG_TRUONG}}' => $headerAdjust,
        '{{CHI_TIEU_KE_HOACH}}' => $headerTarget,
    ]);

    $allTitleParagraphs = khnv_docx_title_paragraphs($body);
    $groupHeadingTemplate = null;
    $tableTemplate = null;
    $sectPr = null;
    foreach ($body->childNodes as $child) {
        if (!$child instanceof DOMElement) {
            continue;
        }
        if ($child->localName === 'p') {
            $text = khnv_node_text($child);
            if ($groupHeadingTemplate === null && stripos($text, '{{PHONG_GIAO_DICH}}') !== false) {
                $groupHeadingTemplate = $child;
            }
        } elseif ($child->localName === 'tbl' && $tableTemplate === null) {
            $tableTemplate = $child;
        } elseif ($child->localName === 'sectPr') {
            $sectPr = $child;
        }
    }

    if (!$groupHeadingTemplate instanceof DOMElement && count($allTitleParagraphs) >= 2) {
        $groupHeadingTemplate = $allTitleParagraphs[1];
    }

    $headingParagraphIndex = array_search($groupHeadingTemplate, $allTitleParagraphs, true);
    if ($headingParagraphIndex === false) {
        $headingParagraphIndex = 1;
    }
    $titleBeforeHeading = array_slice($allTitleParagraphs, 0, $headingParagraphIndex);
    $titleAfterHeading = array_slice($allTitleParagraphs, $headingParagraphIndex + 1);

    if (!$groupHeadingTemplate instanceof DOMElement || !$tableTemplate instanceof DOMElement) {
        $zip->close();
        throw new RuntimeException('DOCX template không có block PGD hợp lệ.');
    }
    if (!$sectPr instanceof DOMElement) {
        $zip->close();
        throw new RuntimeException('DOCX template không có sectPr hợp lệ.');
    }

    $groups = $state['groups'];
    $printedAnyGroup = false;
    foreach ($groups as $index => $group) {
        if (!khnv_group_has_exportable_content($group, $state)) {
            continue;
        }
        $headingNode = $index === 0 ? $groupHeadingTemplate : $groupHeadingTemplate->cloneNode(true);
        $tableNode = $index === 0 ? $tableTemplate : $tableTemplate->cloneNode(true);

        $headingText = khnv_group_heading((string) ($group['pgd'] ?? ''));
        $headingUsesPlaceholder = stripos(khnv_node_text($headingNode), '{{PHONG_GIAO_DICH}}') !== false;
        khnv_replace_docx_placeholders($headingNode, [
            '{{PHONG_GIAO_DICH}}' => $headingText,
        ]);
        if (!$headingUsesPlaceholder) {
            khnv_set_word_text($headingNode, $headingText);
        }
        khnv_replace_docx_placeholders($tableNode, [
            '{{DIEU_CHINH_TANG_TRUONG}}' => $headerAdjust,
            '{{CHI_TIEU_KE_HOACH}}' => $headerTarget,
        ]);
        khnv_docx_update_table_headers_by_position($tableNode, $headerAdjust, $headerTarget);
        khnv_render_docx_table_from_state($tableNode, $group, $state);

        if ($printedAnyGroup) {
            $body->insertBefore(khnv_create_page_break_paragraph($docXml), $sectPr);
            foreach ($titleBeforeHeading as $paragraph) {
                $body->insertBefore($paragraph->cloneNode(true), $sectPr);
            }
            $body->insertBefore($headingNode, $sectPr);
            foreach ($titleAfterHeading as $paragraph) {
                $body->insertBefore($paragraph->cloneNode(true), $sectPr);
            }
            $body->insertBefore($tableNode, $sectPr);
        } else {
            $insertHeadingBefore = $titleAfterHeading[0] ?? $sectPr;
            $body->insertBefore($headingNode, $insertHeadingBefore);
            $body->insertBefore($tableNode, $sectPr);
        }
        $printedAnyGroup = true;
    }

    if (!$printedAnyGroup) {
        $zip->close();
        throw new RuntimeException('Không có dữ liệu hợp lệ để xuất DOCX.');
    }

    khnv_strip_docx_placeholders($body);

    $tempPath = tempnam(sys_get_temp_dir(), 'khnv_docx_');
    if ($tempPath === false) {
        $zip->close();
        throw new RuntimeException('Không thể tạo file DOCX tạm.');
    }
    $tempPath .= '.docx';
    @unlink($tempPath);

    $out = new ZipArchive();
    if ($out->open($tempPath, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
        $zip->close();
        throw new RuntimeException('Không thể ghi DOCX tạm.');
    }

    for ($i = 0; $i < $zip->numFiles; $i++) {
        $stat = $zip->statIndex($i);
        if (!$stat || !isset($stat['name'])) {
            continue;
        }
        $name = $stat['name'];
        if ($name === 'word/document.xml') {
            $out->addFromString($name, $docXml->saveXML());
        } else {
            $content = $zip->getFromIndex($i);
            if ($content !== false) {
                $out->addFromString($name, $content);
            }
        }
    }

    $out->close();
    $zip->close();

    if ($outputPath !== null) {
        @copy($tempPath, $outputPath);
    }

    return $tempPath;
}

function khnv_export_docx_download(array $state, string $templatePath, string $downloadName): void
{
    $tempPath = null;
    try {
        $tempPath = khnv_export_docx_from_state($state, $templatePath);
        if (!file_exists($tempPath)) {
            throw new RuntimeException('File DOCX tạm không được tạo.');
        }
        $fileSize = filesize($tempPath);
        if ($fileSize === false) {
            throw new RuntimeException('Không thể lấy kích thước file.');
        }
        while (ob_get_level() > 0) {
            ob_end_clean();
        }
        header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        header('Content-Disposition: attachment; filename="' . $downloadName . '"');
        header('Content-Length: ' . $fileSize);
        if (readfile($tempPath) === false) {
            throw new RuntimeException('Không thể đọc file để tải xuống.');
        }
    } finally {
        if ($tempPath && file_exists($tempPath)) {
            @unlink($tempPath);
        }
    }
}

function khnv_collect_payload(): array
{
    $raw = $_POST['payload'] ?? '';
    if ($raw === '') {
        return [];
    }
    $decoded = json_decode($raw, true);
    return is_array($decoded) ? $decoded : [];
}

