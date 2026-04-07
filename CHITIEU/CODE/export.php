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

function khnv_docx_update_title_years(DOMElement $body, int $reportYear): void
{
    foreach (khnv_docx_title_paragraphs($body) as $paragraph) {
        $text = khnv_node_text($paragraph);
        if ($text === '' || strpos($text, '{{') !== false) {
            continue;
        }

        $updated = khnv_docx_replace_first_year($text, $reportYear);
        if ($updated !== $text) {
            khnv_set_word_text($paragraph, $updated);
        }
    }
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

function khnv_numeric_value_or_null($value): ?float
{
    $clean = khnv_clean_number_string($value);
    if ($clean === '' || !is_numeric($clean)) {
        return null;
    }

    $number = (float) $clean;
    return is_finite($number) ? $number : null;
}

function khnv_format_nullable_report_number(?float $value): string
{
    if ($value === null) {
        return '';
    }

    return khnv_format_report_number($value);
}

function khnv_to_roman(int $number): string
{
    if ($number <= 0) {
        return '';
    }

    $map = [
        'M' => 1000,
        'CM' => 900,
        'D' => 500,
        'CD' => 400,
        'C' => 100,
        'XC' => 90,
        'L' => 50,
        'XL' => 40,
        'X' => 10,
        'IX' => 9,
        'V' => 5,
        'IV' => 4,
        'I' => 1,
    ];

    $result = '';
    foreach ($map as $roman => $value) {
        while ($number >= $value) {
            $result .= $roman;
            $number -= $value;
        }
    }

    return $result;
}

function khnv_export_source_summary(array $state, string $sourceKey): array
{
    $loanGroups = array_values(array_filter(
        khnv_detect_loan_groups($state['rows'] ?? []),
        static fn(array $loan): bool => trim((string) ($loan['label'] ?? '')) !== ''
    ));

    $overallAdjustTotal = 0.0;
    $programs = [];
    foreach ($loanGroups as $loan) {
        $programKey = khnv_normalize_text_key((string) ($loan['label'] ?? ''));
        $programs[$programKey] = [
            'key' => $programKey,
            'label' => (string) ($loan['label'] ?? ''),
            'adjust_total' => 0.0,
            'pgds' => [],
        ];
    }

    $pgdBlocks = [];
    foreach (($state['groups'] ?? []) as $group) {
        $pgdName = trim((string) ($group['pgd'] ?? ''));
        if ($pgdName === '') {
            continue;
        }

        $pgdKey = khnv_normalize_text_key($pgdName);
        $communes = [];
        foreach (($group['rows'] ?? []) as $rowNum) {
            if (!isset($state['rows'][$rowNum]['cells'])) {
                continue;
            }

            $rowData = $state['rows'][$rowNum]['cells'];
            $communeName = trim((string) ($rowData['C']['value'] ?? ''));
            if ($communeName === '') {
                continue;
            }

            $loanItems = [];
            foreach ($loanGroups as $loan) {
                $programKey = khnv_normalize_text_key((string) ($loan['label'] ?? ''));
                $startValue = khnv_numeric_value_or_null($rowData[(string) ($loan['start'] ?? '')]['value'] ?? '');
                $adjustValue = khnv_numeric_value_or_null($rowData[(string) ($loan['adjust'] ?? '')]['value'] ?? '');
                $targetValue = khnv_numeric_value_or_null($rowData[(string) ($loan['target'] ?? '')]['value'] ?? '');

                if ($adjustValue === null || abs($adjustValue) < 0.00000001) {
                    continue;
                }

                $loanItems[] = [
                    'label' => (string) ($loan['label'] ?? ''),
                    'start' => $startValue,
                    'adjust' => $adjustValue,
                    'target' => $targetValue,
                ];

                if (!isset($programs[$programKey]['pgds'][$pgdKey])) {
                    $programs[$programKey]['pgds'][$pgdKey] = [
                        'key' => $pgdKey,
                        'pgd' => $pgdName,
                        'start_total' => 0.0,
                        'adjust_total' => 0.0,
                        'target_total' => 0.0,
                        'communes' => [],
                    ];
                }

                $programs[$programKey]['adjust_total'] += $adjustValue;
                $overallAdjustTotal += $adjustValue;
                $programs[$programKey]['pgds'][$pgdKey]['start_total'] += $startValue ?? 0.0;
                $programs[$programKey]['pgds'][$pgdKey]['adjust_total'] += $adjustValue;
                $programs[$programKey]['pgds'][$pgdKey]['target_total'] += $targetValue ?? 0.0;
                $programs[$programKey]['pgds'][$pgdKey]['communes'][] = [
                    'name' => $communeName,
                    'start' => $startValue,
                    'adjust' => $adjustValue,
                    'target' => $targetValue,
                ];
            }

            if ($loanItems !== []) {
                $communes[] = [
                    'name' => $communeName,
                    'loans' => $loanItems,
                ];
            }
        }

        if ($communes !== []) {
            $pgdBlocks[$pgdKey] = [
                'key' => $pgdKey,
                'pgd' => $pgdName,
                'communes' => $communes,
            ];
        }
    }

    $programList = [];
    foreach ($programs as $program) {
        if (abs((float) ($program['adjust_total'] ?? 0.0)) < 0.00000001 || ($program['pgds'] ?? []) === []) {
            continue;
        }
        $program['pgds'] = array_values($program['pgds']);
        $programList[] = $program;
    }

    return [
        'key' => $sourceKey,
        'label' => khnv_get_workbook_label($sourceKey),
        'title' => khnv_get_workbook_title($sourceKey),
        'report_year' => khnv_detect_report_year($state['rows'] ?? []),
        'pgd_blocks' => array_values($pgdBlocks),
        'overall_adjust_total' => $overallAdjustTotal,
        'programs' => $programList,
    ];
}

function khnv_export_merge_pgd_blocks(array $sources): array
{
    $merged = [];
    foreach (array_keys(khnv_workbook_configs()) as $sourceKey) {
        foreach (($sources[$sourceKey]['pgd_blocks'] ?? []) as $block) {
            $pgdKey = (string) ($block['key'] ?? '');
            if ($pgdKey === '') {
                continue;
            }

            if (!isset($merged[$pgdKey])) {
                $merged[$pgdKey] = [
                    'key' => $pgdKey,
                    'pgd' => (string) ($block['pgd'] ?? ''),
                    'tw' => null,
                    'dp' => null,
                ];
            }

            $merged[$pgdKey][$sourceKey] = $block;
        }
    }

    return array_values($merged);
}

function khnv_build_export_context(array $states): array
{
    $sources = [];
    $reportYear = 0;
    foreach (array_keys(khnv_workbook_configs()) as $sourceKey) {
        $state = $states[$sourceKey] ?? khnv_parse_workbook(khnv_get_workbook_path($sourceKey));
        $summary = khnv_export_source_summary($state, $sourceKey);
        $sources[$sourceKey] = $summary;
        $reportYear = max($reportYear, (int) ($summary['report_year'] ?? 0));
    }

    if ($reportYear <= 0) {
        $reportYear = (int) date('Y');
    }

    return [
        'report_year' => $reportYear,
        'sources' => $sources,
        'pgd_blocks' => khnv_export_merge_pgd_blocks($sources),
    ];
}

function khnv_docx_row_cell_elements(DOMElement $row): array
{
    $xp = new DOMXPath($row->ownerDocument);
    $xp->registerNamespace('w', KHNV_WORD_NS);
    return khnv_docx_row_cells($xp, $row);
}

function khnv_docx_set_row_cells(DOMElement $row, array $values): void
{
    $cells = khnv_docx_row_cell_elements($row);
    foreach ($values as $index => $value) {
        if (!isset($cells[$index])) {
            continue;
        }
        khnv_set_word_text($cells[$index], (string) $value);
    }
}

function khnv_docx_append_clone(DOMElement $parent, DOMElement $template): DOMElement
{
    $clone = $template->cloneNode(true);
    $parent->appendChild($clone);
    return $clone;
}

function khnv_append_pgd_xa_source_rows(
    DOMElement $table,
    DOMElement $sectionTemplate,
    DOMElement $communeTemplate,
    DOMElement $loanTemplate,
    array $sourceBlock
): void {
    if (($sourceBlock['communes'] ?? []) === []) {
        return;
    }

    khnv_docx_append_clone($table, $sectionTemplate);
    foreach (($sourceBlock['communes'] ?? []) as $commune) {
        $communeRow = khnv_docx_append_clone($table, $communeTemplate);
        khnv_docx_set_row_cells($communeRow, [
            0 => '',
            1 => (string) ($commune['name'] ?? ''),
            2 => '',
            3 => '',
        ]);

        $displayIndex = 1;
        foreach (($commune['loans'] ?? []) as $loan) {
            $loanRow = khnv_docx_append_clone($table, $loanTemplate);
            khnv_docx_set_row_cells($loanRow, [
                0 => (string) $displayIndex,
                1 => (string) ($loan['label'] ?? ''),
                2 => khnv_format_nullable_report_number($loan['adjust'] ?? null),
                3 => khnv_format_nullable_report_number($loan['target'] ?? null),
            ]);
            $displayIndex++;
        }
    }
}

function khnv_render_pgd_xa_table(DOMElement $table, array $pgdBlock): void
{
    $xp = new DOMXPath($table->ownerDocument);
    $xp->registerNamespace('w', KHNV_WORD_NS);
    khnv_strip_docx_page_breaks($table);

    $rows = khnv_docx_table_rows($xp, $table);
    if (count($rows) < 7) {
        throw new RuntimeException('Mẫu Dieu_chinh_chi_tieu.docx không đủ dòng để render.');
    }

    $twSectionTemplate = $rows[1];
    $twCommuneTemplate = $rows[2];
    $twLoanTemplate = $rows[3];
    $dpSectionTemplate = $rows[4];
    $dpCommuneTemplate = $rows[5];
    $dpLoanTemplate = $rows[6];

    for ($i = count($rows) - 1; $i >= 1; $i--) {
        $row = $rows[$i];
        if ($row->parentNode instanceof DOMNode) {
            $row->parentNode->removeChild($row);
        }
    }

    if (($pgdBlock['tw']['communes'] ?? []) !== []) {
        khnv_append_pgd_xa_source_rows($table, $twSectionTemplate, $twCommuneTemplate, $twLoanTemplate, $pgdBlock['tw']);
    }
    if (($pgdBlock['dp']['communes'] ?? []) !== []) {
        khnv_append_pgd_xa_source_rows($table, $dpSectionTemplate, $dpCommuneTemplate, $dpLoanTemplate, $pgdBlock['dp']);
    }
}

function khnv_append_tt_pgd_source_rows(
    DOMElement $table,
    DOMElement $totalTemplate,
    DOMElement $programTemplate,
    DOMElement $pgdTemplate,
    DOMElement $communeTemplate,
    array $sourceSummary
): void {
    if (($sourceSummary['programs'] ?? []) === []) {
        return;
    }

    $totalRow = khnv_docx_append_clone($table, $totalTemplate);
    khnv_docx_set_row_cells($totalRow, [
        2 => '',
        3 => khnv_format_nullable_report_number($sourceSummary['overall_adjust_total'] ?? null),
        4 => '',
    ]);

    $programIndex = 1;
    foreach (($sourceSummary['programs'] ?? []) as $program) {
        $programRow = khnv_docx_append_clone($table, $programTemplate);
        khnv_docx_set_row_cells($programRow, [
            0 => khnv_to_roman($programIndex),
            1 => (string) ($program['label'] ?? ''),
            2 => '',
            3 => khnv_format_nullable_report_number($program['adjust_total'] ?? null),
            4 => '',
        ]);

        $pgdIndex = 1;
        foreach (($program['pgds'] ?? []) as $pgd) {
            $pgdRow = khnv_docx_append_clone($table, $pgdTemplate);
            khnv_docx_set_row_cells($pgdRow, [
                0 => (string) $pgdIndex,
                1 => (string) ($pgd['pgd'] ?? ''),
                2 => khnv_format_nullable_report_number($pgd['start_total'] ?? null),
                3 => khnv_format_nullable_report_number($pgd['adjust_total'] ?? null),
                4 => khnv_format_nullable_report_number($pgd['target_total'] ?? null),
            ]);

            foreach (($pgd['communes'] ?? []) as $commune) {
                $communeRow = khnv_docx_append_clone($table, $communeTemplate);
                khnv_docx_set_row_cells($communeRow, [
                    0 => '',
                    1 => (string) ($commune['name'] ?? ''),
                    2 => khnv_format_nullable_report_number($commune['start'] ?? null),
                    3 => khnv_format_nullable_report_number($commune['adjust'] ?? null),
                    4 => khnv_format_nullable_report_number($commune['target'] ?? null),
                ]);
            }

            $pgdIndex++;
        }

        $programIndex++;
    }
}

function khnv_render_tt_pgd_table(DOMElement $table, array $context): void
{
    $xp = new DOMXPath($table->ownerDocument);
    $xp->registerNamespace('w', KHNV_WORD_NS);
    khnv_strip_docx_page_breaks($table);

    $rows = khnv_docx_table_rows($xp, $table);
    if (count($rows) < 9) {
        throw new RuntimeException('Mẫu To_trinh.docx không đủ dòng để render.');
    }

    $twTotalTemplate = $rows[1];
    $twProgramTemplate = $rows[2];
    $twPgdTemplate = $rows[3];
    $twCommuneTemplate = $rows[4];
    $dpTotalTemplate = $rows[5];
    $dpProgramTemplate = $rows[6];
    $dpPgdTemplate = $rows[7];
    $dpCommuneTemplate = $rows[8];

    for ($i = count($rows) - 1; $i >= 1; $i--) {
        $row = $rows[$i];
        if ($row->parentNode instanceof DOMNode) {
            $row->parentNode->removeChild($row);
        }
    }

    khnv_append_tt_pgd_source_rows(
        $table,
        $twTotalTemplate,
        $twProgramTemplate,
        $twPgdTemplate,
        $twCommuneTemplate,
        $context['sources']['tw'] ?? []
    );

    khnv_append_tt_pgd_source_rows(
        $table,
        $dpTotalTemplate,
        $dpProgramTemplate,
        $dpPgdTemplate,
        $dpCommuneTemplate,
        $context['sources']['dp'] ?? []
    );
}

function khnv_open_docx_document(string $templatePath): array
{
    if (!class_exists('ZipArchive')) {
        throw new RuntimeException('PHP extension ZipArchive chưa được bật.');
    }
    if (!is_file($templatePath)) {
        throw new RuntimeException("Không tìm thấy template DOCX: $templatePath");
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

    return [
        'zip' => $zip,
        'document' => $docXml,
        'body' => $body,
    ];
}

function khnv_write_docx_temp(ZipArchive $zip, DOMDocument $docXml, ?string $outputPath = null): string
{
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
            continue;
        }

        $content = $zip->getFromIndex($i);
        if ($content !== false) {
            $out->addFromString($name, $content);
        }
    }

    $out->close();
    $zip->close();

    if ($outputPath !== null) {
        @copy($tempPath, $outputPath);
    }

    return $tempPath;
}

function khnv_export_pgd_xa_docx(array $context, string $templatePath, ?string $outputPath = null): string
{
    if (($context['pgd_blocks'] ?? []) === []) {
        throw new RuntimeException('Không có dữ liệu hợp lệ để xuất mẫu Điều chỉnh chỉ tiêu.');
    }

    $opened = khnv_open_docx_document($templatePath);
    /** @var ZipArchive $zip */
    $zip = $opened['zip'];
    /** @var DOMDocument $docXml */
    $docXml = $opened['document'];
    /** @var DOMElement $body */
    $body = $opened['body'];

    khnv_strip_docx_page_breaks($body);
    khnv_replace_docx_placeholders($body, [
        '{{nam_ke_hoach}}' => (string) ($context['report_year'] ?? date('Y')),
        '{{don_vi}}' => 'Triệu Đồng',
    ]);
    khnv_docx_update_title_years($body, (int) ($context['report_year'] ?? date('Y')));

    $titleParagraphs = khnv_docx_title_paragraphs($body);
    $tableTemplate = null;
    $sectPr = null;
    foreach ($body->childNodes as $child) {
        if (!$child instanceof DOMElement) {
            continue;
        }
        if ($child->localName === 'tbl' && $tableTemplate === null) {
            $tableTemplate = $child;
        } elseif ($child->localName === 'sectPr') {
            $sectPr = $child;
        }
    }

    if ($tableTemplate === null || !$sectPr instanceof DOMElement || $titleParagraphs === []) {
        $zip->close();
        throw new RuntimeException('Mẫu Dieu_chinh_chi_tieu.docx không có bố cục hợp lệ.');
    }

    foreach ($titleParagraphs as $paragraph) {
        if ($paragraph->parentNode instanceof DOMNode) {
            $paragraph->parentNode->removeChild($paragraph);
        }
    }
    if ($tableTemplate->parentNode instanceof DOMNode) {
        $tableTemplate->parentNode->removeChild($tableTemplate);
    }

    $printed = 0;
    foreach (($context['pgd_blocks'] ?? []) as $pgdBlock) {
        if ($printed > 0) {
            $body->insertBefore(khnv_create_page_break_paragraph($docXml), $sectPr);
        }

        foreach ($titleParagraphs as $paragraphTemplate) {
            $paragraphNode = $paragraphTemplate->cloneNode(true);
            khnv_replace_docx_placeholders($paragraphNode, [
                '{{phong_giao_dich}}' => (string) ($pgdBlock['pgd'] ?? ''),
                '{{PHONG_GIAO_DICH}}' => (string) ($pgdBlock['pgd'] ?? ''),
            ]);
            $body->insertBefore($paragraphNode, $sectPr);
        }

        $tableNode = $tableTemplate->cloneNode(true);
        khnv_render_pgd_xa_table($tableNode, $pgdBlock);
        $body->insertBefore($tableNode, $sectPr);
        $printed++;
    }

    khnv_strip_docx_placeholders($body);
    return khnv_write_docx_temp($zip, $docXml, $outputPath);
}

function khnv_export_tt_pgd_docx(array $context, string $templatePath, ?string $outputPath = null): string
{
    $hasPrograms = false;
    foreach (array_keys(khnv_workbook_configs()) as $sourceKey) {
        if (($context['sources'][$sourceKey]['programs'] ?? []) !== []) {
            $hasPrograms = true;
            break;
        }
    }
    if (!$hasPrograms) {
        throw new RuntimeException('Không có dữ liệu hợp lệ để xuất mẫu Tờ trình.');
    }

    $opened = khnv_open_docx_document($templatePath);
    /** @var ZipArchive $zip */
    $zip = $opened['zip'];
    /** @var DOMDocument $docXml */
    $docXml = $opened['document'];
    /** @var DOMElement $body */
    $body = $opened['body'];

    khnv_strip_docx_page_breaks($body);
    khnv_replace_docx_placeholders($body, [
        '{{don_vi}}' => 'Triệu Đồng',
    ]);
    khnv_docx_update_title_years($body, (int) ($context['report_year'] ?? date('Y')));

    $tableTemplate = null;
    foreach ($body->childNodes as $child) {
        if ($child instanceof DOMElement && $child->localName === 'tbl') {
            $tableTemplate = $child;
            break;
        }
    }

    if (!$tableTemplate instanceof DOMElement) {
        $zip->close();
        throw new RuntimeException('Mẫu To_trinh.docx không có bảng hợp lệ.');
    }

    khnv_render_tt_pgd_table($tableTemplate, $context);
    khnv_strip_docx_placeholders($body);
    return khnv_write_docx_temp($zip, $docXml, $outputPath);
}

function khnv_export_document_from_context(array $context, string $mode, ?string $outputPath = null): string
{
    $mode = khnv_normalize_export_mode($mode);
    $config = khnv_get_export_mode_config($mode);

    if ($mode === 'TT') {
        return khnv_export_pgd_xa_docx($context, (string) ($config['template'] ?? ''), $outputPath);
    }
    if ($mode === 'DMDN') {
        return khnv_export_tt_pgd_docx($context, (string) ($config['template'] ?? ''), $outputPath);
    }

    throw new RuntimeException('Che do export khong hop le.');
}

function khnv_stream_download_file(string $path, string $downloadName, string $contentType): void
{
    if (!is_file($path)) {
        throw new RuntimeException('File tai xuong khong ton tai.');
    }

    $fileSize = filesize($path);
    if ($fileSize === false) {
        throw new RuntimeException('Khong the lay kich thuoc file tai xuong.');
    }

    while (ob_get_level() > 0) {
        ob_end_clean();
    }

    header('Content-Type: ' . $contentType);
    header('Content-Disposition: attachment; filename="' . $downloadName . '"');
    header('Content-Length: ' . $fileSize);

    if (readfile($path) === false) {
        throw new RuntimeException('Khong the doc file de tai xuong.');
    }
}

function khnv_export_all_download(array $context): void
{
    $tempFiles = [];
    $zipPath = null;

    try {
        foreach (['TT', 'DMDN'] as $mode) {
            $config = khnv_get_export_mode_config($mode);
            $tempFiles[$mode] = khnv_export_document_from_context($context, $mode);
            if (!is_file($tempFiles[$mode])) {
                throw new RuntimeException('Khong tao duoc file export tam cho che do ' . $mode . '.');
            }
            $tempFiles[$mode . '_name'] = (string) ($config['download_name'] ?? ($mode . '.docx'));
        }

        $zipPath = tempnam(sys_get_temp_dir(), 'khnv_zip_');
        if ($zipPath === false) {
            throw new RuntimeException('Khong the tao file ZIP tam.');
        }
        $zipPath .= '.zip';
        @unlink($zipPath);

        $zip = new ZipArchive();
        if ($zip->open($zipPath, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
            throw new RuntimeException('Khong the tao file ZIP export.');
        }

        $zip->addFile($tempFiles['TT'], $tempFiles['TT_name']);
        $zip->addFile($tempFiles['DMDN'], $tempFiles['DMDN_name']);
        $zip->close();

        $config = khnv_get_export_mode_config('ALL');
        khnv_stream_download_file($zipPath, (string) ($config['download_name'] ?? 'Xuat_chi_tieu.zip'), 'application/zip');
    } finally {
        foreach ($tempFiles as $tempFile) {
            if (is_string($tempFile) && is_file($tempFile)) {
                @unlink($tempFile);
            }
        }
        if ($zipPath && is_file($zipPath)) {
            @unlink($zipPath);
        }
    }
}

function khnv_export_mode_download(array $states, string $mode): void
{
    $mode = khnv_normalize_export_mode($mode);
    $context = khnv_build_export_context($states);

    if ($mode === 'ALL') {
        khnv_export_all_download($context);
        return;
    }

    $config = khnv_get_export_mode_config($mode);
    $tempPath = null;

    try {
        $tempPath = khnv_export_document_from_context($context, $mode);
        khnv_stream_download_file(
            $tempPath,
            (string) ($config['download_name'] ?? 'Xuat_chi_tieu.docx'),
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        );
    } finally {
        if ($tempPath && is_file($tempPath)) {
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
