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
    $mergeRanges = khnv_parse_merge_ranges($xp);

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

    khnv_apply_derived_target_formulas($rows, $cellsByRef, $formulaRefs);
    $zip->close();

    khnv_recalculate_rows($rows);
    foreach ($rows as $rowNum => $rowData) {
        foreach (($rowData['cells'] ?? []) as $col => $cell) {
            $cellsByRef[$cell['ref']] = $cell;
        }
    }

    $groups = khnv_build_groups($rows);

    return [
        'path' => $path,
        'rows' => $rows,
        'cellsByRef' => $cellsByRef,
        'groups' => $groups,
        'formulaRefs' => $formulaRefs,
        'mergeRanges' => $mergeRanges,
        'sharedStrings' => $sharedStrings,
    ];
}

function khnv_parse_merge_ranges(DOMXPath $xp): array
{
    $ranges = [];
    foreach ($xp->query('//x:mergeCells/x:mergeCell') as $mergeNode) {
        if (!$mergeNode instanceof DOMElement) {
            continue;
        }
        $ref = trim($mergeNode->getAttribute('ref'));
        if (!preg_match('/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/', $ref, $m)) {
            continue;
        }
        $ranges[] = [
            'ref' => $ref,
            'start_col' => $m[1],
            'start_row' => (int) $m[2],
            'end_col' => $m[3],
            'end_row' => (int) $m[4],
            'start_index' => khnv_col_to_index($m[1]),
            'end_index' => khnv_col_to_index($m[3]),
        ];
    }

    return $ranges;
}

function khnv_apply_derived_target_formulas(array &$rows, array &$cellsByRef, array &$formulaRefs): void
{
    $loanGroups = khnv_detect_loan_groups($rows);
    if ($loanGroups === []) {
        return;
    }

    $dataStartRow = khnv_detect_data_start_row($rows);
    foreach ($rows as $rowNum => &$row) {
        if ($rowNum < $dataStartRow || !isset($row['cells']['B'], $row['cells']['C'])) {
            continue;
        }

        $pgd = trim((string) ($row['cells']['B']['value'] ?? ''));
        $commune = trim((string) ($row['cells']['C']['value'] ?? ''));
        if ($pgd === '' || $commune === '') {
            continue;
        }

        foreach ($loanGroups as $loan) {
            $startCol = (string) ($loan['start'] ?? '');
            $adjustCol = (string) ($loan['adjust'] ?? '');
            $targetCol = (string) ($loan['target'] ?? '');
            if (
                $startCol === ''
                || $adjustCol === ''
                || $targetCol === ''
                || !isset($row['cells'][$startCol], $row['cells'][$adjustCol], $row['cells'][$targetCol])
            ) {
                continue;
            }

            $cell = $row['cells'][$targetCol];
            if (!empty($cell['has_formula'])) {
                continue;
            }

            $ref = (string) ($cell['ref'] ?? '');
            if ($ref === '') {
                continue;
            }

            $cell['formula'] = $startCol . $rowNum . '+' . $adjustCol . $rowNum;
            $cell['has_formula'] = true;
            $row['cells'][$targetCol] = $cell;
            $cellsByRef[$ref] = $cell;

            if (!in_array($ref, $formulaRefs, true)) {
                $formulaRefs[] = $ref;
            }
        }
    }
    unset($row);
}

function khnv_find_merge_range(array $mergeRanges, int $rowNum, string $col): ?array
{
    $colIndex = khnv_col_to_index($col);
    foreach ($mergeRanges as $range) {
        if (
            $rowNum >= (int) ($range['start_row'] ?? 0)
            && $rowNum <= (int) ($range['end_row'] ?? 0)
            && $colIndex >= (int) ($range['start_index'] ?? 0)
            && $colIndex <= (int) ($range['end_index'] ?? 0)
        ) {
            return $range;
        }
    }

    return null;
}

function khnv_detect_data_start_row(array $rows): int
{
    ksort($rows);
    foreach ($rows as $rowNum => $row) {
        $stt = khnv_normalize_cell_text((string) (($row['cells']['A']['value'] ?? '')));
        if ($stt !== '' && preg_match('/^\d+$/', $stt)) {
            return (int) $rowNum;
        }
    }

    return 3;
}

function khnv_detect_used_columns(array $rows): array
{
    $loanGroups = khnv_detect_loan_groups($rows);
    if ($loanGroups) {
        $columns = ['A', 'B', 'C'];
        foreach ($loanGroups as $group) {
            foreach (['start', 'adjust', 'target'] as $key) {
                $col = (string) ($group[$key] ?? '');
                if ($col !== '' && !in_array($col, $columns, true)) {
                    $columns[] = $col;
                }
            }
        }
        return $columns;
    }

    $maxIndex = 0;
    foreach ($rows as $row) {
        foreach (($row['cells'] ?? []) as $cell) {
            $value = (string) ($cell['value'] ?? '');
            if ($value === '' && empty($cell['has_formula'])) {
                continue;
            }
            $col = (string) ($cell['col'] ?? '');
            if ($col === '') {
                continue;
            }
            $maxIndex = max($maxIndex, khnv_col_to_index($col));
        }
    }

    if ($maxIndex <= 0) {
        return ['A', 'B', 'C'];
    }

    $columns = [];
    for ($index = 1; $index <= $maxIndex; $index++) {
        $columns[] = khnv_index_to_col($index);
    }

    return $columns;
}

function khnv_build_sheet_header_rows(array $rows, array $displayColumns, int $headerRowCount, array $mergeRanges): array
{
    $headerRows = [];
    for ($rowNum = 1; $rowNum <= $headerRowCount; $rowNum++) {
        $cells = [];
        foreach ($displayColumns as $col) {
            $mergeRange = khnv_find_merge_range($mergeRanges, $rowNum, $col);
            if (
                $mergeRange !== null
                && (
                    (int) $mergeRange['start_row'] !== $rowNum
                    || (string) $mergeRange['start_col'] !== $col
                )
            ) {
                continue;
            }

            $colIndex = khnv_col_to_index($col);
            $colspan = 1;
            $rowspan = 1;
            if ($mergeRange !== null) {
                $colspan = ((int) $mergeRange['end_index']) - ((int) $mergeRange['start_index']) + 1;
                $rowspan = min($headerRowCount, (int) $mergeRange['end_row']) - (int) $mergeRange['start_row'] + 1;
            }

            $cells[] = [
                'col' => $col,
                'index' => $colIndex,
                'value' => (string) ($rows[$rowNum]['cells'][$col]['value'] ?? ''),
                'colspan' => max(1, $colspan),
                'rowspan' => max(1, $rowspan),
            ];
        }
        $headerRows[] = [
            'row_num' => $rowNum,
            'cells' => $cells,
        ];
    }

    return $headerRows;
}

function khnv_formula_ref_value(array $rows, string $ref): float
{
    if (!preg_match('/^([A-Z]+)(\d+)$/', $ref, $m)) {
        return 0.0;
    }

    $rowNum = (int) $m[2];
    $col = $m[1];
    if (!isset($rows[$rowNum]['cells'][$col])) {
        return 0.0;
    }

    return khnv_row_numeric_value($rows[$rowNum]['cells'], $col);
}

function khnv_formula_ref_has_value(array $rows, string $ref): bool
{
    if (!preg_match('/^([A-Z]+)(\d+)$/', $ref, $m)) {
        return false;
    }

    $rowNum = (int) $m[2];
    $col = $m[1];
    if (!isset($rows[$rowNum]['cells'][$col])) {
        return false;
    }

    return trim((string) ($rows[$rowNum]['cells'][$col]['value'] ?? '')) !== '';
}

function khnv_formula_sum_range(array $rows, string $startRef, string $endRef): float
{
    if (!preg_match('/^([A-Z]+)(\d+)$/', $startRef, $startMatch)) {
        return 0.0;
    }
    if (!preg_match('/^([A-Z]+)(\d+)$/', $endRef, $endMatch)) {
        return 0.0;
    }

    $startCol = khnv_col_to_index($startMatch[1]);
    $endCol = khnv_col_to_index($endMatch[1]);
    $startRow = (int) $startMatch[2];
    $endRow = (int) $endMatch[2];

    if ($startCol > $endCol) {
        [$startCol, $endCol] = [$endCol, $startCol];
    }
    if ($startRow > $endRow) {
        [$startRow, $endRow] = [$endRow, $startRow];
    }

    $sum = 0.0;
    for ($row = $startRow; $row <= $endRow; $row++) {
        for ($colIndex = $startCol; $colIndex <= $endCol; $colIndex++) {
            $sum += khnv_formula_ref_value($rows, khnv_index_to_col($colIndex) . $row);
        }
    }

    return $sum;
}

function khnv_formula_range_has_value(array $rows, string $startRef, string $endRef): bool
{
    if (!preg_match('/^([A-Z]+)(\d+)$/', $startRef, $startMatch)) {
        return false;
    }
    if (!preg_match('/^([A-Z]+)(\d+)$/', $endRef, $endMatch)) {
        return false;
    }

    $startCol = khnv_col_to_index($startMatch[1]);
    $endCol = khnv_col_to_index($endMatch[1]);
    $startRow = (int) $startMatch[2];
    $endRow = (int) $endMatch[2];

    if ($startCol > $endCol) {
        [$startCol, $endCol] = [$endCol, $startCol];
    }
    if ($startRow > $endRow) {
        [$startRow, $endRow] = [$endRow, $startRow];
    }

    for ($row = $startRow; $row <= $endRow; $row++) {
        for ($colIndex = $startCol; $colIndex <= $endCol; $colIndex++) {
            if (khnv_formula_ref_has_value($rows, khnv_index_to_col($colIndex) . $row)) {
                return true;
            }
        }
    }

    return false;
}

function khnv_formula_strip_outer_parentheses(string $expression): string
{
    $expression = trim($expression);
    while ($expression !== '' && $expression[0] === '(' && substr($expression, -1) === ')') {
        $depth = 0;
        $isWrapper = true;
        $length = strlen($expression);
        for ($i = 0; $i < $length; $i++) {
            $char = $expression[$i];
            if ($char === '(') {
                $depth++;
            } elseif ($char === ')') {
                $depth--;
                if ($depth < 0) {
                    $isWrapper = false;
                    break;
                }
                if ($depth === 0 && $i < ($length - 1)) {
                    $isWrapper = false;
                    break;
                }
            }
        }
        if (!$isWrapper || $depth !== 0) {
            break;
        }
        $expression = trim(substr($expression, 1, -1));
    }

    return $expression;
}

function khnv_formula_split_top_level_args(string $expression): ?array
{
    $args = [];
    $current = '';
    $depth = 0;
    $length = strlen($expression);

    for ($i = 0; $i < $length; $i++) {
        $char = $expression[$i];
        if ($char === '(') {
            $depth++;
            $current .= $char;
            continue;
        }
        if ($char === ')') {
            $depth--;
            if ($depth < 0) {
                return null;
            }
            $current .= $char;
            continue;
        }
        if (($char === ',' || $char === ';') && $depth === 0) {
            $args[] = trim($current);
            $current = '';
            continue;
        }
        $current .= $char;
    }

    if ($depth !== 0) {
        return null;
    }

    $args[] = trim($current);
    return $args;
}

function khnv_formula_parse_function_args(string $expression, string $name): ?array
{
    $expression = khnv_formula_strip_outer_parentheses($expression);
    $prefix = strtoupper($name) . '(';
    if (strtoupper(substr($expression, 0, strlen($prefix))) !== $prefix || substr($expression, -1) !== ')') {
        return null;
    }

    $inner = substr($expression, strlen($prefix), -1);
    return khnv_formula_split_top_level_args($inner);
}

function khnv_formula_is_zero_literal(string $expression): bool
{
    $expression = khnv_formula_strip_outer_parentheses(trim($expression));
    if ($expression === '' || !is_numeric($expression)) {
        return false;
    }

    return abs((float) $expression) < 0.00000001;
}

function khnv_formula_expression_key(string $expression): string
{
    $expression = preg_replace('/\s+/u', '', trim($expression)) ?? trim($expression);
    $expression = khnv_formula_strip_outer_parentheses($expression);
    return strtoupper(ltrim($expression, '+'));
}

function khnv_formula_extract_if_passthrough_expression(string $condition, string $whenTrue, string $whenFalse): ?string
{
    $condition = khnv_formula_strip_outer_parentheses($condition);
    if (!preg_match('/^(.+?)(>=|<=|<>|=|>|<)(.+)$/', $condition, $matches)) {
        return null;
    }

    $left = khnv_formula_strip_outer_parentheses((string) ($matches[1] ?? ''));
    $operator = (string) ($matches[2] ?? '');
    $right = khnv_formula_strip_outer_parentheses((string) ($matches[3] ?? ''));

    $comparableOperators = ['>', '>=', '<', '<='];
    if (!in_array($operator, $comparableOperators, true)) {
        return null;
    }

    $trueKey = khnv_formula_expression_key($whenTrue);
    $falseKey = khnv_formula_expression_key($whenFalse);
    $leftKey = khnv_formula_expression_key($left);
    $rightKey = khnv_formula_expression_key($right);
    $leftIsZero = khnv_formula_is_zero_literal($left);
    $rightIsZero = khnv_formula_is_zero_literal($right);

    if (khnv_formula_is_zero_literal($whenFalse)) {
        if ($trueKey !== '' && $trueKey === $leftKey && $rightIsZero && in_array($operator, ['>', '>='], true)) {
            return $whenTrue;
        }
        if ($trueKey !== '' && $trueKey === $rightKey && $leftIsZero && in_array($operator, ['<', '<='], true)) {
            return $whenTrue;
        }
    }

    if (khnv_formula_is_zero_literal($whenTrue)) {
        if ($falseKey !== '' && $falseKey === $leftKey && $rightIsZero && in_array($operator, ['<', '<='], true)) {
            return $whenFalse;
        }
        if ($falseKey !== '' && $falseKey === $rightKey && $leftIsZero && in_array($operator, ['>', '>='], true)) {
            return $whenFalse;
        }
    }

    return null;
}

function khnv_formula_extract_passthrough_expression(string $formula): ?string
{
    $formula = khnv_formula_strip_outer_parentheses($formula);

    $maxArgs = khnv_formula_parse_function_args($formula, 'MAX');
    if ($maxArgs !== null && count($maxArgs) === 2) {
        if (khnv_formula_is_zero_literal($maxArgs[0] ?? '')) {
            return (string) ($maxArgs[1] ?? '');
        }
        if (khnv_formula_is_zero_literal($maxArgs[1] ?? '')) {
            return (string) ($maxArgs[0] ?? '');
        }
    }

    $ifArgs = khnv_formula_parse_function_args($formula, 'IF');
    if ($ifArgs !== null && count($ifArgs) === 3) {
        return khnv_formula_extract_if_passthrough_expression(
            (string) ($ifArgs[0] ?? ''),
            (string) ($ifArgs[1] ?? ''),
            (string) ($ifArgs[2] ?? '')
        );
    }

    return null;
}

function khnv_formula_unwrap_passthrough_expression(string $formula): string
{
    $current = khnv_formula_strip_outer_parentheses($formula);
    while (true) {
        $next = khnv_formula_extract_passthrough_expression($current);
        if ($next === null) {
            return $current;
        }

        $next = khnv_formula_strip_outer_parentheses($next);
        if (khnv_formula_expression_key($next) === khnv_formula_expression_key($current)) {
            return $current;
        }

        $current = $next;
    }
}

function khnv_formula_term_value(array $rows, string $term, bool &$hasValue = false): ?float
{
    $normalized = strtoupper(trim($term));
    if ($normalized === '') {
        $hasValue = false;
        return null;
    }

    if (preg_match('/^SUM\(([A-Z]+\d+):([A-Z]+\d+)\)$/i', $normalized, $m)) {
        $hasValue = khnv_formula_range_has_value($rows, strtoupper($m[1]), strtoupper($m[2]));
        return khnv_formula_sum_range($rows, strtoupper($m[1]), strtoupper($m[2]));
    }

    if (preg_match('/^[A-Z]+\d+$/i', $normalized)) {
        $hasValue = khnv_formula_ref_has_value($rows, $normalized);
        return khnv_formula_ref_value($rows, $normalized);
    }

    if (preg_match('/^#[A-Z0-9\/!?\-]+$/i', $normalized)) {
        $hasValue = false;
        return 0.0;
    }

    if (is_numeric($normalized)) {
        $hasValue = true;
        $number = (float) $normalized;
        return is_finite($number) ? $number : null;
    }

    return null;
}

function khnv_evaluate_formula(string $formula, array $rows): ?string
{
    $formula = preg_replace('/\s+/u', '', trim($formula)) ?? trim($formula);
    $formula = ltrim($formula, '=');
    $formula = khnv_formula_unwrap_passthrough_expression($formula);
    if ($formula === '') {
        return null;
    }

    $tokenPattern = '/([+\-]?)(SUM\([A-Z]+\d+:[A-Z]+\d+\)|[A-Z]+\d+|#[A-Z0-9\/!?\-]+|\d+(?:\.\d+)?)/i';
    if (preg_match_all($tokenPattern, $formula, $matches, PREG_SET_ORDER) > 0) {
        $consumed = '';
        $total = 0.0;
        $hasAnyValue = false;

        foreach ($matches as $match) {
            $consumed .= $match[0];
            $termHasValue = false;
            $termValue = khnv_formula_term_value($rows, (string) ($match[2] ?? ''), $termHasValue);
            if ($termValue === null) {
                return null;
            }

            if ($termHasValue) {
                $hasAnyValue = true;
            }

            $sign = (($match[1] ?? '') === '-') ? -1.0 : 1.0;
            $total += $sign * $termValue;
        }

        if ($consumed === $formula) {
            return $hasAnyValue ? khnv_clean_number_string($total) : '';
        }
    }

    return null;
}

function khnv_recalculate_rows(array &$rows): void
{
    ksort($rows);
    foreach ($rows as $rowNum => &$row) {
        if (!isset($row['cells'])) {
            continue;
        }
        foreach ($row['cells'] as $col => &$cell) {
            if (empty($cell['has_formula'])) {
                continue;
            }
            $computed = khnv_evaluate_formula((string) ($cell['formula'] ?? ''), $rows);
            if ($computed === null) {
                continue;
            }
            $cell['value'] = $computed;
        }
        unset($cell);
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
        $hasMeaningfulHeader = khnv_is_meaningful_header_text($label)
            || khnv_is_meaningful_header_text($sub0)
            || khnv_is_meaningful_header_text($sub1)
            || khnv_is_meaningful_header_text($sub2);

        if (!$hasMeaningfulHeader) {
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
            $state['cellsByRef'][$cell['ref']] = $cell;
        }
    }
}

function khnv_clear_workbook_data(array &$state): void
{
    $dataStartRow = khnv_detect_data_start_row($state['rows'] ?? []);

    foreach ($state['rows'] as $rowNum => &$rowData) {
        if ($rowNum < $dataStartRow) {
            continue;
        }

        if (!isset($rowData['cells']) || !is_array($rowData['cells'])) {
            continue;
        }

        foreach ($rowData['cells'] as $col => &$cell) {
            if (in_array($col, ['A', 'B', 'C'], true)) {
                continue;
            }
            $cell['value'] = '';
        }
        unset($cell);
    }
    unset($rowData);

    khnv_recalculate_rows($state['rows']);
    foreach ($state['rows'] as $rowData) {
        foreach (($rowData['cells'] ?? []) as $cell) {
            $state['cellsByRef'][$cell['ref']] = $cell;
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
    foreach ($xp->query('descendant-or-self::w:p', $root) as $pNode) {
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

function khnv_write_workbook_state(string $path, array $state): void
{
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

function khnv_save_workbook(string $path, array $changes): void
{
    $state = khnv_parse_workbook($path);
    khnv_apply_changes($state, $changes);
    khnv_write_workbook_state($path, $state);
}

function khnv_clear_workbook(string $path): void
{
    $state = khnv_parse_workbook($path);
    khnv_clear_workbook_data($state);
    khnv_write_workbook_state($path, $state);
}

function khnv_detect_import_workbook_key(string $filename): string
{
    $filename = basename($filename);
    if (!preg_match('/^CTKHNV_(TW|DP).*\.xlsx$/i', $filename, $m)) {
        throw new RuntimeException('File import phai bat dau bang CTKHNV_TW hoac CTKHNV_DP va co duoi .xlsx');
    }

    return strtolower((string) $m[1]);
}

function khnv_validate_import_filename(string $filename, ?string $expectedKey = null): string
{
    $detectedKey = khnv_detect_import_workbook_key($filename);
    if ($expectedKey !== null && khnv_normalize_workbook_key($expectedKey) !== $detectedKey) {
        throw new RuntimeException('Ten file upload khong khop voi workbook dang chon.');
    }

    return $detectedKey;
}

function khnv_import_uploaded_workbook(array $file, ?string $expectedKey = null): string
{
    $error = (int) ($file['error'] ?? UPLOAD_ERR_NO_FILE);
    if ($error !== UPLOAD_ERR_OK) {
        throw new RuntimeException('Chua chon file Excel de import hoac file tai len bi loi.');
    }

    $name = (string) ($file['name'] ?? '');
    $tmpName = (string) ($file['tmp_name'] ?? '');
    if ($name === '' || $tmpName === '') {
        throw new RuntimeException('File import khong hop le.');
    }

    $targetKey = khnv_validate_import_filename($name, $expectedKey);
    $targetPath = khnv_get_workbook_path($targetKey);

    if (!is_uploaded_file($tmpName)) {
        throw new RuntimeException('Khong xac nhan duoc file upload.');
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
                throw new RuntimeException('Khong the luu file import vao thu muc INPUT.');
            }
        }
    } catch (Throwable $e) {
        if (is_file($backup) && !is_file($targetPath)) {
            @copy($backup, $targetPath);
        }
        throw $e;
    }

    return $targetKey;
}






