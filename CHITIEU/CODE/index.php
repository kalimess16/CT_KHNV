<?php
declare(strict_types=1);

require __DIR__ . '/data.php';
khnv_redirect_localhost_to_ip();

$publicPrefix = defined('KHNV_PUBLIC_PREFIX') ? KHNV_PUBLIC_PREFIX : '';
$homeHref = defined('KHNV_HOME_HREF') ? KHNV_HOME_HREF : '../../index.php';
$selfUrl = defined('KHNV_SELF_URL') ? KHNV_SELF_URL : 'index.php';

function khnv_h(string $value): string
{
    return htmlspecialchars($value, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');
}

function khnv_row_has_content(array $row): bool
{
    foreach (($row['cells'] ?? []) as $cell) {
        $value = (string) ($cell['value'] ?? '');
        if ($value !== '') {
            return true;
        }
    }
    return false;
}

function khnv_input_value(array $cell): string
{
    return (string) ($cell['value'] ?? '');
}

function khnv_group_subtitle(array $group, int $index, int $reportYear): string
{
    $defaults = [
        'Kế hoạch năm ' . $reportYear . ' đã giao',
        'Điều chỉnh tăng trưởng',
        'Chỉ tiêu kế hoạch năm ' . $reportYear,
    ];

    $value = trim((string) ($group['subtitles'][$index] ?? ''));
    if ($value !== '' && !preg_match('/^0(?:\.0+)?$/', $value)) {
        return $value;
    }

    return $defaults[$index] ?? '';
}

$flash = $_GET['status'] ?? '';
$view = $_GET['view'] ?? '';
$error = '';
$state = null;
$zipArchiveAvailable = class_exists('ZipArchive');

try {
    if ($_SERVER['REQUEST_METHOD'] === 'POST') {
        $action = (string) ($_POST['action'] ?? '');

        if ($action === 'import') {
            khnv_import_uploaded_workbook($_FILES['import_file'] ?? []);
            header('Location: ' . $selfUrl . '?status=imported');
            exit;
        }

        $payload = khnv_collect_payload();

        if ($action === 'save') {
            khnv_save_workbook(KHNV_INPUT_XLSX, $payload);
            header('Location: ' . $selfUrl . '?status=saved');
            exit;
        }

        if ($action === 'export') {
            $state = khnv_parse_workbook(KHNV_INPUT_XLSX);
            khnv_export_docx_download($state, KHNV_TEMPLATE_DOCX, 'Xuat_chi_tieu.docx');
            exit;
        }
    }

    if ($state === null) {
        $state = khnv_parse_workbook(KHNV_INPUT_XLSX);
    }
} catch (Throwable $e) {
    $error = $e->getMessage();
    if ($state === null) {
        if ($zipArchiveAvailable) {
            try {
                $state = khnv_parse_workbook(KHNV_INPUT_XLSX);
            } catch (Throwable $inner) {
                $state = [
                    'rows' => [],
                    'groups' => [],
                    'mergeRanges' => [],
                ];
                if ($error === '') {
                    $error = $inner->getMessage();
                }
            }
        } else {
            $state = [
                'rows' => [],
                'groups' => [],
                'mergeRanges' => [],
            ];
        }
    }
}

$rows = $state['rows'] ?? [];
$pgdGroups = $state['groups'] ?? [];
$loanGroups = khnv_detect_loan_groups($rows);
$reportYear = khnv_detect_report_year($rows);
$dataStartRow = khnv_detect_data_start_row($rows);
$headerRowCount = max(1, $dataStartRow - 1);
$displayColumns = khnv_detect_used_columns($rows);
$headerRows = khnv_build_sheet_header_rows($rows, $displayColumns, $headerRowCount, $state['mergeRanges'] ?? []);
$groupCount = count($pgdGroups);
$rowCount = 0;
$formulaMap = [];
foreach ($rows as $rowNum => $row) {
    if ($rowNum >= $dataStartRow && khnv_row_has_content($row)) {
        $rowCount++;
    }
    foreach (($row['cells'] ?? []) as $cell) {
        if (!empty($cell['has_formula'])) {
            $formulaMap[(string) $cell['ref']] = (string) ($cell['formula'] ?? '');
        }
    }
}
$tableColspan = count($displayColumns);
$sourceFile = basename(KHNV_INPUT_XLSX);
$statusText = 'Sẵn sàng chỉnh sửa và xuất file';
if ($flash === 'saved') {
    $statusText = 'Đã lưu trực tiếp vào ' . $sourceFile;
} elseif ($flash === 'imported') {
    $statusText = 'Đã import file local và thay thế ' . $sourceFile;
}
?>
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Chỉ tiêu KHNV</title>
    <link rel="stylesheet" href="<?= khnv_h($publicPrefix . 'style.css') ?>">
</head>
<body class="<?= khnv_h($view !== '' ? 'view-' . preg_replace('/[^a-z0-9_-]/i', '', (string) $view) : '') ?>">
<div class="page-shell">
    <div class="page-topbar">
        <a class="back-home" href="<?= khnv_h($homeHref) ?>">Trang chủ KHNV</a>
        <span class="page-tag">Module Chỉ tiêu</span>
    </div>

    <header class="hero">
        <div class="hero-copy">
            <div class="eyebrow">KHNV / CHITIEU / IMPORT & EXPORT</div>
            <h1>Nhập và xuất chỉ tiêu tín dụng</h1>
            <p>Đọc dữ liệu từ <strong><?= khnv_h($sourceFile) ?></strong>, nhập file local theo chuẩn <strong>CTKHNV*.xlsx</strong>, hiển thị trực tiếp theo cấu trúc workbook đang mở và xuất theo mẫu <strong>OUTPUT/<?= khnv_h(basename(KHNV_TEMPLATE_DOCX)) ?></strong>.</p>
        </div>
        <div class="hero-stats">
            <div class="stat-card">
                <span class="stat-label">Số PGD</span>
                <strong><?= (int) $groupCount ?></strong>
            </div>
            <div class="stat-card">
                <span class="stat-label">Số dòng dữ liệu</span>
                <strong><?= (int) $rowCount ?></strong>
            </div>
            <div class="stat-card">
                <span class="stat-label">Nguồn Excel</span>
                <strong><?= khnv_h($sourceFile) ?></strong>
            </div>
        </div>
    </header>

    <section class="toolbar" id="actionToolbar">
        <div class="toolbar-info">
            <div class="status-line">
                <span class="dot"></span>
                <span><?= khnv_h($statusText) ?></span>
            </div>
            <div class="hint">Bảng đang bám theo số dòng, số cột và ô gộp trong workbook hiện tại.</div>
        </div>
        <div class="toolbar-actions">
            <form id="importForm" method="post" enctype="multipart/form-data" class="import-form">
                <input
                    type="file"
                    id="importFile"
                    name="import_file"
                    class="import-input"
                    accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    required
                >
                <button type="submit" name="action" value="import" class="btn btn-secondary">Nhập chỉ tiêu từ Excel</button>
            </form>
            <button type="button" class="btn btn-ghost" id="reloadBtn">Đọc lại mẫu</button>
            <button type="submit" form="sheetForm" name="action" value="save" class="btn btn-primary">Lưu cập nhật</button>
            <button type="submit" form="sheetForm" name="action" value="export" class="btn btn-accent" id="exportBtn">Xuất chỉ tiêu tín dụng</button>
        </div>
    </section>

    <?php if ($flash === 'saved'): ?>
        <section class="notice notice-success">
            Đã lưu cập nhật vào <strong><?= khnv_h($sourceFile) ?></strong>.
        </section>
    <?php elseif ($flash === 'imported'): ?>
        <section class="notice notice-success">
            Import thành công. File local đã thay thế <strong><?= khnv_h($sourceFile) ?></strong> và bạn có thể bấm <strong>Đọc lại mẫu</strong> để nạp dữ liệu mới.
        </section>
    <?php endif; ?>

    <?php if ($error !== ''): ?>
        <section class="notice notice-error">
            <?= khnv_h($error) ?>
        </section>
    <?php endif; ?>

    <form id="sheetForm" method="post" class="sheet-form" autocomplete="off">
        <textarea id="payload" name="payload" hidden></textarea>
        <div class="table-wrap">
            <table class="sheet-table">
                <colgroup>
                    <?php foreach ($displayColumns as $columnIndex => $column): ?>
                        <?php
                        $colClass = 'col-num';
                        if ($columnIndex === 0) {
                            $colClass = 'col-stt';
                        } elseif ($columnIndex === 1) {
                            $colClass = 'col-pgd';
                        } elseif ($columnIndex === 2) {
                            $colClass = 'col-xa';
                        }
                        ?>
                        <col class="<?= khnv_h($colClass) ?>">
                    <?php endforeach; ?>
                </colgroup>
                <thead>
                    <?php if ($loanGroups): ?>
                        <tr>
                            <th rowspan="2" class="freeze-col freeze-col-1">STT</th>
                            <th rowspan="2" class="freeze-col freeze-col-2">PGD</th>
                            <th rowspan="2" class="freeze-col freeze-col-3">Xã</th>
                            <?php foreach ($loanGroups as $loanGroup): ?>
                                <th colspan="3" class="group-head"><?= khnv_h((string) ($loanGroup['label'] ?? '')) ?></th>
                            <?php endforeach; ?>
                        </tr>
                        <tr>
                            <?php foreach ($loanGroups as $loanGroup): ?>
                                <th><?= khnv_h(khnv_group_subtitle($loanGroup, 0, $reportYear)) ?></th>
                                <th><?= khnv_h(khnv_group_subtitle($loanGroup, 1, $reportYear)) ?></th>
                                <th><?= khnv_h(khnv_group_subtitle($loanGroup, 2, $reportYear)) ?></th>
                            <?php endforeach; ?>
                        </tr>
                    <?php else: ?>
                        <?php foreach ($headerRows as $headerRow): ?>
                            <tr>
                                <?php foreach (($headerRow['cells'] ?? []) as $headerCell): ?>
                                    <?php
                                    $freezeClass = '';
                                    $index = (int) ($headerCell['index'] ?? 0);
                                    if ($index === 1) {
                                        $freezeClass = ' freeze-col freeze-col-1';
                                    } elseif ($index === 2) {
                                        $freezeClass = ' freeze-col freeze-col-2';
                                    } elseif ($index === 3) {
                                        $freezeClass = ' freeze-col freeze-col-3';
                                    }
                                    $groupHeadClass = ((int) ($headerCell['colspan'] ?? 1) > 1) ? ' group-head' : '';
                                    ?>
                                    <th
                                        class="<?= khnv_h(trim($freezeClass . $groupHeadClass)) ?>"
                                        colspan="<?= (int) ($headerCell['colspan'] ?? 1) ?>"
                                        rowspan="<?= (int) ($headerCell['rowspan'] ?? 1) ?>"
                                    >
                                        <?= khnv_h((string) ($headerCell['value'] ?? '')) ?>
                                    </th>
                                <?php endforeach; ?>
                            </tr>
                        <?php endforeach; ?>
                    <?php endif; ?>
                </thead>
                <tbody>
                    <?php
                    ksort($rows);
                    foreach ($rows as $rowNum => $row):
                        if ($rowNum < $dataStartRow || !khnv_row_has_content($row)) {
                            continue;
                        }
                        ?>
                        <tr class="data-row" data-sheet-row="<?= (int) $rowNum ?>">
                            <?php foreach ($displayColumns as $columnIndex => $col): ?>
                                <?php
                                $cell = $row['cells'][$col] ?? null;
                                $value = khnv_input_value($cell ?? ['value' => '']);
                                $isFormula = !empty($cell['has_formula']);
                                $isReadonly = $cell === null || $isFormula || $col === 'A';
                                $isText = $cell === null || !empty($cell['is_text']);
                                $classes = ['cell', $isText ? 'cell-text' : 'cell-num'];
                                if ($isFormula) {
                                    $classes[] = 'cell-formula';
                                }
                                if ($isReadonly) {
                                    $classes[] = 'cell-readonly';
                                }
                                $freezeClass = '';
                                if ($columnIndex === 0) {
                                    $freezeClass = 'freeze-col freeze-col-1';
                                } elseif ($columnIndex === 1) {
                                    $freezeClass = 'freeze-col freeze-col-2';
                                } elseif ($columnIndex === 2) {
                                    $freezeClass = 'freeze-col freeze-col-3';
                                }
                                ?>
                                <td class="<?= khnv_h($freezeClass) ?>">
                                    <input
                                        class="<?= khnv_h(implode(' ', $classes)) ?>"
                                        type="text"
                                        <?= $isText ? '' : 'inputmode="decimal"' ?>
                                        <?= $cell !== null ? 'data-ref="' . khnv_h($col . $rowNum) . '"' : '' ?>
                                        data-col="<?= khnv_h($col) ?>"
                                        data-initial="<?= khnv_h($value) ?>"
                                        value="<?= khnv_h($value) ?>"
                                        <?= $isReadonly ? 'readonly' : '' ?>
                                    >
                                </td>
                            <?php endforeach; ?>
                        </tr>
                    <?php endforeach; ?>
                </tbody>
            </table>
        </div>
    </form>
</div>

<script>
(() => {
    const form = document.getElementById('sheetForm');
    const payload = document.getElementById('payload');
    const reloadBtn = document.getElementById('reloadBtn');
    const importForm = document.getElementById('importForm');
    const importFile = document.getElementById('importFile');
    const exportBtn = document.getElementById('exportBtn');
    const initialView = <?= json_encode($view, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES) ?>;
    const formulaCells = <?= json_encode($formulaMap, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES) ?>;

    function cleanNumber(value) {
        const raw = String(value || '').trim();
        if (!raw) return '';
        const normalized = raw.replace(/,/g, '');
        const num = Number(normalized);
        if (!Number.isFinite(num)) return raw;
        if (Math.abs(num) < 0.00000001 && num !== 0) {
            return '0';
        }
        const rounded = Math.round(num * 100) / 100;
        const str = rounded.toFixed(2);
        return str.replace(/\.?0+$/, '').replace(/\.$/, '');
    }

    function formatNumber(value) {
        const num = Number(value);
        if (!Number.isFinite(num)) return '';
        return cleanNumber(num.toString());
    }

    function formatVND(value) {
        const num = Number(value);
        if (!Number.isFinite(num)) return '';
        const rounded = Math.round(num * 100) / 100;
        const parts = rounded.toFixed(2).split('.');
        const intPart = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
        const decimalPart = parts[1];
        if (decimalPart === '00') {
            return intPart;
        }
        return `${intPart}.${decimalPart}`;
    }

    function colToIndex(col) {
        return String(col || '').toUpperCase().split('').reduce((total, ch) => (total * 26) + (ch.charCodeAt(0) - 64), 0);
    }

    function indexToCol(index) {
        let current = Number(index);
        let col = '';
        while (current > 0) {
            current -= 1;
            col = String.fromCharCode((current % 26) + 65) + col;
            current = Math.floor(current / 26);
        }
        return col;
    }

    function parseRef(ref) {
        const match = String(ref || '').match(/^([A-Z]+)(\d+)$/i);
        if (!match) return null;
        return {
            col: match[1].toUpperCase(),
            row: Number(match[2]),
        };
    }

    function getInputByRef(ref) {
        return form.querySelector(`[data-ref="${ref}"]`);
    }

    function getNumericValueByRef(ref) {
        const input = getInputByRef(ref);
        if (!(input instanceof HTMLInputElement)) return 0;
        const raw = input.dataset.cleanValue || cleanNumber(input.value);
        const num = Number(raw);
        return Number.isFinite(num) ? num : 0;
    }

    function evaluateFormula(formula) {
        const raw = String(formula || '').replace(/\s+/g, '').replace(/^=/, '');
        if (!raw) return null;

        const sumMatch = raw.match(/^SUM\(([A-Z]+\d+):([A-Z]+\d+)\)$/i);
        if (sumMatch) {
            const start = parseRef(sumMatch[1]);
            const end = parseRef(sumMatch[2]);
            if (!start || !end) return null;
            let total = 0;
            const startRow = Math.min(start.row, end.row);
            const endRow = Math.max(start.row, end.row);
            const startCol = Math.min(colToIndex(start.col), colToIndex(end.col));
            const endCol = Math.max(colToIndex(start.col), colToIndex(end.col));
            for (let row = startRow; row <= endRow; row += 1) {
                for (let colIndex = startCol; colIndex <= endCol; colIndex += 1) {
                    total += getNumericValueByRef(`${indexToCol(colIndex)}${row}`);
                }
            }
            return total;
        }

        const binaryMatch = raw.match(/^([A-Z]+\d+)([+\-])([A-Z]+\d+)$/i);
        if (binaryMatch) {
            const left = getNumericValueByRef(binaryMatch[1].toUpperCase());
            const right = getNumericValueByRef(binaryMatch[3].toUpperCase());
            return binaryMatch[2] === '-' ? (left - right) : (left + right);
        }

        return null;
    }

    function syncCleanValue(input) {
        if (!input.classList.contains('cell-num')) return;
        const cleanVal = cleanNumber(input.value);
        input.dataset.cleanValue = cleanVal;
    }

    function formatInputDisplay(input) {
        if (!input.classList.contains('cell-num')) return;
        const cleanVal = input.dataset.cleanValue || cleanNumber(input.value);
        input.dataset.cleanValue = cleanVal;
        if (cleanVal === '') {
            input.value = '';
            return;
        }
        const display = formatVND(cleanVal);
        input.dataset.display = display;
        input.value = input.hasAttribute('readonly') ? display : display;
    }

    function showCleanValue(input) {
        if (!input.classList.contains('cell-num')) return;
        if (input.hasAttribute('readonly')) return;
        const cleanVal = input.dataset.cleanValue || cleanNumber(input.value);
        input.value = cleanVal;
    }

    function recalcAllFormulas() {
        Object.entries(formulaCells).forEach(([ref, formula]) => {
            const input = getInputByRef(ref);
            if (!(input instanceof HTMLInputElement)) return;
            const computed = evaluateFormula(formula);
            if (computed === null) return;
            const cleanVal = formatNumber(computed);
            input.dataset.cleanValue = cleanVal;
            input.dataset.display = formatVND(cleanVal);
            input.value = input.dataset.display || cleanVal;
        });
    }

    function updateDirtyState() {
        const dirty = Array.from(form.querySelectorAll('[data-ref]')).some((input) => {
            if (!(input instanceof HTMLInputElement) || input.hasAttribute('readonly')) {
                return false;
            }
            const initial = input.dataset.initial ?? '';
            const current = input.classList.contains('cell-num')
                ? cleanNumber(input.value)
                : input.value.trim();
            return current !== initial;
        });
        document.body.classList.toggle('dirty', dirty);
    }

    form.addEventListener('input', (event) => {
        const target = event.target;
        if (!(target instanceof HTMLInputElement)) return;
        syncCleanValue(target);
        recalcAllFormulas();
        updateDirtyState();
    });

    form.addEventListener('focus', (event) => {
        const target = event.target;
        if (!(target instanceof HTMLInputElement)) return;
        showCleanValue(target);
    }, true);

    form.addEventListener('blur', (event) => {
        const target = event.target;
        if (!(target instanceof HTMLInputElement)) return;
        syncCleanValue(target);
        recalcAllFormulas();
        formatInputDisplay(target);
    }, true);

    form.addEventListener('submit', (event) => {
        const submitter = event.submitter;
        const action = submitter?.value || 'save';
        if (action === 'export' || action === 'save') {
            const changes = {};
            form.querySelectorAll('[data-ref]').forEach((input) => {
                if (!(input instanceof HTMLInputElement) || input.hasAttribute('readonly')) {
                    return;
                }
                const initial = input.dataset.initial ?? '';
                const current = input.classList.contains('cell-num')
                    ? cleanNumber(input.value)
                    : input.value.trim();
                if (current !== initial) {
                    changes[input.dataset.ref] = current;
                }
            });
            payload.value = JSON.stringify(changes);
        }
    });

    reloadBtn.addEventListener('click', () => {
        window.location.href = window.location.pathname;
    });

    document.querySelectorAll('input.cell-num').forEach((input) => {
        if (!(input instanceof HTMLInputElement)) return;
        syncCleanValue(input);
    });
    recalcAllFormulas();
    document.querySelectorAll('input.cell-num').forEach((input) => {
        if (!(input instanceof HTMLInputElement)) return;
        formatInputDisplay(input);
    });
    updateDirtyState();

    function highlightTarget(target) {
        if (!(target instanceof HTMLElement)) return;
        target.classList.add('pulse-focus');
        target.scrollIntoView({ behavior: 'smooth', block: 'center' });
        window.setTimeout(() => {
            target.classList.remove('pulse-focus');
        }, 2400);
    }

    if (initialView === 'import' && importForm instanceof HTMLElement) {
        highlightTarget(importForm);
        window.setTimeout(() => importFile?.focus(), 350);
    }

    if (initialView === 'export' && exportBtn instanceof HTMLElement) {
        highlightTarget(exportBtn);
        window.setTimeout(() => exportBtn.focus(), 350);
    }
})();
</script>
</body>
</html>
