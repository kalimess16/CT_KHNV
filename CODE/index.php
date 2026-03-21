<?php
declare(strict_types=1);

require __DIR__ . '/data.php';
khnv_redirect_localhost_to_ip();

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
$error = '';
$state = null;
$zipArchiveAvailable = class_exists('ZipArchive');

try {
    if ($_SERVER['REQUEST_METHOD'] === 'POST') {
        $action = (string) ($_POST['action'] ?? '');

        if ($action === 'import') {
            khnv_import_uploaded_workbook($_FILES['import_file'] ?? []);
            header('Location: index.php?status=imported');
            exit;
        }

        $payload = khnv_collect_payload();

        if ($action === 'save') {
            khnv_save_workbook(KHNV_INPUT_XLSX, $payload);
            header('Location: index.php?status=saved');
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
                ];
                if ($error === '') {
                    $error = $inner->getMessage();
                }
            }
        } else {
            $state = [
                'rows' => [],
                'groups' => [],
            ];
        }
    }
}

$rows = $state['rows'] ?? [];
$pgdGroups = $state['groups'] ?? [];
$loanGroups = khnv_detect_loan_groups($rows);
$reportYear = khnv_detect_report_year($rows);
$maxRow = $rows ? max(array_keys($rows)) : 0;
$groupCount = count($pgdGroups);
$rowCount = max(0, $maxRow - 2);
$sheetColumns = [];
foreach ($loanGroups as $loanGroup) {
    $sheetColumns[] = (string) ($loanGroup['start'] ?? '');
    $sheetColumns[] = (string) ($loanGroup['adjust'] ?? '');
    $sheetColumns[] = (string) ($loanGroup['target'] ?? '');
}
$tableColspan = 3 + count($sheetColumns);
$groupStarts = [];
foreach ($pgdGroups as $idx => $group) {
    $groupStarts[(int) ($group['start'] ?? 0)] = $idx;
}
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
    <title>Xuất chỉ tiêu KHNV</title>
    <link rel="stylesheet" href="style.css">
</head>
<body>
<div class="page-shell">
    <header class="hero">
        <div class="hero-copy">
            <div class="eyebrow">KHNV / INPUT / OUTPUT</div>
            <h1>Trang xuất chỉ tiêu tín dụng</h1>
            <p>Đọc dữ liệu từ <strong><?= khnv_h($sourceFile) ?></strong>, import file local theo chuẩn <strong>CTKHNV*.xlsx</strong>, chỉnh trực tiếp trên màn hình và xuất theo mẫu <strong>OUTPUT/<?= khnv_h(basename(KHNV_TEMPLATE_DOCX)) ?></strong>.</p>
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

    <section class="toolbar">
        <div class="toolbar-info">
            <div class="status-line">
                <span class="dot"></span>
                <span><?= khnv_h($statusText) ?></span>
            </div>
            <div class="hint">Các ô công thức sẽ tự tính lại khi bạn thay đổi cột nhập liệu.</div>
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
                <button type="submit" name="action" value="import" class="btn btn-secondary">Import file local</button>
            </form>
            <button type="button" class="btn btn-ghost" id="reloadBtn">Đọc lại mẫu</button>
            <button type="submit" form="sheetForm" name="action" value="save" class="btn btn-primary">Lưu cập nhật</button>
            <button type="submit" form="sheetForm" name="action" value="export" class="btn btn-accent">Xuất chỉ tiêu</button>
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
                    <col class="col-stt">
                    <col class="col-pgd">
                    <col class="col-xa">
                    <?php foreach ($sheetColumns as $_sheetColumn): ?>
                        <col class="col-num">
                    <?php endforeach; ?>
                </colgroup>
                <thead>
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
                </thead>
                <tbody>
                    <?php
                    ksort($rows);
                    foreach ($rows as $rowNum => $row):
                        if ($rowNum < 3 || !khnv_row_has_content($row)) {
                            continue;
                        }
                        if (isset($groupStarts[$rowNum])) {
                            $group = $pgdGroups[$groupStarts[$rowNum]];
                            ?>
                            <tr class="divider-row">
                                <td colspan="<?= (int) $tableColspan ?>">
                                    <span class="divider-label"><?= khnv_h((string) ($group['pgd'] ?? '')) ?></span>
                                    <span class="divider-meta"><?= count((array) ($group['rows'] ?? [])) ?> đơn vị</span>
                                </td>
                            </tr>
                            <?php
                        }
                        $stt = khnv_input_value($row['cells']['A'] ?? ['value' => '']);
                        $pgd = khnv_input_value($row['cells']['B'] ?? ['value' => '']);
                        $xa = khnv_input_value($row['cells']['C'] ?? ['value' => '']);
                        ?>
                        <tr class="data-row" data-sheet-row="<?= (int) $rowNum ?>">
                            <td class="freeze-col freeze-col-1">
                                <input class="cell cell-readonly" type="text" value="<?= khnv_h($stt) ?>" readonly>
                            </td>
                            <td class="freeze-col freeze-col-2">
                                <input class="cell cell-text" type="text" data-ref="B<?= (int) $rowNum ?>" data-col="B" data-initial="<?= khnv_h($pgd) ?>" value="<?= khnv_h($pgd) ?>">
                            </td>
                            <td class="freeze-col freeze-col-3">
                                <input class="cell cell-text" type="text" data-ref="C<?= (int) $rowNum ?>" data-col="C" data-initial="<?= khnv_h($xa) ?>" value="<?= khnv_h($xa) ?>">
                            </td>
                            <?php foreach ($sheetColumns as $col): ?>
                                <?php
                                $cell = $row['cells'][$col] ?? ['value' => '', 'has_formula' => false];
                                $value = khnv_input_value($cell);
                                $readonly = !empty($cell['has_formula']) ? 'readonly' : '';
                                $classes = ['cell', 'cell-num'];
                                if (!empty($cell['has_formula'])) {
                                    $classes[] = 'cell-formula';
                                }
                                ?>
                                <td>
                                    <input
                                        class="<?= khnv_h(implode(' ', $classes)) ?>"
                                        type="text"
                                        inputmode="decimal"
                                        data-ref="<?= khnv_h($col . $rowNum) ?>"
                                        data-col="<?= khnv_h($col) ?>"
                                        data-initial="<?= khnv_h($value) ?>"
                                        value="<?= khnv_h($value) ?>"
                                        <?= $readonly ?>
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
    const groups = <?= json_encode(
        array_map(
            static fn(array $group): array => [
                'base' => (string) ($group['start'] ?? ''),
                'delta' => (string) ($group['adjust'] ?? ''),
                'output' => (string) ($group['target'] ?? ''),
            ],
            $loanGroups
        ),
        JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES
    ) ?>;

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

    function parseNumber(value) {
        const raw = cleanNumber(value);
        if (raw === '') return 0;
        const num = Number(raw);
        return Number.isFinite(num) ? num : 0;
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

    function getInput(row, col) {
        return row.querySelector(`[data-col="${col}"]`);
    }

    function recalcRow(row) {
        for (const group of groups) {
            const base = parseNumber(getInput(row, group.base)?.value);
            const delta = parseNumber(getInput(row, group.delta)?.value);
            const output = getInput(row, group.output);
            if (output && output.hasAttribute('readonly')) {
                const calculated = formatNumber(base + delta);
                output.value = calculated;
                output.dataset.cleanValue = calculated;
                output.setAttribute('data-display', formatVND(calculated));
            }
        }
    }

    function formatInputDisplay(input) {
        if (!input.classList.contains('cell-num')) return;
        if (input.hasAttribute('readonly')) {
            input.value = input.getAttribute('data-display') || input.value;
            return;
        }
        const cleanVal = cleanNumber(input.value);
        input.dataset.cleanValue = cleanVal;
        input.value = formatVND(cleanVal);
    }

    function showCleanValue(input) {
        if (!input.classList.contains('cell-num')) return;
        if (input.hasAttribute('readonly')) return;
        const cleanVal = input.dataset.cleanValue || cleanNumber(input.value);
        input.value = cleanVal;
    }

    function syncCleanValue(input) {
        if (!input.classList.contains('cell-num')) return;
        if (input.hasAttribute('readonly')) return;
        const cleanVal = cleanNumber(input.value);
        input.dataset.cleanValue = cleanVal;
    }

    function updateDirtyState() {
        const dirty = Array.from(form.querySelectorAll('[data-ref]')).some((input) => {
            if (input.hasAttribute('readonly')) return false;
            return input.value !== (input.dataset.initial ?? '');
        });
        document.body.classList.toggle('dirty', dirty);
    }

    form.addEventListener('input', (event) => {
        const target = event.target;
        if (!(target instanceof HTMLInputElement)) return;
        syncCleanValue(target);
        const row = target.closest('tr[data-sheet-row]');
        if (row) {
            recalcRow(row);
        }
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
        formatInputDisplay(target);
    }, true);

    form.addEventListener('submit', (event) => {
        const submitter = event.submitter;
        const action = submitter?.value || 'save';
        if (action === 'export' || action === 'save') {
            const changes = {};
            form.querySelectorAll('[data-ref]').forEach((input) => {
                if (input.hasAttribute('readonly')) return;
                const initial = input.dataset.initial ?? '';
                const current = cleanNumber(input.value);
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

    document.querySelectorAll('tr[data-sheet-row]').forEach((row) => {
        recalcRow(row);
        row.querySelectorAll('[data-col]').forEach((input) => {
            if (input.classList.contains('cell-num')) {
                syncCleanValue(input);
                formatInputDisplay(input);
            }
        });
    });
    updateDirtyState();
})();
</script>
</body>
</html>
