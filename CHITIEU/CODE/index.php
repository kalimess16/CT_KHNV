<?php
declare(strict_types=1);

require_once dirname(__DIR__, 2) . '/access_control.php';
require __DIR__ . '/data.php';

khnv_access_redirect_localhost_to_ip();
khnv_access_enforce_client_ip();

$publicPrefix = defined('KHNV_PUBLIC_PREFIX') ? KHNV_PUBLIC_PREFIX : '';
$selfBaseUrl = defined('KHNV_SELF_URL') ? KHNV_SELF_URL : 'index.php';
$embedded = (string) ($_GET['embedded'] ?? '') === '1';
$view = (string) ($_GET['view'] ?? '');
$activePanel = '';
$dataset = khnv_normalize_workbook_key($_POST['dataset'] ?? $_GET['dataset'] ?? KHNV_DEFAULT_WORKBOOK_KEY);
$exportMode = khnv_normalize_export_mode($_POST['export_mode'] ?? $_GET['export_mode'] ?? KHNV_DEFAULT_EXPORT_MODE);

function khnv_h(string $value): string
{
    return htmlspecialchars($value, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');
}

function khnv_url_with_query(string $baseUrl, array $params): string
{
    $filtered = [];
    foreach ($params as $key => $value) {
        if ($value === null || $value === '') {
            continue;
        }
        $filtered[$key] = (string) $value;
    }

    if ($filtered === []) {
        return $baseUrl;
    }

    return $baseUrl . (str_contains($baseUrl, '?') ? '&' : '?') . http_build_query($filtered);
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

$workbookConfigs = khnv_workbook_configs();
$exportModeConfigs = khnv_export_mode_configs();
$currentWorkbook = khnv_get_workbook_config($dataset);
$currentWorkbookPath = (string) ($currentWorkbook['path'] ?? '');
$flash = (string) ($_GET['status'] ?? '');
$importedKey = khnv_normalize_workbook_key($_GET['imported'] ?? $dataset);
$zipArchiveAvailable = class_exists('ZipArchive');
$error = '';
$state = null;

$baseParams = [
    'embedded' => $embedded ? '1' : null,
];

try {
    if ($_SERVER['REQUEST_METHOD'] === 'POST') {
        $action = (string) ($_POST['action'] ?? '');

        if ($action === 'import') {
            $expectedKey = khnv_normalize_workbook_key($_POST['workbook_key'] ?? $dataset);
            $importedKey = khnv_import_uploaded_workbook($_FILES['import_file'] ?? [], $expectedKey);
            header('Location: ' . khnv_url_with_query($selfBaseUrl, array_merge($baseParams, [
                'dataset' => $importedKey,
                'status' => 'imported',
                'imported' => $importedKey,
            ])));
            exit;
        }

        if ($action === 'clear') {
            khnv_clear_workbook($currentWorkbookPath);
            header('Location: ' . khnv_url_with_query($selfBaseUrl, array_merge($baseParams, [
                'dataset' => $dataset,
                'status' => 'cleared',
            ])));
            exit;
        }

        $payload = khnv_collect_payload();

        if ($action === 'save') {
            khnv_save_workbook($currentWorkbookPath, $payload);
            header('Location: ' . khnv_url_with_query($selfBaseUrl, array_merge($baseParams, [
                'dataset' => $dataset,
                'status' => 'saved',
            ])));
            exit;
        }

        if ($action === 'export') {
            $states = khnv_parse_all_workbooks();
            if ($payload !== []) {
                khnv_apply_changes($states[$dataset], $payload);
            }
            khnv_export_mode_download($states, (string) ($_POST['export_mode'] ?? KHNV_DEFAULT_EXPORT_MODE));
            exit;
        }
    }

    if ($state === null) {
        $state = khnv_parse_workbook($currentWorkbookPath);
    }
} catch (Throwable $e) {
    $error = $e->getMessage();
    if ($state === null) {
        if ($zipArchiveAvailable) {
            try {
                $state = khnv_parse_workbook($currentWorkbookPath);
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
$loanGroups = khnv_detect_loan_groups($rows);
$reportYear = khnv_detect_report_year($rows);
$dataStartRow = khnv_detect_data_start_row($rows);
$displayColumns = khnv_detect_used_columns($rows);
$headerRowCount = max(1, $dataStartRow - 1);
$headerRows = khnv_build_sheet_header_rows($rows, $displayColumns, $headerRowCount, $state['mergeRanges'] ?? []);
$formulaMap = [];

foreach ($rows as $row) {
    foreach (($row['cells'] ?? []) as $cell) {
        if (!empty($cell['has_formula'])) {
            $formulaMap[(string) ($cell['ref'] ?? '')] = (string) ($cell['formula'] ?? '');
        }
    }
}

$sourceFile = basename($currentWorkbookPath);
$currentWorkbookLabel = khnv_get_workbook_label($dataset);
$currentWorkbookTitle = khnv_get_workbook_title($dataset);
$statusText = 'Đang xem ' . $currentWorkbookLabel . ' từ ' . $sourceFile . '.';
if ($flash === 'saved') {
    $statusText = 'Đã lưu cập nhật trực tiếp vào ' . $sourceFile . '.';
} elseif ($flash === 'imported') {
    $importedFile = basename(khnv_get_workbook_path($importedKey));
    $statusText = 'Đã import file local và thay thế ' . $importedFile . '.';
} elseif ($flash === 'cleared') {
    $statusText = 'Đã xóa toàn bộ số liệu trong ' . $sourceFile . '.';
}
?>
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Chỉ tiêu KHNV</title>
    <link rel="icon" type="image/png" href="<?= khnv_h($publicPrefix !== '' ? '../iconweb.png' : '../../iconweb.png') ?>">
    <link rel="stylesheet" href="<?= khnv_h($publicPrefix . 'style.css') ?>">
</head>
<?php
$bodyClasses = [];
if ($view !== '') {
    $bodyClasses[] = 'view-' . preg_replace('/[^a-z0-9_-]/i', '', $view);
}
if ($embedded) {
    $bodyClasses[] = 'embedded';
}
$bodyClasses[] = 'dataset-' . preg_replace('/[^a-z0-9_-]/i', '', $dataset);
?>
<body class="<?= khnv_h(implode(' ', $bodyClasses)) ?>">
<div class="page-shell">
    <section class="toolbar" id="actionToolbar">
        <?php
        $importLauncherUrl = khnv_url_with_query($selfBaseUrl, array_merge($baseParams, [
            'dataset' => $dataset,
        ]));
        $exportLauncherUrl = khnv_url_with_query($selfBaseUrl, array_merge($baseParams, [
            'dataset' => $dataset,
        ]));
        ?>

        <div class="toolbar-top">
            <div class="toolbar-info">
                <div class="embedded-note">Nhập file local theo chuẩn <strong>CTKHNV_DP/TW*.xlsx</strong>.</div>
                <div class="status-line">
                    <span class="dot"></span>
                    <span><?= khnv_h($statusText) ?></span>
                </div>
            </div>

            <div class="action-launchers" id="actionLaunchers">
                <a
                    class="panel-launcher panel-launcher-import<?= $activePanel === 'import' ? ' active' : '' ?>"
                    href="<?= khnv_h($importLauncherUrl) ?>"
                    data-panel-target="import"
                    aria-controls="importPanel"
                    aria-expanded="<?= $activePanel === 'import' ? 'true' : 'false' ?>"
                    role="button"
                >
                    <span class="panel-launcher-kicker">Import Excel</span>
                </a>

                <a
                    class="panel-launcher panel-launcher-export<?= $activePanel === 'export' ? ' active' : '' ?>"
                    href="<?= khnv_h($exportLauncherUrl) ?>"
                    data-panel-target="export"
                    aria-controls="exportPanel"
                    aria-expanded="<?= $activePanel === 'export' ? 'true' : 'false' ?>"
                    role="button"
                >
                    <span class="panel-launcher-kicker">Xuất Chỉ Tiêu</span>
                </a>
            </div>

            <div class="toolbar-side">
                <div class="source-switch" id="datasetSwitch">
                    <?php foreach ($workbookConfigs as $key => $config): ?>
                        <?php
                        $switchUrl = khnv_url_with_query($selfBaseUrl, array_merge($baseParams, ['dataset' => $key]));
                        $switchClasses = ['source-link'];
                        if ($dataset === $key) {
                            $switchClasses[] = 'active';
                        }
                        ?>
                        <a class="<?= khnv_h(implode(' ', $switchClasses)) ?>" href="<?= khnv_h($switchUrl) ?>">
                            <span><?= khnv_h((string) ($config['label'] ?? '')) ?></span>
                            <small><?= khnv_h(basename((string) ($config['path'] ?? ''))) ?></small>
                        </a>
                    <?php endforeach; ?>
                </div>

                <div class="toolbar-actions">
                    <button type="submit" form="sheetForm" name="action" value="clear" class="btn btn-danger" id="clearBtn" data-workbook-label="<?= khnv_h($currentWorkbookLabel) ?>">Xóa số liệu</button>
                    <button type="button" class="btn btn-ghost" id="reloadBtn">Đọc lại mẫu</button>
                    <button type="submit" form="sheetForm" name="action" value="save" class="btn btn-primary">Lưu cập nhật</button>
                </div>
            </div>
        </div>

        <div class="toolbar-panels<?= $activePanel !== '' ? ' has-active-panel' : '' ?>" id="toolbarPanels">
            <form
                method="post"
                enctype="multipart/form-data"
                class="action-card action-panel import-card<?= $activePanel === 'import' ? ' is-active' : '' ?>"
                id="importPanel"
                data-panel="import"
                <?= $activePanel === 'import' ? '' : 'hidden' ?>
            >
                <div class="action-card-copy">
                    <span class="card-kicker">Import Excel</span>
                    <strong>Nhập lại workbook nguồn</strong>
                    <span class="card-meta">Sau khi mở panel, chọn đúng nguồn `TW/DP`, chọn file local rồi nhập vào workbook cần thay.</span>
                </div>
                <div class="action-card-controls">
                    <div class="choice-group">
                        <?php foreach ($workbookConfigs as $key => $config): ?>
                            <?php $checked = $dataset === $key; ?>
                            <label class="choice-pill<?= $checked ? ' active' : '' ?>">
                                <input type="radio" name="workbook_key" value="<?= khnv_h($key) ?>" <?= $checked ? 'checked' : '' ?> <?= $activePanel === 'import' ? '' : 'disabled' ?>>
                                <span><?= khnv_h((string) ($config['label'] ?? '')) ?></span>
                                <small><?= khnv_h((string) ($config['title'] ?? '')) ?></small>
                            </label>
                        <?php endforeach; ?>
                    </div>
                    <input
                        type="file"
                        name="import_file"
                        class="import-input"
                        accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        <?= $activePanel === 'import' ? '' : 'disabled' ?>
                        required
                    >
                    <button type="submit" name="action" value="import" class="btn btn-secondary" id="importBtn" <?= $activePanel === 'import' ? '' : 'disabled' ?>>Nhập file Excel</button>
                </div>
            </form>

            <div
                class="action-card action-panel export-card<?= $activePanel === 'export' ? ' is-active' : '' ?>"
                id="exportPanel"
                data-panel="export"
                <?= $activePanel === 'export' ? '' : 'hidden' ?>
            >
                <div class="action-card-copy">
                    <span class="card-kicker">Xuất DOCX</span>
                    <strong>Chọn mẫu và tải xuống</strong>
                    <span class="card-meta">Sau khi mở panel, chọn loại mẫu cần lấy rồi bấm xuất một lần.</span>
                </div>
                <div class="action-card-controls">
                    <div class="choice-group export-options">
                        <?php foreach ($exportModeConfigs as $modeKey => $config): ?>
                            <?php $checked = $exportMode === $modeKey; ?>
                            <label class="choice-pill export-option<?= $checked ? ' active' : '' ?>">
                                <input
                                    type="radio"
                                    name="export_mode"
                                    value="<?= khnv_h($modeKey) ?>"
                                    form="sheetForm"
                                    <?= $checked ? 'checked' : '' ?>
                                    <?= $activePanel === 'export' ? '' : 'disabled' ?>
                                >
                                <span><?= khnv_h((string) ($config['label'] ?? '')) ?></span>
                                <small><?= khnv_h((string) ($config['title'] ?? '')) ?></small>
                            </label>
                        <?php endforeach; ?>
                    </div>
                    <button type="submit" form="sheetForm" name="action" value="export" class="btn btn-accent" id="exportBtn" <?= $activePanel === 'export' ? '' : 'disabled' ?>>Xuất chỉ tiêu tín dụng</button>
                </div>
            </div>
        </div>

    </section>

    <?php if ($flash === 'saved'): ?>
        <section class="notice notice-success">
            Đã lưu cập nhật vào <strong><?= khnv_h($sourceFile) ?></strong>.
        </section>
    <?php elseif ($flash === 'imported'): ?>
        <section class="notice notice-success">
            Import thành công. File local đã thay thế <strong><?= khnv_h(basename(khnv_get_workbook_path($importedKey))) ?></strong> và bạn có thể bấm <strong>Đọc lại mẫu</strong> để nạp dữ liệu mới.
        </section>
    <?php elseif ($flash === 'cleared'): ?>
        <section class="notice notice-success">
            Đã xóa số liệu của workbook <strong><?= khnv_h($currentWorkbookLabel) ?></strong>. Các ô số liệu được đưa về trạng thái rỗng.
        </section>
    <?php endif; ?>

    <?php if ($error !== ''): ?>
        <section class="notice notice-error">
            <?= khnv_h($error) ?>
        </section>
    <?php endif; ?>

    <form id="sheetForm" method="post" class="sheet-form" autocomplete="off">
        <input type="hidden" name="dataset" value="<?= khnv_h($dataset) ?>">
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
                                        <?= $isText ? '' : 'inputmode="numeric"' ?>
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
    const clearBtn = document.getElementById('clearBtn');
    const toolbarPanels = document.getElementById('toolbarPanels');
    const panelLaunchers = Array.from(document.querySelectorAll('[data-panel-target]'));
    const importPanel = document.getElementById('importPanel');
    const exportPanel = document.getElementById('exportPanel');
    const exportBtn = document.getElementById('exportBtn');
    const initialView = <?= json_encode($activePanel, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES) ?>;
    const formulaCells = <?= json_encode($formulaMap, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES) ?>;
    const supportedPanels = new Set(['import', 'export']);
    const panels = {
        import: importPanel instanceof HTMLElement ? importPanel : null,
        export: exportPanel instanceof HTMLElement ? exportPanel : null,
    };
    let currentPanel = supportedPanels.has(initialView) ? initialView : '';

    function cleanNumber(value) {
        const raw = String(value || '').trim();
        if (!raw) return '';
        const normalized = raw.replace(/,/g, '');
        const num = Number(normalized);
        if (!Number.isFinite(num)) return raw;
        if (Math.abs(num) < 0.00000001 && num !== 0) {
            return '0';
        }
        return String(Math.round(num));
    }

    function formatNumber(value) {
        const num = Number(value);
        if (!Number.isFinite(num)) return '';
        return cleanNumber(num.toString());
    }

    function formatVND(value) {
        const num = Number(value);
        if (!Number.isFinite(num)) return '';
        return String(Math.round(num)).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
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

    function hasValueByRef(ref) {
        const input = getInputByRef(ref);
        if (!(input instanceof HTMLInputElement)) {
            return false;
        }

        const raw = input.classList.contains('cell-num')
            ? (input.dataset.cleanValue || cleanNumber(input.value))
            : input.value.trim();

        return String(raw).trim() !== '';
    }

    function stripOuterParens(expression) {
        let current = String(expression || '').trim();
        while (current.startsWith('(') && current.endsWith(')')) {
            let depth = 0;
            let isWrapper = true;
            for (let i = 0; i < current.length; i += 1) {
                const char = current[i];
                if (char === '(') {
                    depth += 1;
                } else if (char === ')') {
                    depth -= 1;
                    if (depth < 0) {
                        isWrapper = false;
                        break;
                    }
                    if (depth === 0 && i < current.length - 1) {
                        isWrapper = false;
                        break;
                    }
                }
            }
            if (!isWrapper || depth !== 0) {
                break;
            }
            current = current.slice(1, -1).trim();
        }
        return current;
    }

    function splitTopLevelArgs(expression) {
        const args = [];
        let current = '';
        let depth = 0;
        for (const char of String(expression || '')) {
            if (char === '(') {
                depth += 1;
                current += char;
                continue;
            }
            if (char === ')') {
                depth -= 1;
                if (depth < 0) return null;
                current += char;
                continue;
            }
            if ((char === ',' || char === ';') && depth === 0) {
                args.push(current.trim());
                current = '';
                continue;
            }
            current += char;
        }
        if (depth !== 0) return null;
        args.push(current.trim());
        return args;
    }

    function parseFunctionArgs(expression, name) {
        const normalized = stripOuterParens(expression);
        const prefix = `${String(name || '').toUpperCase()}(`;
        if (!normalized.toUpperCase().startsWith(prefix) || !normalized.endsWith(')')) {
            return null;
        }
        return splitTopLevelArgs(normalized.slice(prefix.length, -1));
    }

    function isZeroLiteral(expression) {
        const normalized = stripOuterParens(String(expression || '').trim());
        if (!normalized) return false;
        const num = Number(normalized);
        return Number.isFinite(num) && Math.abs(num) < 0.00000001;
    }

    function expressionKey(expression) {
        return stripOuterParens(String(expression || '').replace(/\s+/g, '').trim()).replace(/^\+/, '').toUpperCase();
    }

    function extractIfPassthroughExpression(condition, whenTrue, whenFalse) {
        const normalizedCondition = stripOuterParens(condition);
        const match = normalizedCondition.match(/^(.+?)(>=|<=|<>|=|>|<)(.+)$/);
        if (!match) return null;

        const [, leftRaw, operator, rightRaw] = match;
        if (!['>', '>=', '<', '<='].includes(operator)) {
            return null;
        }

        const left = stripOuterParens(leftRaw);
        const right = stripOuterParens(rightRaw);
        const trueKey = expressionKey(whenTrue);
        const falseKey = expressionKey(whenFalse);
        const leftKey = expressionKey(left);
        const rightKey = expressionKey(right);
        const leftIsZero = isZeroLiteral(left);
        const rightIsZero = isZeroLiteral(right);

        if (isZeroLiteral(whenFalse)) {
            if (trueKey && trueKey === leftKey && rightIsZero && ['>', '>='].includes(operator)) {
                return whenTrue;
            }
            if (trueKey && trueKey === rightKey && leftIsZero && ['<', '<='].includes(operator)) {
                return whenTrue;
            }
        }

        if (isZeroLiteral(whenTrue)) {
            if (falseKey && falseKey === leftKey && rightIsZero && ['<', '<='].includes(operator)) {
                return whenFalse;
            }
            if (falseKey && falseKey === rightKey && leftIsZero && ['>', '>='].includes(operator)) {
                return whenFalse;
            }
        }

        return null;
    }

    function extractPassthroughExpression(formula) {
        const normalized = stripOuterParens(formula);

        const maxArgs = parseFunctionArgs(normalized, 'MAX');
        if (Array.isArray(maxArgs) && maxArgs.length === 2) {
            if (isZeroLiteral(maxArgs[0])) return maxArgs[1];
            if (isZeroLiteral(maxArgs[1])) return maxArgs[0];
        }

        const ifArgs = parseFunctionArgs(normalized, 'IF');
        if (Array.isArray(ifArgs) && ifArgs.length === 3) {
            return extractIfPassthroughExpression(ifArgs[0], ifArgs[1], ifArgs[2]);
        }

        return null;
    }

    function unwrapPassthroughExpression(formula) {
        let current = stripOuterParens(formula);
        while (true) {
            const next = extractPassthroughExpression(current);
            if (next === null) {
                return current;
            }
            const normalizedNext = stripOuterParens(next);
            if (expressionKey(normalizedNext) === expressionKey(current)) {
                return current;
            }
            current = normalizedNext;
        }
    }

    function getFormulaTermValue(term) {
        const normalized = String(term || '').trim().toUpperCase();
        if (!normalized) return null;

        const sumMatch = normalized.match(/^SUM\(([A-Z]+\d+):([A-Z]+\d+)\)$/i);
        if (sumMatch) {
            const start = parseRef(sumMatch[1]);
            const end = parseRef(sumMatch[2]);
            if (!start || !end) return null;

            let total = 0;
            const startRow = Math.min(start.row, end.row);
            const endRow = Math.max(start.row, end.row);
            const startCol = Math.min(colToIndex(start.col), colToIndex(end.col));
            const endCol = Math.max(colToIndex(start.col), colToIndex(end.col));
            let hasAnyValue = false;

            for (let row = startRow; row <= endRow; row += 1) {
                for (let colIndex = startCol; colIndex <= endCol; colIndex += 1) {
                    const ref = `${indexToCol(colIndex)}${row}`;
                    hasAnyValue = hasAnyValue || hasValueByRef(ref);
                    total += getNumericValueByRef(ref);
                }
            }

            return { value: total, hasValue: hasAnyValue };
        }

        if (/^[A-Z]+\d+$/i.test(normalized)) {
            return {
                value: getNumericValueByRef(normalized),
                hasValue: hasValueByRef(normalized),
            };
        }

        if (/^#[A-Z0-9\/!?\-]+$/i.test(normalized)) {
            return { value: 0, hasValue: false };
        }

        if (/^\d+(?:\.\d+)?$/.test(normalized)) {
            return { value: Number(normalized), hasValue: true };
        }

        return null;
    }

    function evaluateFormula(formula) {
        const raw = unwrapPassthroughExpression(String(formula || '').replace(/\s+/g, '').replace(/^=/, ''));
        if (!raw) return null;

        const tokenPattern = /([+\-]?)(SUM\([A-Z]+\d+:[A-Z]+\d+\)|[A-Z]+\d+|#[A-Z0-9\/!?\-]+|\d+(?:\.\d+)?)/ig;
        let total = 0;
        let hasAnyValue = false;
        let consumed = '';
        let match;

        while ((match = tokenPattern.exec(raw)) !== null) {
            if (match.index !== consumed.length) {
                return null;
            }

            consumed += match[0];
            const term = getFormulaTermValue(match[2]);
            if (!term) {
                return null;
            }

            if (term.hasValue) {
                hasAnyValue = true;
            }

            total += (match[1] === '-' ? -1 : 1) * term.value;
        }

        if (consumed === raw && consumed !== '') {
            return hasAnyValue ? total : '';
        }

        return null;
    }

    function syncCleanValue(input) {
        if (!input.classList.contains('cell-num')) return;
        input.dataset.cleanValue = cleanNumber(input.value);
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
        input.value = display;
    }

    function showCleanValue(input) {
        if (!input.classList.contains('cell-num')) return;
        if (input.hasAttribute('readonly')) return;
        input.value = input.dataset.cleanValue || cleanNumber(input.value);
    }

    function recalcAllFormulas() {
        Object.entries(formulaCells).forEach(([ref, formula]) => {
            const input = getInputByRef(ref);
            if (!(input instanceof HTMLInputElement)) return;
            const computed = evaluateFormula(formula);
            if (computed === null) return;
            const cleanVal = computed === '' ? '' : formatNumber(computed);
            input.dataset.cleanValue = cleanVal;
            input.dataset.display = cleanVal === '' ? '' : formatVND(cleanVal);
            input.value = input.dataset.display || cleanVal;
        });
    }

    function updateDirtyState() {
        const dirty = Array.from(form.querySelectorAll('[data-ref]')).some((input) => {
            if (!(input instanceof HTMLInputElement) || input.hasAttribute('readonly')) {
                return false;
            }
            const initial = input.dataset.initial ?? '';
            const current = input.classList.contains('cell-num') ? cleanNumber(input.value) : input.value.trim();
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
        if (action === 'clear') {
            payload.value = '';
            const workbookLabel = clearBtn instanceof HTMLButtonElement
                ? (clearBtn.getAttribute('data-workbook-label') || '')
                : '';
            const confirmed = window.confirm(`Xóa toàn bộ số liệu của workbook ${workbookLabel || 'đang xem'}?`);
            if (!confirmed) {
                event.preventDefault();
            }
            return;
        }
        if (action === 'export' || action === 'save') {
            const changes = {};
            form.querySelectorAll('[data-ref]').forEach((input) => {
                if (!(input instanceof HTMLInputElement) || input.hasAttribute('readonly')) {
                    return;
                }
                const initial = input.dataset.initial ?? '';
                const current = input.classList.contains('cell-num') ? cleanNumber(input.value) : input.value.trim();
                if (current !== initial) {
                    changes[input.dataset.ref] = current;
                }
            });
            payload.value = JSON.stringify(changes);
        }
    });

    reloadBtn.addEventListener('click', () => {
        const url = new URL(window.location.href);
        url.searchParams.delete('status');
        url.searchParams.delete('imported');
        window.location.href = url.toString();
    });

    function syncChoiceGroupState(group) {
        if (!(group instanceof HTMLElement)) return;
        group.querySelectorAll('.choice-pill').forEach((pill) => {
            if (!(pill instanceof HTMLElement)) return;
            const input = pill.querySelector('input[type="radio"]');
            pill.classList.toggle('active', input instanceof HTMLInputElement && input.checked);
        });
    }

    document.querySelectorAll('.choice-group').forEach((group) => {
        syncChoiceGroupState(group);
    });

    document.addEventListener('change', (event) => {
        const target = event.target;
        if (!(target instanceof HTMLInputElement) || target.type !== 'radio') {
            return;
        }
        const group = target.closest('.choice-group');
        syncChoiceGroupState(group);
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

    function syncViewQuery(panelName) {
        const url = new URL(window.location.href);
        url.searchParams.delete('view');
        window.history.replaceState({}, '', url.toString());
    }

    function focusPanel(panelName) {
        const panel = panels[panelName] ?? null;
        if (!(panel instanceof HTMLElement)) {
            return;
        }
        highlightTarget(panel);
        if (panelName === 'import') {
            const activeInput = panel.querySelector('.import-input');
            window.setTimeout(() => activeInput?.focus(), 350);
            return;
        }
        if (panelName === 'export') {
            window.setTimeout(() => exportBtn?.focus(), 350);
        }
    }

    function setPanelInteractivity(panel, enabled) {
        if (!(panel instanceof HTMLElement)) {
            return;
        }

        panel.setAttribute('aria-disabled', enabled ? 'false' : 'true');
        panel.querySelectorAll('input, button, select, textarea').forEach((control) => {
            if (
                control instanceof HTMLInputElement ||
                control instanceof HTMLButtonElement ||
                control instanceof HTMLSelectElement ||
                control instanceof HTMLTextAreaElement
            ) {
                control.disabled = !enabled;
            }
        });
    }

    function setActivePanel(panelName, options = {}) {
        const normalizedPanel = supportedPanels.has(panelName) ? panelName : '';
        const focusRequested = options.focus === true;
        const updateUrl = options.updateUrl !== false;
        currentPanel = normalizedPanel;

        if (toolbarPanels instanceof HTMLElement) {
            toolbarPanels.classList.toggle('has-active-panel', normalizedPanel !== '');
        }

        Object.entries(panels).forEach(([name, panel]) => {
            if (!(panel instanceof HTMLElement)) {
                return;
            }
            const isActive = name === normalizedPanel;
            panel.hidden = !isActive;
            panel.classList.toggle('is-active', isActive);
            setPanelInteractivity(panel, isActive);
        });

        panelLaunchers.forEach((launcher) => {
            if (!(launcher instanceof HTMLElement)) {
                return;
            }
            const isActive = launcher.dataset.panelTarget === normalizedPanel;
            launcher.classList.toggle('active', isActive);
            launcher.setAttribute('aria-expanded', isActive ? 'true' : 'false');
        });

        if (updateUrl) {
            syncViewQuery(normalizedPanel);
        }
        if (focusRequested && normalizedPanel !== '') {
            focusPanel(normalizedPanel);
        }
    }

    panelLaunchers.forEach((launcher) => {
        launcher.addEventListener('click', (event) => {
            const targetPanel = launcher.dataset.panelTarget || '';
            if (!supportedPanels.has(targetPanel)) {
                return;
            }
            event.preventDefault();
            const nextPanel = currentPanel === targetPanel ? '' : targetPanel;
            setActivePanel(nextPanel, { focus: nextPanel !== '' });
        });
    });

    const isEmbedded = document.body.classList.contains('embedded');
    if (isEmbedded && window.parent !== window) {
        const getEmbeddedContentHeight = () => {
            const root = document.querySelector('.page-shell');
            if (!(root instanceof HTMLElement)) {
                return Math.max(document.documentElement.scrollHeight, document.body.scrollHeight);
            }

            const candidates = Array.from(root.children).filter((node) => node instanceof HTMLElement);
            const lastElement = candidates.length > 0 ? candidates[candidates.length - 1] : root;
            if (!(lastElement instanceof HTMLElement)) {
                return Math.ceil(root.getBoundingClientRect().height);
            }

            const rootRect = root.getBoundingClientRect();
            const lastRect = lastElement.getBoundingClientRect();
            const computed = window.getComputedStyle(lastElement);
            const marginBottom = parseFloat(computed.marginBottom || '0') || 0;
            return Math.ceil(lastRect.bottom - rootRect.top + marginBottom);
        };

        const notifyParentHeight = () => {
            const height = getEmbeddedContentHeight();
            window.parent.postMessage({
                type: 'khnv-embedded-height',
                height,
            }, window.location.origin);
        };

        const scheduleNotify = () => {
            window.requestAnimationFrame(() => {
                window.requestAnimationFrame(notifyParentHeight);
            });
        };

        scheduleNotify();
        window.addEventListener('load', scheduleNotify);
        window.addEventListener('resize', scheduleNotify);

        if ('ResizeObserver' in window) {
            const resizeObserver = new ResizeObserver(() => {
                scheduleNotify();
            });
            resizeObserver.observe(document.body);
        }

        const mutationObserver = new MutationObserver(() => {
            scheduleNotify();
        });
        mutationObserver.observe(document.body, {
            subtree: true,
            childList: true,
            attributes: true,
            characterData: true,
        });
    }

    setActivePanel(currentPanel, {
        focus: currentPanel !== '',
        updateUrl: false,
    });
})();
</script>
</body>
</html>
