<?php


const KHNV_BASE_DIR = __DIR__ . '/..';
const KHNV_INPUT_XLSX = KHNV_BASE_DIR . '/INPUT/test.xlsx';
const KHNV_TEMPLATE_DOCX_SAMPLE_1 = KHNV_BASE_DIR . '/OUTPUT/MAU.docx';
const KHNV_TEMPLATE_DOCX_SAMPLE_2 = KHNV_BASE_DIR . '/OUTPUT/MAU_NEW.docx';
const KHNV_TEMPLATE_DOCX_SAMPLE_3 = KHNV_BASE_DIR . '/OUTPUT/1.docx';
const KHNV_TEMPLATE_DOCX_ACTIVE = KHNV_TEMPLATE_DOCX_SAMPLE_3;
const KHNV_TEMPLATE_DOCX = KHNV_TEMPLATE_DOCX_ACTIVE;
const KHNV_MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const KHNV_WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function khnv_is_loopback_host(string $host): bool
{
    $normalized = trim(strtolower($host), '[]');
    return in_array($normalized, ['localhost', '127.0.0.1', '::1'], true);
}

function khnv_is_preferred_lan_ip(string $ip): bool
{
    if (!filter_var($ip, FILTER_VALIDATE_IP, FILTER_FLAG_IPV4)) {
        return false;
    }

    if (preg_match('/^(10\.|192\.168\.|172\.(1[6-9]|2\d|3[0-1])\.)/', $ip)) {
        return true;
    }

    return false;
}

function khnv_detect_server_ip(): string
{
    $candidates = [];

    $serverAddr = (string) ($_SERVER['SERVER_ADDR'] ?? '');
    if ($serverAddr !== '') {
        $candidates[] = $serverAddr;
    }

    $hostIps = @gethostbynamel(gethostname());
    if (is_array($hostIps)) {
        foreach ($hostIps as $ip) {
            $candidates[] = (string) $ip;
        }
    }

    $validIps = [];
    foreach ($candidates as $ip) {
        $ip = trim($ip);
        if (!filter_var($ip, FILTER_VALIDATE_IP, FILTER_FLAG_IPV4)) {
            continue;
        }
        if (khnv_is_loopback_host($ip)) {
            continue;
        }
        $validIps[] = $ip;
    }

    $validIps = array_values(array_unique($validIps));
    foreach ($validIps as $ip) {
        if (khnv_is_preferred_lan_ip($ip)) {
            return $ip;
        }
    }

    return $validIps[0] ?? '';
}

function khnv_current_scheme(): string
{
    $https = strtolower((string) ($_SERVER['HTTPS'] ?? ''));
    if ($https !== '' && $https !== 'off' && $https !== '0') {
        return 'https';
    }

    $scheme = strtolower((string) ($_SERVER['REQUEST_SCHEME'] ?? ''));
    if ($scheme === 'https') {
        return 'https';
    }

    if ((string) ($_SERVER['SERVER_PORT'] ?? '') === '443') {
        return 'https';
    }

    return 'http';
}

function khnv_build_ip_url(string $ipHost): string
{
    $scheme = khnv_current_scheme();
    $requestUri = (string) ($_SERVER['REQUEST_URI'] ?? '/');
    $port = (string) ($_SERVER['SERVER_PORT'] ?? '');

    $portSuffix = '';
    if ($port !== '' && !in_array([$scheme, $port], [['http', '80'], ['https', '443']], true)) {
        $portSuffix = ':' . $port;
    }

    return $scheme . '://' . $ipHost . $portSuffix . $requestUri;
}

function khnv_redirect_localhost_to_ip(): void
{
    if (PHP_SAPI === 'cli') {
        return;
    }

    $hostHeader = (string) ($_SERVER['HTTP_HOST'] ?? $_SERVER['SERVER_NAME'] ?? '');
    if ($hostHeader === '') {
        return;
    }

    $host = strtolower((string) preg_replace('/:\d+$/', '', trim($hostHeader)));
    if (!khnv_is_loopback_host($host)) {
        return;
    }

    $serverIp = khnv_detect_server_ip();
    if ($serverIp === '') {
        return;
    }

    header('Location: ' . khnv_build_ip_url($serverIp), true, 302);
    exit;
}

function khnv_col_to_index(string $col): int
{
    $col = strtoupper($col);
    $n = 0;
    $len = strlen($col);
    for ($i = 0; $i < $len; $i++) {
        $n = ($n * 26) + (ord($col[$i]) - 64);
    }
    return $n;
}

function khnv_index_to_col(int $index): string
{
    $col = '';
    while ($index > 0) {
        $index--;
        $col = chr(($index % 26) + 65) . $col;
        $index = intdiv($index, 26);
    }
    return $col;
}

function khnv_normalize_text_key(string $value): string
{
    $value = preg_replace('/\s+/u', ' ', trim($value)) ?? $value;
    return $value;
}

function khnv_normalize_cell_text(string $value): string
{
    return preg_replace('/\s+/u', ' ', trim($value)) ?? trim($value);
}

function khnv_is_meaningful_header_text(string $value): bool
{
    $value = khnv_normalize_cell_text($value);
    if ($value === '') {
        return false;
    }
    if (preg_match('/^0(?:\.0+)?$/', $value)) {
        return false;
    }
    return true;
}

function khnv_detect_report_year(array $rows, int $fallback = 2026): int
{
    $rowNumbers = array_keys($rows);
    sort($rowNumbers);
    $scanRows = array_values(array_unique(array_merge([1, 2], $rowNumbers)));

    foreach ($scanRows as $rowNum) {
        if (!isset($rows[$rowNum]['cells'])) {
            continue;
        }
        foreach (($rows[$rowNum]['cells'] ?? []) as $cell) {
            $value = khnv_normalize_cell_text((string) ($cell['value'] ?? ''));
            if ($value === '') {
                continue;
            }
            if (preg_match('/(?<!\d)((?:19|20)\d{2})(?!\d)/', $value, $m)) {
                return (int) $m[1];
            }
        }
    }

    return $fallback;
}

function khnv_clean_number_string($value): string
{
    if ($value === null || $value === '') {
        return '';
    }
    if (is_string($value)) {
        $value = str_replace([',', ' '], ['', ''], trim($value));
    }
    if (!is_numeric($value)) {
        return trim((string) $value);
    }
    $num = (float) $value;
    $rounded = (int) round($num);
    return $rounded === 0 ? '0' : (string) $rounded;
}

function khnv_format_report_number($value): string
{
    if ($value === null || $value === '') {
        return '';
    }
    $clean = khnv_clean_number_string($value);
    if ($clean === '') {
        return '';
    }
    if (!is_numeric($clean)) {
        return $clean;
    }
    return number_format((int) round((float) $clean), 0, '.', ',');
}

function khnv_safe_upper(string $value): string
{
    if (function_exists('mb_strtoupper')) {
        return mb_strtoupper($value, 'UTF-8');
    }
    return strtoupper($value);
}


function khnv_node_text(DOMNode $node): string
{
    $text = '';
    foreach ($node->childNodes as $child) {
        if ($child->nodeType === XML_TEXT_NODE || $child->nodeType === XML_CDATA_SECTION_NODE) {
            $text .= $child->nodeValue;
        } else {
            $text .= khnv_node_text($child);
        }
    }
    return $text;
}


function khnv_read_zip_xml(ZipArchive $zip, string $entryName): string
{
    $content = $zip->getFromName($entryName);
    if ($content === false) {
        throw new RuntimeException("Không tìm thấy file $entryName trong file ZIP.");
    }
    return $content;
}

function khnv_load_shared_strings(ZipArchive $zip): array
{
    if ($zip->locateName('xl/sharedStrings.xml') === false) {
        return [];
    }
    $xml = new DOMDocument();
    $xml->preserveWhiteSpace = true;
    $xml->formatOutput = false;
    $xml->loadXML(khnv_read_zip_xml($zip, 'xl/sharedStrings.xml'), LIBXML_NONET);
    $xp = new DOMXPath($xml);
    $xp->registerNamespace('x', KHNV_MAIN_NS);
    $strings = [];
    foreach ($xp->query('//x:si') as $si) {
        $strings[] = khnv_node_text($si);
    }
    return $strings;
}

require_once __DIR__ . '/import.php';
require_once __DIR__ . '/export.php';

