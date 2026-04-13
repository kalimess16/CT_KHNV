<?php
declare(strict_types=1);

function khnv_access_allowed_clients(): array
{
    // Them IP vao day de cho phep truy cap.
    // Xoa IP khoi day de chan truy cap qua PHP.
    // Neu them IP moi, cap nhat them file .htaccess de Apache khong chan truoc khi vao PHP.
    return [
        '10.64.0.108' => 'may Tin hoc',
        '10.64.0.62' => 'may Ms Trang',
        '10.64.0.60' => 'may Ms Tu',
        '10.64.0.83' => 'may Mr Doanh',
        '10.64.0.234' => 'may ptp tin hoc',
    ];
}

function khnv_access_is_loopback_host(string $host): bool
{
    $normalized = trim(strtolower($host), '[]');
    return in_array($normalized, ['localhost', '127.0.0.1', '::1'], true);
}

function khnv_access_is_preferred_lan_ip(string $ip): bool
{
    if (!filter_var($ip, FILTER_VALIDATE_IP, FILTER_FLAG_IPV4)) {
        return false;
    }

    return (bool) preg_match('/^(10\.|192\.168\.|172\.(1[6-9]|2\d|3[0-1])\.)/', $ip);
}

function khnv_access_detect_server_ip(): string
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
        if (khnv_access_is_loopback_host($ip)) {
            continue;
        }
        $validIps[] = $ip;
    }

    $validIps = array_values(array_unique($validIps));
    foreach ($validIps as $ip) {
        if (khnv_access_is_preferred_lan_ip($ip)) {
            return $ip;
        }
    }

    return $validIps[0] ?? '';
}

function khnv_access_current_scheme(): string
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

function khnv_access_build_ip_url(string $ipHost): string
{
    $scheme = khnv_access_current_scheme();
    $requestUri = (string) ($_SERVER['REQUEST_URI'] ?? '/');
    $port = (string) ($_SERVER['SERVER_PORT'] ?? '');

    $portSuffix = '';
    if ($port !== '' && !in_array([$scheme, $port], [['http', '80'], ['https', '443']], true)) {
        $portSuffix = ':' . $port;
    }

    return $scheme . '://' . $ipHost . $portSuffix . $requestUri;
}

function khnv_access_redirect_localhost_to_ip(): void
{
    if (PHP_SAPI === 'cli') {
        return;
    }

    $hostHeader = (string) ($_SERVER['HTTP_HOST'] ?? $_SERVER['SERVER_NAME'] ?? '');
    if ($hostHeader === '') {
        return;
    }

    $host = strtolower((string) preg_replace('/:\d+$/', '', trim($hostHeader)));
    if (!khnv_access_is_loopback_host($host)) {
        return;
    }

    $serverIp = khnv_access_detect_server_ip();
    if ($serverIp === '') {
        return;
    }

    header('Location: ' . khnv_access_build_ip_url($serverIp), true, 302);
    exit;
}

function khnv_access_client_ip(): string
{
    if (PHP_SAPI === 'cli') {
        return 'cli';
    }

    return trim((string) ($_SERVER['REMOTE_ADDR'] ?? ''));
}

function khnv_access_is_allowed_client(?string $clientIp = null): bool
{
    $candidate = trim((string) ($clientIp ?? khnv_access_client_ip()));
    if ($candidate === '') {
        return false;
    }

    return array_key_exists($candidate, khnv_access_allowed_clients());
}

function khnv_access_forbidden_page(string $clientIp): string
{
    $allowedIps = implode(', ', array_keys(khnv_access_allowed_clients()));
    $displayIp = $clientIp !== '' ? $clientIp : 'khong xac dinh';

    return '<!DOCTYPE html>'
        . '<html lang="vi"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">'
        . '<title>403 - Forbidden</title>'
        . '<style>'
        . 'body{margin:0;min-height:100vh;display:grid;place-items:center;padding:24px;background:#fff4f8;color:#4a1634;font:18px/1.6 "Times New Roman",Times,serif;}'
        . '.card{max-width:760px;padding:28px 32px;border:1px solid rgba(114,18,71,.16);border-radius:20px;background:#fff;box-shadow:0 20px 40px rgba(97,20,65,.08);}'
        . 'h1{margin:0 0 12px;font-size:34px;}'
        . 'p{margin:0 0 10px;}'
        . 'code{padding:2px 6px;border-radius:6px;background:#fff0f7;color:#721247;}'
        . '</style></head><body><section class="card">'
        . '<h1>Truy cap bi tu choi</h1>'
        . '<p>He thong chi cho phep cac IP noi bo da duoc khai bao truy cap.</p>'
        . '<p>IP hien tai: <code>' . htmlspecialchars($displayIp, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8') . '</code></p>'
        . '<p>IP duoc phep: <code>' . htmlspecialchars($allowedIps, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8') . '</code></p>'
        . '</section></body></html>';
}

function khnv_access_enforce_client_ip(): void
{
    if (PHP_SAPI === 'cli') {
        return;
    }

    $clientIp = khnv_access_client_ip();
    if (khnv_access_is_allowed_client($clientIp)) {
        return;
    }

    http_response_code(403);
    header('Content-Type: text/html; charset=UTF-8');
    header('Cache-Control: no-store, no-cache, must-revalidate, max-age=0');
    header('Pragma: no-cache');
    echo khnv_access_forbidden_page($clientIp);
    exit;
}
