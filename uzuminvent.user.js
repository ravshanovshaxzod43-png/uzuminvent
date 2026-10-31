<?php
declare(strict_types=1);

// CORS: faqat Uzum sahifasidan so'rovlarga ruxsat
$allowedOrigin = 'https://dp.uzum.uz';
header("Access-Control-Allow-Origin: $allowedOrigin");
header('Access-Control-Allow-Methods: POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');
header('Content-Type: application/json; charset=utf-8');

if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    http_response_code(204);
    exit;
}

function respond(int $code, array $data): void {
    http_response_code($code);
    echo json_encode($data, JSON_UNESCAPED_UNICODE);
    exit;
}

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    respond(405, ['ok' => false, 'error' => 'method_not_allowed']);
}

// .env o'qish
function load_env(string $path): array {
    $env = [];
    if (!is_readable($path)) return $env;
    foreach (file($path, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES) as $line) {
        $line = trim($line);
        if ($line === '' || $line[0] === '#' || strpos($line, '=') === false) continue;
        [$k, $v] = array_map('trim', explode('=', $line, 2));
        $v = trim($v, "\"'");
        $env[$k] = $v;
    }
    return $env;
}

$env = load_env(__DIR__ . '/.env');
$token  = $env['TELEGRAM_BOT_TOKEN'] ?? '';
$chatId = $env['TELEGRAM_CHAT_ID'] ?? '';
$sendKey = $env['SEND_KEY'] ?? '';

if ($token === '' || $chatId === '' || $sendKey === '') {
    respond(500, ['ok' => false, 'error' => 'server_not_configured']);
}

// Kalit tekshiruvi (hash_equals — vaqt hujumidan himoya)
$key = $_POST['key'] ?? '';
if (!hash_equals($sendKey, (string)$key)) {
    respond(403, ['ok' => false, 'error' => 'forbidden']);
}

// Matnni tekshirish
$text = trim((string)($_POST['text'] ?? ''));
if ($text === '') {
    respond(400, ['ok' => false, 'error' => 'empty_text']);
}
$text = mb_substr($text, 0, 4000);

// Oddiy rate limit: IP boshiga soatiga 30 ta (fayl asosida)
$ip = $_SERVER['REMOTE_ADDR'] ?? 'unknown';
$rlFile = sys_get_temp_dir() . '/tg_rl_' . md5($ip) . '.json';
$now = time();
$hits = [];
if (is_file($rlFile)) {
    $hits = array_filter(json_decode((string)file_get_contents($rlFile), true) ?: [], fn($ts) => $ts > $now - 3600);
}
if (count($hits) >= 30) {
    respond(429, ['ok' => false, 'error' => 'rate_limited']);
}
$hits[] = $now;
file_put_contents($rlFile, json_encode(array_values($hits)));

// Telegram'ga yuborish (cURL)
$ch = curl_init("https://api.telegram.org/bot{$token}/sendMessage");
curl_setopt_array($ch, [
    CURLOPT_POST => true,
    CURLOPT_POSTFIELDS => http_build_query([
        'chat_id' => $chatId,
        'text' => $text,
        'disable_web_page_preview' => 'true',
    ]),
    CURLOPT_RETURNTRANSFER => true,
    CURLOPT_TIMEOUT => 12,
]);
$response = curl_exec($ch);
$httpCode = (int)curl_getinfo($ch, CURLINFO_HTTP_CODE);
$curlErr  = curl_error($ch);
curl_close($ch);

if ($response === false || $httpCode !== 200) {
    respond(502, ['ok' => false, 'error' => 'telegram_failed', 'detail' => $curlErr ?: $httpCode]);
}

respond(200, ['ok' => true]);
