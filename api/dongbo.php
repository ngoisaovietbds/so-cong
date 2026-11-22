<?php
/*
 * Sổ Công 1.4.0 – chỗ cất đồng bộ nhiều máy (https://quyenbds.com/socong/api/dongbo.php).
 *
 * Máy chủ CHỈ thấy chuỗi đã mã hoá: app khoá dữ liệu bằng AES-GCM, khoá suy từ mật khẩu nhóm (PBKDF2) ngay trên máy.
 * Máy chủ không có mật khẩu, không có khoá AES. Mỗi lời gọi kèm "thẻ vào nhóm" (cũng suy từ mật khẩu);
 * máy chủ chỉ lưu băm SHA-256 của thẻ để so.
 *
 * Dữ liệu nhóm để NGOÀI public_html: <thư mục nhà>/socong-dongbo/ (web không mở tới được).
 * Việc (POST, thân JSON): tao (tạo nhóm), lay (lấy về), gui (gửi lên, kèm rev đang có),
 * rev (1.4.1: hỏi nhẹ số lần sửa của nhóm, đọc file nhỏ <mã>.rev, app đang mở hỏi 20–60 giây một lần),
 * qtri-nhan / qtri-kiem / doi (1.4.3: mã quản trị, chỉ máy quản trị đổi được mật khẩu nhóm; máy chủ lưu băm mã quản trị 'ha').
 * Sếp chốt 15/11 "Cho dùng thử": máy đang dùng thử, chưa có khoá mở, cũng tạo / vào nhóm được, nên API không đòi khoá mở.
 * Chống lạm dụng: giới hạn lời gọi theo IP và theo nhóm, nhóm tạo mỗi IP mỗi ngày, 1 MB mỗi nhóm, tổng số nhóm và dung lượng,
 * mã nhóm 16 ký tự (80 bit) + thẻ 256 bit.
 */
declare(strict_types=1);

header('Content-Type: application/json; charset=utf-8');
header('Cache-Control: no-store');
header('X-Content-Type-Options: nosniff');
header('Referrer-Policy: no-referrer');

const THAN_TOI_DA = 1500000;          // thân lời gọi tối đa (byte)
const DU_LIEU_TOI_DA = 1400000;       // chuỗi base64 tối đa mỗi nhóm (~1 MB sau giải base64)
const GOI_TOI_DA = 600;               // lời gọi lay/gui/tao mỗi IP trong 10 phút (cả đội chung wifi văn phòng)
const TAO_TOI_DA_IP = 5;              // nhóm tạo mỗi IP mỗi ngày
const GOI_NHOM_TOI_DA = 1200;         // lời gọi lay/gui vào một nhóm mỗi giờ
const REV_IP_TOI_DA = 2000;           // lời gọi hỏi rev mỗi IP trong 10 phút (đếm riêng)
const REV_NHOM_TOI_DA = 6000;         // lời gọi hỏi rev vào một nhóm mỗi giờ (20 máy × 3 lần/phút = 3.600, dư cho nhiều tab)
const TONG_NHOM = 2000;
const TONG_DUNG_LUONG = 200 * 1024 * 1024;
const BO_NHOM_SAU = 400 * 86400;      // nhóm không dùng quá 400 ngày thì xoá
const MA_NHOM_RE = '/^[0-9A-HJKMNP-TV-Z]{16}$/';
const THE_RE = '/^[0-9a-f]{64}$/';
const ADM_RE = '/^[0-9A-HJKMNP-TV-Z]{26}$/';   // mã quản trị 26 ký tự (~128 bit)

function tra(array $kq, int $ma = 200): void
{
    http_response_code($ma);
    echo json_encode($kq, JSON_UNESCAPED_SLASHES);
    exit;
}

function loi(string $loi, int $ma): void
{
    tra(['ok' => false, 'loi' => $loi], $ma);
}

/** Thư mục dữ liệu: <thư mục nhà>/socong-dongbo (api nằm ở <nhà>/public_html/socong/api). */
function thu_muc(string $con = ''): string
{
    $goc = dirname(__DIR__, 3) . '/socong-dongbo';
    $dir = $con === '' ? $goc : $goc . '/' . $con;
    if (!is_dir($dir) && !@mkdir($dir, 0700, true) && !is_dir($dir)) loi('may-chu', 500);
    return $dir;
}

/** Ghi file an toàn: ghi file tạm rồi đổi tên (không bao giờ để file ghi dở). */
function ghi(string $file, string $noi_dung): bool
{
    $tam = $file . '.tam' . bin2hex(random_bytes(4));
    if (@file_put_contents($tam, $noi_dung, LOCK_EX) !== strlen($noi_dung)) { @unlink($tam); return false; }
    @chmod($tam, 0600);
    if (!@rename($tam, $file)) { @unlink($tam); return false; }
    return true;
}

/** Đếm số lần trong một khoảng thời gian (khoá file để hai lời gọi cùng lúc không đếm sai). Trả false nếu vượt. */
function dem(string $ten, int $khoang, int $toi_da): bool
{
    $file = thu_muc('dem') . '/' . $ten . '.json';
    $f = @fopen($file, 'c+');
    if (!$f) return true;                    // không đếm được thì cho qua, không chặn nhầm người dùng
    flock($f, LOCK_EX);
    $d = json_decode((string) stream_get_contents($f), true);
    $now = time();
    if (!is_array($d) || !isset($d['t'], $d['n']) || $now - (int) $d['t'] >= $khoang) $d = ['t' => $now, 'n' => 0];
    $d['n'] = (int) $d['n'] + 1;
    ftruncate($f, 0);
    rewind($f);
    fwrite($f, json_encode($d));
    fflush($f);
    flock($f, LOCK_UN);
    fclose($f);
    return $d['n'] <= $toi_da;
}

/** Muối riêng của máy chủ để băm địa chỉ IP (không lưu IP thật). */
function muoi(): string
{
    $file = thu_muc() . '/.muoi';
    $m = @file_get_contents($file);
    if (!is_string($m) || strlen($m) !== 64) {
        $m = bin2hex(random_bytes(32));
        ghi($file, $m);
    }
    return $m;
}

function file_nhom(string $g): string { return thu_muc('nhom') . '/' . $g . '.json'; }
function file_rev(string $g): string { return thu_muc('nhom') . '/' . $g . '.rev'; }
/** File nhỏ cho việc hỏi rev: băm thẻ + rev, để không phải đọc file dữ liệu tới 1 MB. */
function ghi_rev(string $g, string $h, int $rev): void { ghi(file_rev($g), json_encode(['h' => $h, 'rev' => $rev])); }
function xoa_nhom_file(string $f): void
{
    @unlink($f);
    @unlink(substr($f, 0, -5) . '.khoa');
    @unlink(substr($f, 0, -5) . '.rev');
}

function doc_nhom(string $g): ?array
{
    $d = json_decode((string) @file_get_contents(file_nhom($g)), true);
    return is_array($d) && isset($d['h'], $d['rev']) ? $d : null;
}

/** Nhóm có thật và thẻ đúng. Sai thẻ hay không có nhóm đều trả cùng một lỗi (không dò được nhóm nào tồn tại). */
function vao_nhom(string $g, string $t): array
{
    $d = doc_nhom($g);
    if (!$d || !hash_equals((string) $d['h'], hash('sha256', $t))) loi('khong-vao-duoc', 403);
    return $d;
}

/** Dọn nhóm lâu không dùng; trả [số nhóm, tổng dung lượng] còn lại. */
function don_va_dem(): array
{
    $n = 0;
    $tong = 0;
    $han = time() - BO_NHOM_SAU;
    foreach (glob(thu_muc('nhom') . '/*.json') ?: [] as $f) {
        $mt = @filemtime($f);
        if ($mt !== false && $mt < $han) { xoa_nhom_file($f); continue; }
        // Nhóm tạo xong mà chưa từng gửi dữ liệu (file rất nhỏ) quá 2 ngày: bỏ, để không ai chiếm chỗ bằng nhóm rỗng (rà G-a)
        if ($mt !== false && $mt < time() - 2 * 86400 && (int) @filesize($f) < 300) { xoa_nhom_file($f); continue; }
        $n++;
        $tong += (int) @filesize($f);
    }
    // Bộ đếm cũ hơn 2 ngày không còn dùng tới
    foreach (glob(thu_muc('dem') . '/*.json') ?: [] as $f) {
        $mt = @filemtime($f);
        if ($mt !== false && $mt < time() - 2 * 86400) @unlink($f);
    }
    return [$n, $tong];
}

// ---------- Nhận lời gọi ----------
if (($_SERVER['REQUEST_METHOD'] ?? '') !== 'POST') loi('chi-post', 405);
$than = file_get_contents('php://input', false, null, 0, THAN_TOI_DA + 1);
if (!is_string($than) || strlen($than) > THAN_TOI_DA) loi('qua-lon', 413);
$q = json_decode($than, true);
if (!is_array($q)) loi('sai-dang', 400);

$viec = $q['a'] ?? '';
$ip = hash('sha256', muoi() . '|' . ($_SERVER['REMOTE_ADDR'] ?? ''));
// Hỏi rev đếm riêng (nhẹ, gọi thường xuyên); các việc khác đếm chung
if ($viec === 'rev') {
    if (!dem('rv-ip-' . substr($ip, 0, 32), 600, REV_IP_TOI_DA)) loi('qua-nhieu', 429);
} elseif (!dem('ip-' . substr($ip, 0, 32), 600, GOI_TOI_DA)) loi('qua-nhieu', 429);

$g = $q['g'] ?? '';
$t = $q['t'] ?? '';
if (!is_string($g) || !preg_match(MA_NHOM_RE, $g) || !is_string($t) || !preg_match(THE_RE, $t)) loi('sai-dang', 400);

if ($viec === 'tao') {
    if (!dem('tao-' . substr($ip, 0, 32), 86400, TAO_TOI_DA_IP)) loi('qua-nhieu', 429);
    [$so, $tong] = don_va_dem();
    if ($so >= TONG_NHOM || $tong >= TONG_DUNG_LUONG) loi('het-cho', 507);
    $f = @fopen(file_nhom($g), 'x');            // 'x': chỉ tạo khi chưa có
    if (!$f) loi('ton-tai', 409);
    $now = time();
    $d = ['h' => hash('sha256', $t), 'rev' => 0, 'data' => '', 'at' => $now, 'tao' => $now];
    $adm = $q['adm'] ?? null;
    if (is_string($adm) && preg_match(ADM_RE, $adm)) $d['ha'] = hash('sha256', $adm);   // máy tạo nhóm là máy quản trị
    fwrite($f, json_encode($d));
    fclose($f);
    @chmod(file_nhom($g), 0600);
    ghi_rev($g, $d['h'], 0);
    tra(['ok' => true, 'rev' => 0]);
}

$gb = substr(hash('sha256', muoi() . '|' . $g), 0, 32);   // mã nhóm băm cho bộ đếm (không lưu mã nhóm thật)

if ($viec === 'rev') {
    if (!dem('rv-nhom-' . $gb, 3600, REV_NHOM_TOI_DA)) loi('qua-nhieu', 429);
    $r = json_decode((string) @file_get_contents(file_rev($g)), true);
    if (!is_array($r) || !isset($r['h'], $r['rev'])) {
        // Nhóm tạo từ 1.4.0 chưa có file rev: đọc file chính một lần rồi tạo (giữ khoá nhóm để không ghi rev cũ đè rev mới)
        if (!is_file(file_nhom($g))) loi('khong-vao-duoc', 403);
        $k = @fopen(thu_muc('nhom') . '/' . $g . '.khoa', 'c');
        if (!$k) loi('may-chu', 500);
        flock($k, LOCK_EX);
        $d = doc_nhom($g);
        if ($d) ghi_rev($g, (string) $d['h'], (int) $d['rev']);
        flock($k, LOCK_UN);
        fclose($k);
        if (!$d) loi('khong-vao-duoc', 403);
        $r = ['h' => (string) $d['h'], 'rev' => (int) $d['rev']];
    }
    if (!hash_equals((string) $r['h'], hash('sha256', $t))) loi('khong-vao-duoc', 403);
    tra(['ok' => true, 'rev' => (int) $r['rev']]);
}

// Mỗi nhóm tối đa GOI_NHOM_TOI_DA lời gọi mỗi giờ (đếm theo mã nhóm băm, không lưu mã nhóm thật)
if (!dem('nhom-' . $gb, 3600, GOI_NHOM_TOI_DA)) loi('qua-nhieu', 429);

if ($viec === 'lay') {
    $d = vao_nhom($g, $t);
    // Nhóm chỉ đọc lâu ngày vẫn là nhóm đang dùng: làm mới giờ file để không bị dọn
    $mt = @filemtime(file_nhom($g));
    if ($mt !== false && $mt < time() - 30 * 86400) @touch(file_nhom($g));
    tra(['ok' => true, 'rev' => (int) $d['rev'], 'data' => (string) ($d['data'] ?? ''), 'at' => (int) ($d['at'] ?? 0), 'qt' => isset($d['ha'])]);
}

if ($viec === 'gui') {
    $data = $q['data'] ?? null;
    $rev = $q['rev'] ?? null;
    if (!is_string($data) || $data === '' || !is_int($rev)) loi('sai-dang', 400);
    if (strlen($data) > DU_LIEU_TOI_DA) loi('qua-lon', 413);
    if (!preg_match('/^[A-Za-z0-9+\/]+={0,2}$/', $data)) loi('sai-dang', 400);
    vao_nhom($g, $t);
    // Khoá riêng của nhóm: hai máy gửi cùng lúc thì lần lượt, máy sau thấy rev đã đổi và báo xung đột
    $k = @fopen(thu_muc('nhom') . '/' . $g . '.khoa', 'c');
    if (!$k) loi('may-chu', 500);
    flock($k, LOCK_EX);
    $d = vao_nhom($g, $t);
    if ((int) $d['rev'] !== $rev) {
        flock($k, LOCK_UN);
        fclose($k);
        tra(['ok' => false, 'loi' => 'xung-dot', 'rev' => (int) $d['rev']], 409);
    }
    if (strlen($data) > strlen((string) ($d['data'] ?? ''))) {
        [, $tong] = don_va_dem();
        if ($tong + strlen($data) > TONG_DUNG_LUONG) { flock($k, LOCK_UN); fclose($k); loi('het-cho', 507); }
    }
    $d['rev'] = $rev + 1;
    $d['data'] = $data;
    $d['at'] = time();
    $ok = ghi(file_nhom($g), json_encode($d));
    if ($ok) ghi_rev($g, (string) $d['h'], $d['rev']);
    flock($k, LOCK_UN);
    fclose($k);
    if (!$ok) loi('may-chu', 500);
    tra(['ok' => true, 'rev' => $d['rev']]);
}

/** Giữ khoá riêng của nhóm trong lúc đọc - sửa - ghi. */
function khoa_nhom(string $g)
{
    $k = @fopen(thu_muc('nhom') . '/' . $g . '.khoa', 'c');
    if (!$k) loi('may-chu', 500);
    flock($k, LOCK_EX);
    return $k;
}
function mo_khoa($k): void { flock($k, LOCK_UN); fclose($k); }

if ($viec === 'qtri-nhan' || $viec === 'qtri-kiem' || $viec === 'doi') {
    $adm = $q['adm'] ?? null;
    if (!is_string($adm) || !preg_match(ADM_RE, $adm)) loi('sai-dang', 400);
    $ha = hash('sha256', $adm);

    if ($viec === 'qtri-kiem') {
        $d = vao_nhom($g, $t);
        if (!isset($d['ha']) || !hash_equals((string) $d['ha'], $ha)) loi('sai-ma-quan-tri', 403);
        tra(['ok' => true]);
    }

    if ($viec === 'qtri-nhan') {
        // Nhóm tạo trước 1.4.3 chưa có quản trị: máy đầu tiên nhận (phải có thẻ vào nhóm)
        $k = khoa_nhom($g);
        $d = doc_nhom($g);
        if (!$d || !hash_equals((string) $d['h'], hash('sha256', $t))) { mo_khoa($k); loi('khong-vao-duoc', 403); }
        if (isset($d['ha'])) { mo_khoa($k); loi('da-co-quan-tri', 409); }
        $d['ha'] = $ha;
        $ok = ghi(file_nhom($g), json_encode($d));
        mo_khoa($k);
        if (!$ok) loi('may-chu', 500);
        tra(['ok' => true]);
    }

    // doi: đổi mật khẩu nhóm = thay băm thẻ + dữ liệu đã khoá bằng khoá mới, chỉ máy quản trị
    $t2 = $q['t2'] ?? '';
    $data = $q['data'] ?? null;
    $rev = $q['rev'] ?? null;
    if (!is_string($t2) || !preg_match(THE_RE, $t2) || !is_string($data) || $data === '' || !is_int($rev)) loi('sai-dang', 400);
    if (strlen($data) > DU_LIEU_TOI_DA) loi('qua-lon', 413);
    if (!preg_match('/^[A-Za-z0-9+\/]+={0,2}$/', $data)) loi('sai-dang', 400);
    $k = khoa_nhom($g);
    $d = doc_nhom($g);
    if (!$d || !hash_equals((string) $d['h'], hash('sha256', $t))) { mo_khoa($k); loi('khong-vao-duoc', 403); }
    if (!isset($d['ha']) || !hash_equals((string) $d['ha'], $ha)) { mo_khoa($k); loi('sai-ma-quan-tri', 403); }
    if ((int) $d['rev'] !== $rev) { mo_khoa($k); tra(['ok' => false, 'loi' => 'xung-dot', 'rev' => (int) $d['rev']], 409); }
    $d['h'] = hash('sha256', $t2);
    $d['rev'] = $rev + 1;
    $d['data'] = $data;
    $d['at'] = time();
    $ok = ghi(file_nhom($g), json_encode($d));
    if ($ok) ghi_rev($g, $d['h'], $d['rev']);
    mo_khoa($k);
    if (!$ok) loi('may-chu', 500);
    tra(['ok' => true, 'rev' => $d['rev']]);
}

loi('sai-viec', 400);
