/*
 * Sổ Công: lưu sẵn các file của app trong máy để mở được khi không có mạng.
 * Chỉ lưu file của app (cùng thư mục), không đụng tới dữ liệu chấm công (dữ liệu nằm trong localStorage).
 * Cách cập nhật: mở app thì dùng ngay bản đã lưu, đồng thời tải bản mới ở nền; lần mở sau dùng bản mới.
 * PHIEN_BAN do build_web.py ghi lại mỗi lần đóng gói.
 */
var PHIEN_BAN = 'socong-1.4.9';
var FILE_APP = [
  './',
  'index.html',
  'style.css',
  'lunar.js',
  'ngay-nghi.js',
  'khoa-cong-khai.js',
  'app.js',
  'pwa.js',
  'manifest.webmanifest',
  'icons/apple-touch-icon.png',
  'icons/icon-192.png',
  'icons/icon-512.png',
  'icons/icon-maskable-512.png',
  'icons/qr-momo.png'
];

self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(PHIEN_BAN).then(function (c) {
      // cache: 'reload' = tải thẳng từ máy chủ, không lấy bản cũ trong bộ nhớ đệm của trình duyệt
      return c.addAll(FILE_APP.map(function (u) { return new Request(u, { cache: 'reload' }); }));
    }).then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.filter(function (k) { return k.indexOf('socong-') === 0 && k !== PHIEN_BAN; })
        .map(function (k) { return caches.delete(k); }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET') return;
  var url = new URL(req.url);
  if (url.origin !== self.location.origin) return;               // chỉ lo file của chính app
  var scope = self.registration.scope;
  if (req.url.indexOf(scope) !== 0) return;                       // ngoài thư mục /socong/ thì bỏ qua
  // Chỉ lo trang chính và đúng các file của app; thư mục khác (ví dụ /socong/apk/) đi thẳng lên máy chủ
  var rel = url.pathname.slice(new URL(scope).pathname.length);
  var isHome = rel === '' || rel === 'index.html';
  if (!isHome && FILE_APP.indexOf(rel) === -1) return;
  if (req.mode === 'navigate' && !isHome) return;
  e.respondWith(
    caches.open(PHIEN_BAN).then(function (c) {
      var key = req.mode === 'navigate' ? './' : req;
      return c.match(key, { ignoreSearch: true }).then(function (cached) {
        var net = fetch(req, { cache: 'no-cache' }).then(function (res) {
          if (res && res.ok && res.type === 'basic') c.put(key, res.clone());
          return res;
        }).catch(function () { return cached; });
        return cached || net;
      });
    })
  );
});
