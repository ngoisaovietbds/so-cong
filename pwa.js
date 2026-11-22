/*
 * Sổ Công bản web (iPhone, Android Chrome): đăng ký sw.js để app mở được khi không có mạng.
 * Trong app Android (file://) không chạy gì.
 */
(function () {
  'use strict';
  if (!('serviceWorker' in navigator)) return;
  var ok = location.protocol === 'https:' || location.hostname === 'localhost' || location.hostname === '127.0.0.1';
  if (!ok) return;
  window.addEventListener('load', function () {
    navigator.serviceWorker.register('sw.js').catch(function () { /* không chạy được thì vẫn dùng khi có mạng */ });
  });
})();
