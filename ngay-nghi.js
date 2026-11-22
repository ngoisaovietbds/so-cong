/*
 * Ngày nghỉ lễ, Tết theo quy định của Nhà nước (tra cứu 15/11/2026).
 *
 * Căn cứ:
 *  - Bộ luật Lao động 2019, Điều 112: Tết Dương lịch 1/1; Tết Âm lịch 5 ngày; Giỗ Tổ 10/3 âm lịch;
 *    30/4 và 1/5; Quốc khánh 2/9 và 1 ngày liền kề trước hoặc sau. Điều 111: ngày lễ trùng ngày nghỉ
 *    hằng tuần thì được nghỉ bù vào ngày làm việc kế tiếp.
 *  - Nghị quyết 28/2026/QH16: 24/11 hằng năm là Ngày Văn hóa Việt Nam, nghỉ hưởng nguyên lương (từ 1/7/2026).
 *  - Thông báo 9441/TB-BNV (năm 2026) và 11483/TB-BNV ngày 09/10/2026 (Tết, Quốc khánh 2027) của Bộ Nội vụ.
 *
 * Năm có thông báo chính thức thì dùng đúng lịch nghỉ của công chức trong thông báo đó (gồm cả ngày hoán đổi). Năm chưa có thì tính theo luật (Tết: 1 ngày cuối năm âm
 * + mùng 1 đến mùng 4; Quốc khánh: 2/9 + 1/9), ghi "dự kiến". Ngày nghỉ hằng tuần của Sổ Công là Chủ nhật,
 * nên chỉ ngày lễ rơi vào Chủ nhật mới được nghỉ bù sang ngày kế tiếp.
 */
(function (global) {
  'use strict';

  // Lịch chính thức đã công bố: 'YYYY-MM-DD': tên
  var OFFICIAL = {
    2026: {
      '2026-01-01': 'Tết Dương lịch',
      '2026-02-14': 'Nghỉ Tết Âm lịch', '2026-02-15': 'Nghỉ Tết Âm lịch', '2026-02-16': 'Nghỉ Tết Âm lịch',
      '2026-02-17': 'Tết Âm lịch', '2026-02-18': 'Tết Âm lịch', '2026-02-19': 'Tết Âm lịch',
      '2026-02-20': 'Tết Âm lịch', '2026-02-21': 'Tết Âm lịch', '2026-02-22': 'Nghỉ Tết Âm lịch',
      '2026-04-26': 'Giỗ Tổ Hùng Vương', '2026-04-27': 'Nghỉ bù Giỗ Tổ',
      '2026-04-30': 'Giải phóng miền Nam', '2026-05-01': 'Quốc tế Lao động',
      '2026-08-31': 'Nghỉ hoán đổi dịp Quốc khánh',   // công chức làm bù thứ Bảy 22/8 (TB 9441/TB-BNV)
      '2026-09-01': 'Quốc khánh', '2026-09-02': 'Quốc khánh',
      '2026-11-24': 'Ngày Văn hóa Việt Nam'
    },
    2027: {
      '2027-01-01': 'Tết Dương lịch',
      '2027-02-04': 'Nghỉ Tết Âm lịch', '2027-02-05': 'Nghỉ Tết Âm lịch',
      '2027-02-06': 'Tết Âm lịch', '2027-02-07': 'Tết Âm lịch', '2027-02-08': 'Tết Âm lịch',
      '2027-02-09': 'Tết Âm lịch', '2027-02-10': 'Tết Âm lịch',
      '2027-04-30': 'Giải phóng miền Nam', '2027-05-01': 'Quốc tế Lao động',
      '2027-09-02': 'Quốc khánh', '2027-09-03': 'Quốc khánh',
      '2027-11-24': 'Ngày Văn hóa Việt Nam'
      // Giỗ Tổ 2027 tính theo âm lịch ở dưới (chưa có thông báo nghỉ bù riêng)
    }
  };

  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function key(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }

  /** Tìm ngày dương của các ngày âm trong một năm dương. */
  function lunarDatesInYear(year) {
    var L = global.SoCongLunar;
    var out = { tet: null, gioTo: null, lastOfLunarYear: null };
    if (!L) return out;
    for (var d = new Date(year, 0, 1); d.getFullYear() === year; d = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1)) {
      var l = L.solarToLunar(d.getDate(), d.getMonth() + 1, d.getFullYear());
      if (l.leap) continue;
      if (l.day === 1 && l.month === 1 && !out.tet) out.tet = d;
      if (l.day === 10 && l.month === 3 && !out.gioTo) out.gioTo = d;
    }
    return out;
  }

  var cache = {};

  function buildYear(year) {
    var map = {};
    var add = function (d, name) { if (d.getFullYear() === year && !map[key(d)]) map[key(d)] = name; };
    var lunar = lunarDatesInYear(year);
    var official = OFFICIAL[year];

    if (official) {
      Object.keys(official).forEach(function (k) { map[k] = official[k]; });
      if (lunar.gioTo && !Object.keys(map).some(function (k) { return map[k].indexOf('Giỗ Tổ') !== -1; })) add(lunar.gioTo, 'Giỗ Tổ Hùng Vương');
    } else {
      add(new Date(year, 0, 1), 'Tết Dương lịch');
      if (lunar.tet) {
        add(new Date(lunar.tet.getFullYear(), lunar.tet.getMonth(), lunar.tet.getDate() - 1), 'Tết Âm lịch (dự kiến)');
        for (var i = 0; i < 4; i++) add(new Date(lunar.tet.getFullYear(), lunar.tet.getMonth(), lunar.tet.getDate() + i), 'Tết Âm lịch (dự kiến)');
      }
      if (lunar.gioTo) add(lunar.gioTo, 'Giỗ Tổ Hùng Vương');
      add(new Date(year, 3, 30), 'Giải phóng miền Nam');
      add(new Date(year, 4, 1), 'Quốc tế Lao động');
      add(new Date(year, 8, 1), 'Quốc khánh (dự kiến)');
      add(new Date(year, 8, 2), 'Quốc khánh');
      if (year >= 2026) add(new Date(year, 10, 24), 'Ngày Văn hóa Việt Nam');
    }

    // Nghỉ bù: ngày lễ trùng Chủ nhật thì nghỉ ngày làm việc kế tiếp.
    // Ngày lấy từ lịch chính thức thì không tính lại (thông báo đã gồm nghỉ bù); Tết dự kiến cũng không tính.
    Object.keys(map).sort().forEach(function (k) {
      if (official && official[k]) return;
      if (/Tết Âm lịch/.test(map[k])) return;
      var p = k.split('-');
      var d = new Date(+p[0], +p[1] - 1, +p[2]);
      if (d.getDay() !== 0) return;
      var n = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1);
      while (map[key(n)] || n.getDay() === 0) n = new Date(n.getFullYear(), n.getMonth(), n.getDate() + 1);
      add(n, 'Nghỉ bù ' + map[k]);
    });
    return map;
  }

  function holidayName(d) {
    var y = d.getFullYear();
    if (!cache[y]) cache[y] = buildYear(y);
    return cache[y][key(d)] || null;
  }

  var api = { holidayName: holidayName, _buildYear: buildYear };
  global.SoCongNgayNghi = api;
})(typeof window !== 'undefined' ? window : this);
