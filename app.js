/*
 * Sổ Công — lát 1 (bảng + lịch âm), lát 2 (chuyển kỳ), lát 3 (cài đặt).
 * Dữ liệu lưu trong localStorage của máy, khoá "socong.v1".
 */
(function () {
  'use strict';

  var APP_VERSION = '1.4.9';
  var DATA_VERSION = 1;               // phiên bản cấu trúc dữ liệu app này hiểu được
  var STORE_KEY = 'socong.v1';
  var BROKEN_PREFIX = 'socong.v1.hong-';
  var ORIGINAL_PREFIX = 'socong.v1.goc-';
  var DOW = ['CN', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7'];
  var COLOR_RE = /^(#[0-9a-fA-F]{6}|var\(--c[1-9]\))$/;
  var COL_ID_RE = /^[a-z0-9_-]{1,20}$/;
  var DATE_KEY_RE = /^\d{4}-\d{2}-\d{2}$/;
  var AUTO_COL = 'full';              // cột tự điền theo thứ trong tuần
  var OFF_COLS = ['q', 'p', 'full'];  // Chủ nhật và ngày lễ tự điền "nghỉ" ở các cột này
  // 1.4.1: cột thiếu công tự tính theo cặp mã cột (tên cột đổi được nên bám mã). Sếp chốt 15/11: áp cho mọi kỳ, kể cả kỳ cũ
  var MINUS_OF = { qm: 'q', pm: 'p' };
  var KNOWN_KEYS = ['version', 'columns', 'settings', 'cells', 'notes', 'meta', 'mod'];
  var MAX_COLS = 8;
  var NOTE_MAX = 300;

  var DEFAULT_SETTINGS = {
    startDay: 21, payDay: 21,
    showLunar: true, hideEmpty: false, weekTotals: false, showName: false, name: '',
    autoFull: true, fitScreen: true, fontScale: 1,
    feelVibrate: true, feelSound: true   // 1.4.4: rung, tiếng tích khi đổi kỳ (riêng từng máy)
  };
  var DEFAULT_STATE = {
    version: DATA_VERSION,
    columns: [
      { id: 'q', name: 'Q', color: 'var(--c1)' },
      { id: 'p', name: 'P', color: 'var(--c2)' },
      { id: 'full', name: 'Full', color: 'var(--c3)' },
      { id: 'qm', name: 'Q-', color: 'var(--c4)' },
      { id: 'pm', name: 'P-', color: 'var(--c5)' }
    ],
    settings: DEFAULT_SETTINGS,
    cells: {},   // { 'YYYY-MM-DD': { colId: 'giá trị' } }; '' = cố ý để trống (đè lên ô tự điền)
    notes: {},   // { 'YYYY-MM-DD': 'ghi chú' } (lát 4)
    meta: { imports: [] },
    // Giờ sửa (1.4.0, cho gộp file và đồng bộ): c = ô "ngày|cột", n = ghi chú theo ngày, s = cols/startDay/payDay.
    // Có giờ mà không có giá trị nghĩa là đã xoá lúc đó. Số cũ chưa có giờ coi như giờ 0.
    mod: { c: {}, n: {}, s: {} }
  };
  var MOD_SET_KEYS = ['cols', 'startDay', 'payDay'];

  // ---------- Tiện ích ----------
  function clone(o) { return JSON.parse(JSON.stringify(o)); }
  function isPlainObject(o) { return !!o && typeof o === 'object' && !Array.isArray(o); }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function stamp() {
    var d = new Date();
    return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + '-' + pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds());
  }
  function validDay(n, fallback) {
    n = Number(n);
    return Number.isInteger(n) && n >= 1 && n <= 31 ? n : fallback;
  }
  /** Bỏ ký tự điều khiển (trừ xuống dòng, tab): chúng làm hỏng file Excel và không có ích gì. */
  function stripCtl(v) { return String(v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, ''); }
  /** Giờ sửa hợp lệ (mili giây), sai thì 0. */
  function cleanTs(v) { return typeof v === 'number' && isFinite(v) && v > 0 && v < 1e14 ? Math.floor(v) : 0; }
  function esc(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  // ---------- Số liệu ----------
  function normalizeValue(v) {
    var s = String(v).trim();
    if (/^-?\d+\.\d+$/.test(s)) s = s.replace('.', ',');   // 1.5 → 1,5
    return s;
  }
  function parseNumber(v) {
    if (v === undefined || v === null) return null;
    var s = String(v).trim().replace(',', '.');
    if (s === '' || !/^-?\d+(\.\d+)?$/.test(s)) return null;
    return parseFloat(s);
  }
  function formatNumber(n) {
    var r = Math.round(n * 100) / 100;
    return String(r).replace('.', ',');
  }

  // ---------- Đọc / làm sạch dữ liệu ----------
  /**
   * Kiểm từng trường, bỏ phần sai, giữ phần đúng; giữ nguyên các trường app chưa hiểu.
   * Trả { state, dropped } — dropped > 0 nghĩa là có phần bị bỏ, cần cất bản gốc trước khi lưu.
   */
  function sanitizeState(s) {
    var out = clone(DEFAULT_STATE);
    var dropped = 0;
    if (!isPlainObject(s)) return { state: out, dropped: 1 };

    Object.keys(s).forEach(function (k) {
      if (KNOWN_KEYS.indexOf(k) === -1) out[k] = s[k];   // trường của bản mới hơn: giữ nguyên
    });

    if (Array.isArray(s.columns)) {
      var seen = {};
      var cols = [];
      s.columns.forEach(function (c) {
        if (!isPlainObject(c) || typeof c.id !== 'string' || !COL_ID_RE.test(c.id) || seen[c.id]) { dropped++; return; }
        seen[c.id] = true;
        var col = {};
        Object.keys(c).forEach(function (k) {
          var v = c[k];
          if (v === null || typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') col[k] = v;
        });
        col.id = c.id;
        if (typeof c.name !== 'string' || !stripCtl(c.name).trim()) { col.name = c.id; dropped++; } else col.name = stripCtl(c.name).slice(0, 12);
        if (typeof c.color !== 'string' || !COLOR_RE.test(c.color)) {
          col.color = 'var(--c' + ((cols.length % 9) + 1) + ')';
          dropped++;
        }
        cols.push(col);
      });
      if (cols.length) out.columns = cols;
    } else if (s.columns !== undefined) dropped++;

    if (isPlainObject(s.settings)) {
      var st = s.settings;
      Object.keys(st).forEach(function (k) { if (!(k in DEFAULT_SETTINGS)) out.settings[k] = st[k]; });
      if (st.startDay !== undefined) out.settings.startDay = validDay(st.startDay, 21);
      if (st.payDay !== undefined) out.settings.payDay = validDay(st.payDay, 21);
      ['showLunar', 'hideEmpty', 'weekTotals', 'showName', 'autoFull', 'fitScreen', 'feelVibrate', 'feelSound'].forEach(function (k) {
        if (typeof st[k] === 'boolean') out.settings[k] = st[k];
      });
      if (typeof st.name === 'string') out.settings.name = stripCtl(st.name).slice(0, 40);
      var fs = Number(st.fontScale);
      if (isFinite(fs) && fs >= 0.8 && fs <= 1.4) out.settings.fontScale = Math.round(fs * 100) / 100;
    }

    if (isPlainObject(s.cells)) {
      Object.keys(s.cells).forEach(function (k) {
        var row = s.cells[k];
        if (!DATE_KEY_RE.test(k) || !isPlainObject(row)) { dropped++; return; }
        var clean = {};
        Object.keys(row).forEach(function (col) {
          var v = row[col];
          if (!COL_ID_RE.test(col)) { dropped++; return; }
          if (typeof v === 'number' && isFinite(v)) v = String(v);
          if (typeof v !== 'string') { dropped++; return; }
          clean[col] = normalizeValue(stripCtl(v)).slice(0, 12);
        });
        if (Object.keys(clean).length) out.cells[k] = clean;
      });
    }

    if (isPlainObject(s.notes)) {
      Object.keys(s.notes).forEach(function (k) {
        var v = s.notes[k];
        if (!DATE_KEY_RE.test(k) || typeof v !== 'string') { dropped++; return; }
        v = stripCtl(v);
        if (v.trim()) out.notes[k] = v.slice(0, NOTE_MAX);
      });
    } else if (s.notes !== undefined) dropped++;

    if (isPlainObject(s.meta)) {
      Object.keys(s.meta).forEach(function (k) { out.meta[k] = s.meta[k]; });
      if (!Array.isArray(out.meta.imports)) out.meta.imports = [];
    }

    if (isPlainObject(s.mod)) {
      var m = s.mod;
      if (isPlainObject(m.c)) Object.keys(m.c).forEach(function (k) {
        var p = k.split('|'), t = cleanTs(m.c[k]);
        if (p.length === 2 && DATE_KEY_RE.test(p[0]) && COL_ID_RE.test(p[1]) && t) out.mod.c[k] = t; else dropped++;
      });
      if (isPlainObject(m.n)) Object.keys(m.n).forEach(function (k) {
        var t = cleanTs(m.n[k]);
        if (DATE_KEY_RE.test(k) && t) out.mod.n[k] = t; else dropped++;
      });
      if (isPlainObject(m.s)) MOD_SET_KEYS.forEach(function (k) {
        var t = cleanTs(m.s[k]);
        if (t) out.mod.s[k] = t;
      });
    } else if (s.mod !== undefined) dropped++;
    return { state: out, dropped: dropped };
  }

  var loadWarning = '';
  // 1.4.4: máy mới tinh (chưa có dữ liệu): bộ 3 cột cho người dùng cá nhân và hỏi tên lần đầu
  var firstRun = false;
  function newUserState() {
    var st = clone(DEFAULT_STATE);
    st.columns = [
      { id: 'q', name: 'Tôi', color: 'var(--c1)' },
      { id: 'full', name: 'Full', color: 'var(--c3)' },
      { id: 'qm', name: 'Nghỉ', color: 'var(--c4)' }
    ];
    return st;
  }
  var saveBlocked = false;

  /** Giữ 5 bản cất mới nhất của một loại, xoá bản cũ hơn (tên có ngày giờ nên xếp theo chữ là đúng thứ tự). */
  function pruneStash(prefix) {
    try {
      var keys = [];
      for (var i = 0; i < localStorage.length; i++) {
        var k = localStorage.key(i);
        if (k && k.indexOf(prefix) === 0) keys.push(k);
      }
      keys.sort();
      while (keys.length > 5) localStorage.removeItem(keys.shift());
    } catch (e) { /* không sao */ }
  }
  function stash(prefix, raw) {
    var key = prefix + stamp();
    localStorage.setItem(key, raw);
    pruneStash(prefix);
    return key;
  }

  function loadState() {
    var raw = null;
    try {
      raw = localStorage.getItem(STORE_KEY);
    } catch (e) {
      loadWarning = 'Máy không cho app đọc bộ nhớ. Dữ liệu chấm công sẽ không được lưu.';
      saveBlocked = true;
      return clone(DEFAULT_STATE);
    }
    if (!raw) { firstRun = true; return newUserState(); }

    var parsed;
    try {
      parsed = JSON.parse(raw);
    } catch (e) {
      // Dữ liệu cũ đọc không được: cất riêng trước khi dùng bảng trống, để không bị ghi đè mất.
      try {
        var k = stash(BROKEN_PREFIX, raw);
        localStorage.removeItem(STORE_KEY);
        loadWarning = 'Dữ liệu cũ bị hỏng nên đọc không được. App đã cất riêng bản cũ (' + k + '), bảng bắt đầu lại từ trống.';
      } catch (e2) {
        loadWarning = 'Dữ liệu cũ bị hỏng và không cất riêng được. App tạm không lưu để giữ nguyên bản cũ.';
        saveBlocked = true;
      }
      return clone(DEFAULT_STATE);
    }

    if (isPlainObject(parsed) && typeof parsed.version === 'number' && parsed.version > DATA_VERSION) {
      loadWarning = 'Dữ liệu trên máy được lưu bởi một bản Sổ Công mới hơn. App này chỉ xem, không lưu, để khỏi làm hỏng dữ liệu. Hãy cài bản mới nhất.';
      saveBlocked = true;
    }
    var res = sanitizeState(parsed);
    if (res.dropped && !saveBlocked) {
      try {
        var k2 = stash(ORIGINAL_PREFIX, raw);
        loadWarning = 'Có ' + res.dropped + ' mục dữ liệu không đúng dạng nên app bỏ qua. Bản gốc đã cất riêng (' + k2 + ').';
      } catch (e3) {
        loadWarning = 'Có dữ liệu không đúng dạng và không cất riêng được. App tạm không lưu để giữ nguyên bản cũ.';
        saveBlocked = true;
      }
    }
    return res.state;
  }

  function saveState() {
    if (saveBlocked) {
      alert('App đang tạm không lưu để giữ nguyên dữ liệu cũ. Xem dòng thông báo màu vàng ở trên.');
      return false;
    }
    try {
      localStorage.setItem(STORE_KEY, JSON.stringify(state));
      return true;
    } catch (e) {
      alert('Không lưu được dữ liệu trên máy. Hãy kiểm tra bộ nhớ máy.');
      return false;
    }
  }

  // ---------- Giờ sửa, gộp (1.4.0) ----------
  /** Giờ cho lần sửa mới: giờ máy, nhưng luôn lớn hơn giờ cũ của mục đó (đồng hồ các máy lệch chút vẫn đúng thứ tự). */
  function nextTs(prev) { var n = Date.now(); prev = prev || 0; return n > prev ? n : prev + 1; }
  function markCell(key, col) { var k = key + '|' + col; state.mod.c[k] = nextTs(state.mod.c[k]); syncSoon(); }
  function markNote(key) { state.mod.n[key] = nextTs(state.mod.n[key]); syncSoon(); }
  function markSet(name) { state.mod.s[name] = nextTs(state.mod.s[name]); syncSoon(); }

  /** Phần dữ liệu được gộp / đồng bộ (không có khoá mở, dùng thử, cỡ chữ…). */
  function docOf(st) {
    return {
      cells: st.cells, notes: st.notes,
      cols: { columns: st.columns, removed: Array.isArray(st.meta.removedColumns) ? st.meta.removedColumns : [] },
      startDay: st.settings.startDay, payDay: st.settings.payDay, mod: st.mod
    };
  }
  function applyDoc(doc) {
    state.cells = doc.cells;
    state.notes = doc.notes;
    state.columns = doc.cols.columns;
    state.meta.removedColumns = doc.cols.removed;
    state.settings.startDay = doc.startDay;
    state.settings.payDay = doc.payDay;
    state.mod = doc.mod;
  }
  function has(o, k) { return Object.prototype.hasOwnProperty.call(o, k); }

  /**
   * Gộp O vào L theo từng mục, lấy giờ sửa mới hơn. mode 'sync' (đồng bộ): bên "đã xoá" mới hơn thì xoá theo,
   * cùng giờ khác số thì lấy O (bản trên máy chủ) để mọi máy ra cùng kết quả. mode 'file' (gộp từ file): không xoá gì,
   * cùng giờ thì giữ L. Trả { doc, changed (L đổi), push (L có phần O chưa có), đếm số mục }.
   */
  function mergeDocs(L, O, mode) {
    var out = clone(L);
    var r = { doc: out, changed: false, push: false, cellsAdd: 0, cellsUpd: 0, cellsDel: 0, notes: 0, cols: false, sets: [] };
    function same(a, b) { return JSON.stringify(a) === JSON.stringify(b); }
    // 'o' lấy O · 'l' giữ L (O thiếu hoặc cũ hơn) · '=' như nhau
    function decide(lv, lt, ov, ot, always) {
      var lHas = always || lv !== undefined || lt > 0;
      var oHas = always || ov !== undefined || ot > 0;
      if (!oHas) return lHas ? 'l' : '=';
      // Đồng bộ: máy chủ có dấu "đã xoá" mà máy này chưa từng có mục đó thì vẫn chép dấu về (rà 1.4.0 lát 1, B1),
      // nếu không lần gửi sau sẽ làm máy chủ mất dấu và máy thứ ba gửi số cũ lên lại.
      if (!lHas) return (ov === undefined && !(mode === 'sync' && ot > 0)) ? '=' : 'o';
      if (mode === 'file') {
        if (ov === undefined) return 'l';
        if (ot > lt) return 'o';
        return same(lv, ov) ? '=' : 'l';
      }
      if (ot > lt) return 'o';
      if (ot < lt) return 'l';
      return same(lv, ov) ? '=' : 'o';
    }
    function setTs(map, k, t) { if (t > 0) map[k] = t; else delete map[k]; }

    // Ô chấm công
    var keys = {};
    [L, O].forEach(function (d) {
      Object.keys(d.cells).forEach(function (day) { Object.keys(d.cells[day]).forEach(function (c) { keys[day + '|' + c] = 1; }); });
      Object.keys(d.mod.c).forEach(function (k) { keys[k] = 1; });
    });
    Object.keys(keys).forEach(function (k) {
      var p = k.split('|'), day = p[0], c = p[1];
      var lv = L.cells[day] && has(L.cells[day], c) ? L.cells[day][c] : undefined;
      var ov = O.cells[day] && has(O.cells[day], c) ? O.cells[day][c] : undefined;
      var lt = L.mod.c[k] || 0, ot = O.mod.c[k] || 0;
      var d = decide(lv, lt, ov, ot, false);
      if (d === 'l') r.push = true;
      if (d !== 'o') return;
      if (ov === undefined) {
        if (out.cells[day]) { delete out.cells[day][c]; if (!Object.keys(out.cells[day]).length) delete out.cells[day]; }
      } else {
        if (!out.cells[day]) out.cells[day] = {};
        out.cells[day][c] = ov;
      }
      setTs(out.mod.c, k, ot);
      r.changed = true;
      if (lv === ov) return;
      if (ov === undefined) r.cellsDel++; else if (lv === undefined) r.cellsAdd++; else r.cellsUpd++;
    });

    // Ghi chú
    var nk = {};
    [L, O].forEach(function (d) {
      Object.keys(d.notes).forEach(function (k) { nk[k] = 1; });
      Object.keys(d.mod.n).forEach(function (k) { nk[k] = 1; });
    });
    Object.keys(nk).forEach(function (k) {
      var lv = has(L.notes, k) ? L.notes[k] : undefined, ov = has(O.notes, k) ? O.notes[k] : undefined;
      var lt = L.mod.n[k] || 0, ot = O.mod.n[k] || 0;
      var d = decide(lv, lt, ov, ot, false);
      if (d === 'l') r.push = true;
      if (d !== 'o') return;
      if (ov === undefined) delete out.notes[k]; else out.notes[k] = ov;
      setTs(out.mod.n, k, ot);
      r.changed = true;
      if (lv !== ov) r.notes++;
    });

    // Bộ cột (gộp cả bộ) và 2 cài đặt kỳ
    ['cols', 'startDay', 'payDay'].forEach(function (f) {
      var lt = L.mod.s[f] || 0, ot = O.mod.s[f] || 0;
      var d = decide(L[f], lt, O[f], ot, true);
      if (d === 'l') r.push = true;
      if (f === 'cols' && mode === 'file' && d !== 'o') {
        // Gộp từ file mà giữ bộ cột máy này: cột trong file máy này chưa có thì đưa vào "cột đã bỏ"
        // để số đã gộp không bị ẩn mất, bấm Lấy lại là hiện (rà 1.4.4 lần 4, N-3)
        var have = {};
        out.cols.columns.concat(out.cols.removed).forEach(function (c) { if (c && c.id) have[c.id] = 1; });
        var extra = O.cols.columns.concat(O.cols.removed).filter(function (c) { return c && c.id && !have[c.id] && (have[c.id] = 1); });
        if (extra.length) {
          out.cols.removed = out.cols.removed.concat(clone(extra));
          out.mod.s.cols = nextTs(lt);
          r.cols = true;
          r.changed = true;
          r.push = true;
        }
      }
      if (d !== 'o') return;
      if (!same(L[f], O[f])) { if (f === 'cols') r.cols = true; else r.sets.push(f); }
      out[f] = clone(O[f]);
      setTs(out.mod.s, f, ot);
      r.changed = true;
      if (f === 'cols') {
        // Cột máy này có mà bộ cột của O không biết tới: đưa vào "cột đã bỏ" để còn Lấy lại (rà 1.4.0 lát 1, N1)
        var known = {};
        out.cols.columns.concat(out.cols.removed).forEach(function (c) { if (c && c.id) known[c.id] = 1; });
        var lost = L.cols.columns.concat(L.cols.removed).filter(function (c) { return c && c.id && !known[c.id] && (known[c.id] = 1); });
        if (lost.length) {
          out.cols.removed = out.cols.removed.concat(clone(lost));
          out.mod.s.cols = nextTs(ot);       // như một lần sửa mới, để các máy khác cũng nhận
          r.push = true;
        }
      }
    });
    return r;
  }

  /** Nạp số liệu chép từ ảnh app cũ khi Sếp bấm nút, chỉ vào ô còn trống. Trả số ô đã điền. */
  function importOldData() {
    var src = window.SoCongAnhMau;
    if (!src || saveBlocked) return -1;
    var n = 0;
    Object.keys(src.cells).forEach(function (key) {
      var row = state.cells[key] || {};
      Object.keys(src.cells[key]).forEach(function (col) {
        if (row[col] === undefined) { row[col] = src.cells[key][col]; markCell(key, col); n++; }
      });
      if (Object.keys(row).length) state.cells[key] = row;
    });
    if (state.meta.imports.indexOf(src.id) === -1) state.meta.imports.push(src.id);
    saveState();
    return n;
  }

  // ---------- Ngày tháng, kỳ công ----------
  function dateKey(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function keyToDate(k) { var p = k.split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
  function daysInMonth(year, month1) { return new Date(year, month1, 0).getDate(); }
  function addDays(d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); }
  function startOfToday() { var n = new Date(); return new Date(n.getFullYear(), n.getMonth(), n.getDate()); }

  /** Ngày bắt đầu của kỳ có tên "tháng month1 / year". */
  function periodStart(year, month1, startDay) {
    if (startDay <= 1) return new Date(year, month1 - 1, 1);
    var py = month1 === 1 ? year - 1 : year;
    var pm = month1 === 1 ? 12 : month1 - 1;
    return new Date(py, pm - 1, Math.min(startDay, daysInMonth(py, pm)));
  }
  function shiftPeriod(p, n) {
    var idx = p.year * 12 + (p.month - 1) + n;
    return { year: Math.floor(idx / 12), month: (idx % 12) + 1 };
  }
  function samePeriod(a, b) { return a.year === b.year && a.month === b.month; }
  function periodIndex(p) { return p.year * 12 + p.month - 1; }

  function periodDays(year, month1, startDay) {
    var start = periodStart(year, month1, startDay);
    var n = shiftPeriod({ year: year, month: month1 }, 1);
    var next = periodStart(n.year, n.month, startDay);
    var days = [];
    for (var d = start; d < next; d = addDays(d, 1)) days.push(d);
    return days;
  }
  /** Kỳ chứa ngày "day". Kỳ mang tên tháng kết thúc. */
  function periodOf(day, startDay) {
    var p = { year: day.getFullYear(), month: day.getMonth() + 1 };
    if (startDay <= 1) return p;
    var n = shiftPeriod(p, 1);
    return day >= periodStart(n.year, n.month, startDay) ? n : p;
  }
  function todayPeriod() { return periodOf(today, state.settings.startDay); }

  /** Ngày sớm nhất có dữ liệu (khoá dạng YYYY-MM-DD so sánh theo chữ là đúng thứ tự ngày). */
  function earliestKey() {
    var min = null;
    for (var k in state.cells) if (min === null || k < min) min = k;
    return min;
  }

  /** Ô tự điền chỉ bắt đầu từ kỳ có dữ liệu sớm nhất (hoặc kỳ hôm nay nếu chưa có gì). */
  function autoFromDate() {
    var k = earliestKey();
    var p = k ? periodOf(keyToDate(k), state.settings.startDay) : todayPeriod();
    var tp = todayPeriod();
    if (periodIndex(p) > periodIndex(tp)) p = tp;
    return periodStart(p.year, p.month, state.settings.startDay);
  }

  function holidayOf(d) {
    return window.SoCongNgayNghi ? window.SoCongNgayNghi.holidayName(d) : null;
  }
  function isOffDay(d) { return d.getDay() === 0 || !!holidayOf(d); }

  /**
   * Giá trị hiện trong ô: số đã chấm; nếu chưa chấm thì giá trị tự điền:
   *  - Chủ nhật, ngày lễ: "nghỉ" ở cột Q, P, Full. Nếu ngày đó Q hoặc P có chấm số thì Full = số lớn hơn.
   *  - Ngày thường: Full = 2 (T2–T6), 1 (T7).
   */
  function displayValue(key, d, colId, autoFrom) {
    var row = state.cells[key];
    if (row && row[colId] !== undefined) return row[colId];
    if (!state.settings.autoFull || d < autoFrom) return '';
    if (has(MINUS_OF, colId)) {
      // Q- / P-: ngày làm việc, Full là số. Trống → trống; "nghỉ" → Full; số < Full → Full − số; số ≥ Full → trống
      if (isOffDay(d)) return '';
      var full = parseNumber(displayValue(key, d, AUTO_COL, autoFrom));
      if (full === null) return '';
      var base = displayValue(key, d, MINUS_OF[colId], autoFrom);
      if (base === '') return '';
      if (base === 'nghỉ') return formatNumber(full);
      var worked = parseNumber(base);
      if (worked === null || worked >= full) return '';
      return formatNumber(full - worked);
    }
    if (isOffDay(d)) {
      if (OFF_COLS.indexOf(colId) === -1) return '';
      if (colId === AUTO_COL && row) {
        var worked = Math.max(parseNumber(row.q) || 0, parseNumber(row.p) || 0);
        if (worked > 0) return formatNumber(worked);
      }
      return 'nghỉ';
    }
    if (colId === AUTO_COL) return d.getDay() === 6 ? '1' : '2';
    return '';
  }

  /** Ngày nhận lương: đúng số ngày Sếp chọn; tháng không có ngày đó thì lấy ngày cuối tháng. */
  function isPayday(d) {
    return d.getDate() === Math.min(state.settings.payDay, daysInMonth(d.getFullYear(), d.getMonth() + 1));
  }

  // ---------- Trạng thái ----------
  var state = loadState();
  var today = startOfToday();
  var current = todayPeriod();

  var root = document.documentElement;
  var elHead = document.getElementById('sheetHead');
  var elBody = document.getElementById('sheetBody');
  var elWrap = document.querySelector('.sheet-wrap');
  var elPeriod = document.getElementById('periodLabel');
  var elBar = document.getElementById('periodBar');
  var elToday = document.getElementById('btnToday');
  var elNotice = document.getElementById('notice');
  var elNameRow = document.getElementById('nameRow');
  var elName = document.getElementById('nameInput');

  /** Hiện / ẩn bảng có hiệu ứng trượt (CSS .shown). Trạng thái mở đọc qua isShown, không qua .hidden. */
  function setShown(el, show) {
    if (!show && el.dataset.open !== '1') return;   // đã đóng hoặc đang đóng: không đụng hẹn giờ gắn hidden
    clearTimeout(el._hideTimer);
    if (show) {
      el.hidden = false;
      el.classList.remove('leaving');
      el.dataset.open = '1';
      void el.offsetWidth;            // để trình duyệt vẽ vị trí đầu rồi mới trượt vào
      el.classList.add('shown');
    } else {
      el.dataset.open = '';
      el.classList.remove('shown');
      el.classList.add('leaving');
      el._hideTimer = setTimeout(function () { el.hidden = true; el.classList.remove('leaving'); }, 220);
    }
  }
  function isShown(el) { return el.dataset.open === '1'; }

  var noticeSticky = false;
  var noticeTimer = null;
  /** autoMs: tự ẩn sau bấy nhiêu mili giây nếu dòng báo vẫn là câu này (1.4.5) */
  function showNotice(text, sticky, autoMs) {
    clearTimeout(noticeTimer);
    noticeAction = null;
    if (autoMs && text) noticeTimer = setTimeout(function () { if (elNotice.textContent === text) showNotice(''); }, autoMs);
    noticeSticky = !!sticky;
    elNotice.textContent = text;
    elNotice.hidden = !text;
    if (elBody.rows.length) fitRows();   // dòng báo chiếm chỗ: tính lại chiều cao dòng
  }
  if (loadWarning) showNotice(loadWarning);
  var noticeAction = null;   // 1.4.9: dòng báo có việc khi chạm (ví dụ mở trang tải bản mới)
  elNotice.addEventListener('click', function () {
    if (noticeAction) { var act = noticeAction; noticeAction = null; showNotice(''); act(); return; }
    if (!loadWarning && !noticeSticky) showNotice('');
  });

  function applyLook() {
    root.style.setProperty('--font', (16 * state.settings.fontScale) + 'px');
    elNameRow.hidden = !state.settings.showName;
    if (document.activeElement !== elName) elName.value = state.settings.name;
  }

  // ---------- Vẽ bảng ----------
  function render(scrollMode, recenterBar) {
    applyLook();
    var st = state.settings;
    var days = periodDays(current.year, current.month, st.startDay);
    var first = days[0];
    var last = days[days.length - 1];
    var autoFrom = autoFromDate();
    elPeriod.innerHTML = 'Kỳ <b>' + current.month + '/' + current.year + '</b> · ' +
      first.getDate() + '/' + (first.getMonth() + 1) + ' – ' + last.getDate() + '/' + (last.getMonth() + 1);

    // Lấy giá trị hiển thị của cả kỳ một lần
    var grid = days.map(function (d) {
      var key = dateKey(d);
      var vals = {};
      var any = false;   // có ô nào chấm tay chưa (ô tự điền không tính), dùng cho "Ẩn ngày chưa chấm"
      var stored = state.cells[key] || {};
      state.columns.forEach(function (c) {
        vals[c.id] = displayValue(key, d, c.id, autoFrom);
        if (stored[c.id] !== undefined && stored[c.id] !== '') any = true;
      });
      return { d: d, key: key, vals: vals, any: any };
    });

    function sumRows(rows) {
      var t = {}, has = {};
      state.columns.forEach(function (c) { t[c.id] = 0; has[c.id] = false; });
      rows.forEach(function (r) {
        state.columns.forEach(function (c) {
          var n = parseNumber(r.vals[c.id]);
          if (n !== null) { t[c.id] += n; has[c.id] = true; }
        });
      });
      return { t: t, has: has };
    }

    var total = sumRows(grid);
    elSheet.classList.toggle('future-period', first > today);
    var h = '<tr class="head-names"><th class="col-day">Ngày</th>';
    state.columns.forEach(function (c) { h += '<th style="background:' + esc(c.color) + '">' + esc(c.name) + '</th>'; });
    h += '</tr><tr class="head-total"><th class="col-day">Tổng</th>';
    state.columns.forEach(function (c) {
      h += '<th style="background:' + esc(c.color) + '">' + (total.has[c.id] ? formatNumber(total.t[c.id]) : '') + '</th>';
    });
    h += '</tr>';
    elHead.innerHTML = h;

    var b = '';
    var todayKey = dateKey(today);
    var week = [];
    grid.forEach(function (r, i) {
      var d = r.d;
      var dow = d.getDay();
      week.push(r);
      var isToday = r.key === todayKey;
      if (!(st.hideEmpty && !r.any && !isToday)) {
        var cls = [];
        if (dow === 0 || dow === 6) cls.push('weekend');
        if (holidayOf(d)) cls.push('holiday');
        if (i === 0) cls.push('start');
        if (isPayday(d)) cls.push('payday');
        if (isToday) cls.push('today');
        if (d > today) cls.push('future');
        if (state.notes[r.key]) cls.push('has-note');
        var sub = '<span class="dow">' + DOW[dow] + '</span>';
        if (st.showLunar && window.SoCongLunar) {
          var l = window.SoCongLunar.solarToLunar(d.getDate(), d.getMonth() + 1, d.getFullYear());
          sub += '<span class="lunar">' + l.day + '/' + l.month + (l.leap ? 'n' : '') + '</span>';
        }
        b += '<tr class="' + cls.join(' ') + '" data-key="' + r.key + '">';
        b += '<td class="td-day"><div class="day-cell"><span class="day-num">' + d.getDate() +
          '</span><span class="day-sub">' + sub + '</span></div></td>';
        state.columns.forEach(function (c) {
          var v = r.vals[c.id];
          var isText = v !== '' && parseNumber(v) === null;
          b += '<td class="cell' + (isText ? ' text' : '') + '" data-col="' + esc(c.id) + '" style="background:' + esc(c.color) + '">' +
            esc(v) + '</td>';
        });
        b += '</tr>';
      }
      if (st.weekTotals && (dow === 0 || i === grid.length - 1)) {
        var w = sumRows(week);
        b += '<tr class="week-total"><td class="td-day">Tuần</td>';
        state.columns.forEach(function (c) {
          b += '<td style="background:' + esc(c.color) + '">' + (w.has[c.id] ? formatNumber(w.t[c.id]) : '') + '</td>';
        });
        b += '</tr>';
        week = [];
      }
    });
    elBody.innerHTML = b;
    fitRows();

    if (scrollMode === 'today') {
      var todayRow = elBody.querySelector('tr.today');
      if (todayRow) todayRow.scrollIntoView({ block: 'center' });
      else elWrap.scrollTop = 0;
    } else if (scrollMode === 'top') {
      elWrap.scrollTop = 0;
    }
    renderBar(recenterBar !== false && scrollMode !== undefined);
  }

  // ---------- Xem gọn cả kỳ trên một màn hình ----------
  var elSheet = document.getElementById('sheet');
  var ROW_MIN = 17;     // dưới mức này chữ quá nhỏ: để cuộn như thường
  var ROW_NORMAL = 46;  // chiều cao dòng ở chế độ thường

  /** 1.4.8: tên cột dài hơn ô (nhất là khi tăng cỡ chữ) thì thu nhỏ chữ tên cột cho vừa, tối thiểu 55%.
   *  Đọc hết rồi mới ghi, tính thẳng tỉ lệ; nhớ kết quả theo chìa (bề rộng bảng, cỡ chữ, gọn, tên/số) để mỗi lần chấm ô
   *  chỉ gán lại cỡ chữ đã tính, không đo lại (rà 1.4.8 N-1). */
  var headKey = '', headSizes = [];
  function fitHeads() {
    var ths = elHead.querySelectorAll('tr.head-names th:not(.col-day), tr.head-total th:not(.col-day)');
    var texts = [];
    for (var i = 0; i < ths.length; i++) texts.push(ths[i].textContent);
    var key = elSheet.clientWidth + '|' + getComputedStyle(elHead).fontSize + '|' + elSheet.classList.contains('compact') + '|' + texts.join('');
    if (key !== headKey) {
      headKey = key;
      headSizes = [];
      for (i = 0; i < ths.length; i++) ths[i].style.fontSize = '';
      for (i = 0; i < ths.length; i++) {          // chỉ đọc
        var th = ths[i], sw = th.scrollWidth, cw = th.clientWidth;
        if (sw <= cw + 1) { headSizes.push(''); continue; }
        var base = parseFloat(getComputedStyle(th).fontSize) || 16;
        headSizes.push(Math.max(base * (cw - 2) / sw, base * 0.55) + 'px');
      }
    }
    for (i = 0; i < ths.length; i++) ths[i].style.fontSize = headSizes[i] || '';   // chỉ ghi
  }

  function fitRows() {
    elSheet.classList.remove('compact');
    elSheet.style.removeProperty('--row-h');
    if (!state.settings.fitScreen) { fitHeads(); return; }
    var rows = elBody.rows.length;
    if (!rows) { fitHeads(); return; }
    elSheet.classList.add('compact');          // đầu bảng gọn trước rồi mới đo
    fitHeads();
    var avail = elWrap.clientHeight - elHead.offsetHeight - 1;
    var h = Math.floor(avail / rows);
    if (h >= ROW_NORMAL) h = ROW_NORMAL;
    if (h < ROW_MIN) h = ROW_MIN;
    elSheet.style.setProperty('--row-h', h + 'px');
  }
  var fitTimer = null;
  window.addEventListener('resize', function () {
    clearTimeout(fitTimer);
    fitTimer = setTimeout(fitRows, 120);
  });

  // ---------- Thanh chọn kỳ ----------
  function barRange() {
    var tp = todayPeriod();
    var from = shiftPeriod(tp, -12);
    var minAllowed = shiftPeriod(tp, -24);
    var k = earliestKey();
    if (k) {
      var dp = periodOf(keyToDate(k), state.settings.startDay);
      if (periodIndex(dp) < periodIndex(from)) from = dp;
    }
    if (periodIndex(from) < periodIndex(minAllowed)) from = minAllowed;
    var to = shiftPeriod(tp, 3);
    if (periodIndex(current) < periodIndex(from)) from = current;
    if (periodIndex(current) > periodIndex(to)) to = current;
    return { from: from, to: to, todayP: tp };
  }

  function renderBar(recenter) {
    var r = barRange();
    var html = '';
    for (var p = r.from; periodIndex(p) <= periodIndex(r.to); p = shiftPeriod(p, 1)) {
      var cls = 'chip';
      if (samePeriod(p, current)) cls += ' active';
      if (samePeriod(p, r.todayP)) cls += ' now';
      if (periodIndex(p) > periodIndex(r.todayP)) cls += ' future';
      html += '<button type="button" class="' + cls + '" data-y="' + p.year + '" data-m="' + p.month + '">' +
        p.month + '/' + p.year + '</button>';
    }
    var keepScroll = elBar.scrollTop;
    var firstPaint = !elBar.firstChild;
    elBar.innerHTML = html;
    elBar.scrollTop = keepScroll;
    elToday.hidden = samePeriod(current, r.todayP);
    // 1.4.4: kỳ đang xem luôn nằm giữa thanh dưới (recenter giữ để tương thích, không còn dùng)
    centerActiveChip(!firstPaint);
  }

  /** Cuộn thanh kỳ để nút kỳ đang xem nằm giữa. smooth = trượt mượt. */
  function centerActiveChip(smooth) {
    if (barUser) return; // đang vuốt thanh: để người dùng chọn xong đã
    var active = elBar.querySelector('.chip.active');
    if (!active) return;
    // 1.4.6: bánh xe dọc kiểu chọn giờ của iPhone: cuộn theo chiều dọc
    var top = elBar.scrollTop + (active.getBoundingClientRect().top - elBar.getBoundingClientRect().top) -
      (elBar.clientHeight - active.offsetHeight) / 2;
    barAuto = true;
    barUser = false;
    clearTimeout(barAutoTimer);
    barAutoTimer = setTimeout(function () { barAuto = false; }, smooth ? 900 : 80);
    try { elBar.scrollTo({ top: top, behavior: smooth ? 'smooth' : 'auto' }); }
    catch (e) { elBar.scrollTop = top; }
    wheelShape();
  }

  /** Chiều cao đúng của một nấc (offsetHeight làm tròn số nguyên nên lệch dần): lấy từ khoảng cuộn tối đa. */
  function wheelStep(chips) {
    var n = chips.length;
    var span = elBar.scrollHeight - elBar.clientHeight;
    return (n > 1 && span > 0) ? span / (n - 1) : (chips[0].getBoundingClientRect().height || 1);
  }

  /** Hình bánh xe: kỳ ở giữa rõ, kỳ trên / dưới nghiêng ra sau và mờ dần (1.4.6). */
  var wheelRaf = 0;
  function wheelShape() {
    if (wheelRaf) return;
    wheelRaf = requestAnimationFrame(function () {
      wheelRaf = 0;
      // Tính theo vị trí cuộn, không đo khung nút (khung đã bị nghiêng/thu nhỏ sẽ đo sai, gây nhấp nháy; rà 1.4.6 N-1).
      // Nút i nằm đúng giữa khi scrollTop = i × chiều cao một nấc (nhờ khoảng chừa ::before).
      var chips = elBar.querySelectorAll('button.chip');
      if (!chips.length) return;
      var h = wheelStep(chips);
      var pos = elBar.scrollTop / h;
      for (var i = 0; i < chips.length; i++) {
        var k = i - pos;   // cách giữa bao nhiêu nấc
        if (Math.abs(k) > 3) { chips[i].style.transform = ''; chips[i].style.opacity = '0'; chips[i].classList.remove('centered'); continue; }
        var ang = Math.max(-75, Math.min(75, k * 38));
        chips[i].style.transform = 'rotateX(' + (-ang) + 'deg) scale(' + (1 - Math.min(Math.abs(k), 1.5) * 0.12) + ')';
        chips[i].style.opacity = String(Math.max(0, 1 - Math.abs(k) * 0.55));
        chips[i].classList.toggle('centered', Math.abs(k) < 0.5);
      }
    });
  }

  // ---------- 1.4.4: rung + tiếng tích khi đổi kỳ ----------
  var audioCtx = null, iosHaptic = null;
  function tick() {
    if (state.settings.feelVibrate) vibrateTick();
    if (state.settings.feelSound) soundTick();
  }
  function vibrateTick() {
    try {
      var A = window.SoCongAndroid;
      if (A && typeof A.haptic === 'function') { A.haptic(); return; }
      if (navigator.vibrate) { navigator.vibrate(12); return; }
      // iPhone/iPad (iOS 18 trở lên): bấm một công tắc ẩn kiểu "switch" thì máy rung nhẹ. Máy cũ hơn: không rung, không lỗi.
      if (!iosHaptic) {
        iosHaptic = document.createElement('label');
        iosHaptic.setAttribute('aria-hidden', 'true');
        iosHaptic.style.cssText = 'position:fixed;left:-9999px;top:0;width:1px;height:1px;overflow:hidden;opacity:0;pointer-events:none';
        var inp = document.createElement('input');
        inp.type = 'checkbox';
        inp.setAttribute('switch', '');
        inp.tabIndex = -1;
        iosHaptic.appendChild(inp);
        document.body.appendChild(iosHaptic);
      }
      iosHaptic.click();
    } catch (e) { /* máy không hỗ trợ rung: bỏ qua */ }
  }
  function unlockAudio() {
    if (!state.settings.feelSound) return;
    try {
      var AC = window.AudioContext || window.webkitAudioContext;
      if (!AC) return;
      if (!audioCtx) audioCtx = new AC();
      if (audioCtx.state === 'suspended') audioCtx.resume();
    } catch (e) { /* bỏ qua */ }
  }
  ['touchend', 'click'].forEach(function (ev) {
    document.addEventListener(ev, function once() {
      unlockAudio();
      if (audioCtx && audioCtx.state === 'running') document.removeEventListener(ev, once, true);
    }, true);
  });
  function soundTick() {
    try {
      var AC = window.AudioContext || window.webkitAudioContext;
      if (!AC) return;
      if (!audioCtx) audioCtx = new AC();
      if (audioCtx.state === 'suspended') audioCtx.resume();
      var t = audioCtx.currentTime;
      var osc = audioCtx.createOscillator(), g = audioCtx.createGain();
      osc.type = 'triangle';
      osc.frequency.setValueAtTime(1500, t);
      osc.frequency.exponentialRampToValueAtTime(900, t + 0.05);
      g.gain.setValueAtTime(0.0001, t);
      g.gain.exponentialRampToValueAtTime(0.18, t + 0.005);
      g.gain.exponentialRampToValueAtTime(0.0001, t + 0.06);
      osc.connect(g); g.connect(audioCtx.destination);
      osc.start(t); osc.stop(t + 0.07);
    } catch (e) { /* không phát được tiếng: bỏ qua */ }
  }

  function goTo(p, silent) {
    if (samePeriod(p, current)) return;
    if (noticeSticky) showNotice('');
    var dir = periodIndex(p) > periodIndex(current) ? 'next' : 'prev';
    current = p;
    barUser = false; clearTimeout(barScrollTimer); // đổi kỳ xong thì thanh luôn về giữa
    if (!silent) tick();
    closePicker();
    render(samePeriod(p, todayPeriod()) ? 'today' : 'top');
    elSheet.classList.remove('enter-next', 'enter-prev');
    void elSheet.offsetWidth;
    elSheet.classList.add('enter-' + dir);
  }

  elBar.addEventListener('click', function (e) {
    var btn = e.target.closest('button.chip');
    if (!btn) return;
    barUser = false;
    clearTimeout(barScrollTimer);
    goTo({ year: +btn.getAttribute('data-y'), month: +btn.getAttribute('data-m') });
  });
  elToday.addEventListener('click', function () { goTo(todayPeriod()); });

  // 1.4.4: vuốt thanh kỳ như bánh xe: nút nào dừng ở giữa thì mở kỳ đó
  var barAuto = false, barAutoTimer = null, barUser = false, barScrollTimer = null;
  function chipAtCenter() {
    // 1.4.6: tính theo vị trí cuộn (nút i ở giữa khi scrollTop = i × chiều cao một nấc), không đo khung đã nghiêng
    var chips = elBar.querySelectorAll('button.chip');
    if (!chips.length) return null;
    var i = Math.round(elBar.scrollTop / wheelStep(chips));
    return chips[Math.max(0, Math.min(chips.length - 1, i))];
  }
  // barMoved: từ lúc chạm, thanh có thật sự cuộn chưa. Chạm mà không cuộn (kéo dọc, bấm giữ, lăn dọc) thì thả tay là
  // bỏ cờ barUser ngay, để cuộn không do người dùng (xoay máy…) không bị coi là vuốt (rà 1.4.4 lần 2, N-2).
  var barMoved = false, barWheelTimer = null;
  function barRelease() { if (!barMoved) barUser = false; }
  ['touchstart', 'pointerdown', 'wheel'].forEach(function (ev) {
    elBar.addEventListener(ev, function () {
      barUser = true; barMoved = false; barAuto = false; clearTimeout(barAutoTimer);
      if (ev === 'wheel') { clearTimeout(barWheelTimer); barWheelTimer = setTimeout(barRelease, 300); }
    }, { passive: true });
  });
  // Không nghe pointercancel: trình duyệt phát nó ngay khi ngón tay bắt đầu kéo, trước sự kiện cuộn đầu tiên (rà lần 3, C-1)
  ['touchend', 'touchcancel', 'pointerup'].forEach(function (ev) {
    elBar.addEventListener(ev, barRelease, { passive: true });
  });
  window.addEventListener('resize', function () { if (!barUser) centerActiveChip(false); });
  var wheelLast = null, wheelTickAt = 0;   // nút đang ở giữa lúc vuốt, để kêu "tích" mỗi nấc như bánh xe iPhone
  elBar.addEventListener('scroll', function () {
    wheelShape();
    if (barUser && !barAuto) barMoved = true;
    if (barAuto || !barUser) { wheelLast = null; return; }
    var c0 = chipAtCenter();
    // Vuốt mạnh: tối đa một tiếng "tích" mỗi 45 ms (rà 1.4.6 G-1)
    if (c0 && wheelLast && c0 !== wheelLast && Date.now() - wheelTickAt > 45) { wheelTickAt = Date.now(); tick(); }
    wheelLast = c0;
    clearTimeout(barScrollTimer);
    barScrollTimer = setTimeout(function () {
      barUser = false;
      wheelLast = null;
      var c = chipAtCenter();
      if (!c) return;
      var p = { year: +c.getAttribute('data-y'), month: +c.getAttribute('data-m') };
      if (samePeriod(p, current)) centerActiveChip(true);
      else goTo(p, true);
    }, 160);
  }, { passive: true });

  // Vuốt ngang trên bảng để sang kỳ trước / sau
  // 1.4.1: bảng đang ở trên cùng mà kéo xuống quá 80 px rồi thả → Đồng bộ ngay (không chặn cuộn)
  var touchX = null, touchY = null, touchTop = 0, pullReady = false;
  var elPull = document.getElementById('pullHint');
  function setPull(on) { if (pullReady !== on) { pullReady = on; elPull.hidden = !on; } }
  elWrap.addEventListener('touchstart', function (e) {
    setPull(false);
    if (e.touches.length !== 1) { touchX = null; return; }
    touchX = e.touches[0].clientX;
    touchY = e.touches[0].clientY;
    touchTop = elWrap.scrollTop;
  }, { passive: true });
  elWrap.addEventListener('touchmove', function (e) {
    if (touchX === null || !syncCfg || touchTop > 0) return;
    var dx = e.touches[0].clientX - touchX, dy = e.touches[0].clientY - touchY;
    setPull(elWrap.scrollTop <= 0 && dy > 80 && dy > Math.abs(dx) * 2);
  }, { passive: true });
  elWrap.addEventListener('touchend', function (e) {
    if (touchX === null) return;
    var dx = e.changedTouches[0].clientX - touchX;
    var dy = e.changedTouches[0].clientY - touchY;
    touchX = null;
    if (pullReady) { setPull(false); if (syncCfg) syncNow(true); return; }
    if (Math.abs(dx) > 70 && Math.abs(dx) > Math.abs(dy) * 2) goTo(shiftPeriod(current, dx < 0 ? 1 : -1));
  }, { passive: true });
  elWrap.addEventListener('touchcancel', function () { touchX = null; setPull(false); }, { passive: true });

  // ---------- Bảng nút nhanh: 2 · 1 · nghỉ ----------
  var elPicker = document.getElementById('picker');
  var elBackdrop = document.getElementById('pickerBackdrop');
  var elTitle = document.getElementById('pickerTitle');
  var editing = null; // { key, col }

  function openPicker(td) {
    if (isLocked()) return openKey('Hết ' + TRIAL_DAYS + ' ngày dùng thử. Nhập khoá mở để chấm tiếp.');
    var key = td.parentNode.getAttribute('data-key');
    var col = td.getAttribute('data-col');
    var colDef = state.columns.filter(function (c) { return c.id === col; })[0];
    var p = key.split('-');
    editing = { key: key, col: col };
    td.classList.add('selected');
    var hol = holidayOf(keyToDate(key));
    if (keyToDate(key) > today) hol = (hol ? hol + ' · ' : '') + 'Ngày chưa tới – chỉ nên ghi trước ngày nghỉ dự kiến';
    elTitle.innerHTML = 'Ngày <b>' + parseInt(p[2], 10) + '/' + parseInt(p[1], 10) + '</b> · cột <b>' + esc(colDef ? colDef.name : col) + '</b>' +
      (hol ? '<div class="picker-hol">' + esc(hol) + '</div>' : '');
    // 1.4.2: cột tự tính (Full, Q-, P-) có nút "Tự tính" để bỏ số chấm tay, quay về số tự tính
    var canAuto = state.settings.autoFull && (col === AUTO_COL || has(MINUS_OF, col));
    elPickAuto.hidden = !canAuto;
    elPickRow.classList.toggle('three', canAuto);
    setShown(elPicker, true);
    setShown(elBackdrop, true);
  }

  var elPickAuto = document.querySelector('#picker [data-action="auto"]');
  var elPickRow = elPickAuto.parentNode;

  function closePicker() {
    setShown(elPicker, false);
    if (!settingsOpen() && !dayOpen() && !helpOpen() && !(elKey && keyOpen()) && !(elSyncSheet && syncSheetOpen())) setShown(elBackdrop, false);
    editing = null;
    var sel = elBody.querySelector('td.selected');
    if (sel) sel.classList.remove('selected');
  }

  function setValue(value) {
    if (!editing) return;
    var v = normalizeValue(stripCtl(value)).slice(0, 12);
    var row = state.cells[editing.key] || {};
    if (v === '') {
      // Ô tự điền: lưu '' để giữ ô trống. Ô thường: xoá hẳn.
      if ((OFF_COLS.indexOf(editing.col) !== -1 || has(MINUS_OF, editing.col)) && state.settings.autoFull) row[editing.col] = '';
      else delete row[editing.col];
    } else row[editing.col] = v;
    // 1.4.2: sửa Q (hoặc P) thì Q- (hoặc P-) ngày đó tự tính lại: bỏ số cũ trong ô (kể cả số nạp từ app cũ)
    Object.keys(MINUS_OF).forEach(function (m) {
      if (MINUS_OF[m] === editing.col && row[m] !== undefined) { delete row[m]; markCell(editing.key, m); }
    });
    if (Object.keys(row).length) state.cells[editing.key] = row;
    else delete state.cells[editing.key];
    markCell(editing.key, editing.col);
    saveState();
    var keep = elWrap.scrollTop;
    closePicker();
    render();
    elWrap.scrollTop = keep;
  }

  /** 1.4.2: bỏ số chấm tay ở ô tự tính (Full, Q-, P-) để ô hiện lại số tự tính. */
  function setAuto() {
    if (!editing) return;
    var row = state.cells[editing.key];
    if (row && row[editing.col] !== undefined) {
      delete row[editing.col];
      if (!Object.keys(row).length) delete state.cells[editing.key];
      markCell(editing.key, editing.col);
      saveState();
    }
    var keep = elWrap.scrollTop;
    closePicker();
    render();
    elWrap.scrollTop = keep;
  }

  elBody.addEventListener('click', function (e) {
    var td = e.target.closest('td.cell');
    if (td) return openPicker(td);
    var day = e.target.closest('td.td-day');
    if (day && day.parentNode.getAttribute('data-key')) openDay(day.parentNode.getAttribute('data-key'));
  });

  // ---------- Ghi chú theo ngày (lát 4) ----------
  var elDay = document.getElementById('daySheet');
  var elDayTitle = document.getElementById('dayTitle');
  var elDaySub = document.getElementById('daySub');
  var elNote = document.getElementById('noteInput');
  var noteKey = null;
  var noteOrig = '';                  // chữ lúc mở bảng: đóng bảng chỉ lưu khi người dùng đã sửa

  function dayOpen() { return isShown(elDay); }

  function openDay(key) {
    if (isLocked()) return openKey('Hết ' + TRIAL_DAYS + ' ngày dùng thử. Nhập khoá mở để ghi chú tiếp.');
    closePicker();
    var d = keyToDate(key);
    var l = window.SoCongLunar ? window.SoCongLunar.solarToLunar(d.getDate(), d.getMonth() + 1, d.getFullYear()) : null;
    var hol = holidayOf(d);
    noteKey = key;
    elDayTitle.textContent = (d.getDay() === 0 ? 'Chủ nhật' : 'Thứ ' + (d.getDay() + 1)) + ', ' + d.getDate() + '/' + (d.getMonth() + 1) + '/' + d.getFullYear();
    elDaySub.textContent = (l ? 'Âm lịch ' + l.day + '/' + l.month + (l.leap ? ' nhuận' : '') : '') + (hol ? ' · ' + hol : '') +
      (isPayday(d) ? ' · Ngày nhận lương' : '');
    elDaySub.classList.toggle('is-payday', isPayday(d));
    elNote.value = noteOrig = state.notes[key] || '';
    setShown(elDay, true);
    setShown(elBackdrop, true);
  }
  /** Đóng bảng ghi chú. Mặc định tự lưu nếu chữ khác bản đã lưu (Quay lại, chạm nền mờ, Đóng đều không mất chữ). */
  function closeDay(skipSave) {
    if (!skipSave && noteKey && elNote.value !== noteOrig && elNote.value !== (state.notes[noteKey] || '')) return saveNote(elNote.value);
    setShown(elDay, false);
    setShown(elBackdrop, false);
    noteKey = null;
    if (document.activeElement === elNote) elNote.blur();
  }
  function saveNote(text) {
    if (!noteKey) return;
    var v = stripCtl(text).slice(0, NOTE_MAX);
    var had = state.notes[noteKey];
    if (v.trim()) state.notes[noteKey] = v; else delete state.notes[noteKey];
    if (state.notes[noteKey] !== had) markNote(noteKey);
    saveState();
    var keep = elWrap.scrollTop;
    closeDay(true);
    render();
    elWrap.scrollTop = keep;
  }
  elDay.addEventListener('click', function (e) {
    var btn = e.target.closest('button');
    if (!btn) return;
    var a = btn.getAttribute('data-action');
    if (a === 'note-save') saveNote(elNote.value);
    else if (a === 'note-clear') { if (!state.notes[noteKey] || confirm('Xoá ghi chú của ngày này?')) saveNote(''); }
    else if (a === 'note-close') closeDay();
  });

  /** Chép chữ ra bộ nhớ tạm; WebView mở từ file:// có thể không có navigator.clipboard nên có cách dự phòng. */
  function copyText(text) {
    function fallback() {
      var ta = document.createElement('textarea');
      ta.value = text;
      ta.setAttribute('readonly', '');
      ta.style.position = 'fixed';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      var ok = false;
      try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
      document.body.removeChild(ta);
      return ok;
    }
    // App Android: chép qua cầu nối của app (chắc chắn nhất trong WebView)
    if (window.SoCongAndroid && typeof window.SoCongAndroid.copy === 'function') {
      try { return Promise.resolve(!!window.SoCongAndroid.copy(text)); } catch (e) { /* thử cách khác */ }
    }
    if (navigator.clipboard && window.isSecureContext) {
      return navigator.clipboard.writeText(text).then(function () { return true; }, function () { return fallback(); });
    }
    return Promise.resolve(fallback());
  }

  function notesText(p) {
    var days = periodDays(p.year, p.month, state.settings.startDay);
    var lines = [];
    days.forEach(function (d) {
      var k = dateKey(d);
      if (state.notes[k]) lines.push(DOW[d.getDay()] + ' ' + d.getDate() + '/' + (d.getMonth() + 1) + ': ' + state.notes[k].replace(/\s*\n\s*/g, '; '));
    });
    if (!lines.length) return '';
    var head = 'Ghi chú kỳ ' + p.month + '/' + p.year + (state.settings.name ? ' – ' + state.settings.name : '');
    return head + '\n' + lines.join('\n');
  }

  elPicker.addEventListener('click', function (e) {
    var btn = e.target.closest('button');
    if (!btn) return;
    var action = btn.getAttribute('data-action');
    if (action === 'close') return closePicker();
    if (action === 'clear') return setValue('');
    if (action === 'auto') return setAuto();
    var val = btn.getAttribute('data-value');
    if (val !== null) setValue(val);
  });

  // ---------- Lát 5: file sao lưu, Excel ----------
  function utf8Bytes(text) { return new TextEncoder().encode(text); }
  function toBase64(bytes) {
    var bin = '';
    for (var i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(bin);
  }

  /**
   * Xuất file. Trên app Android: lưu vào Tải xuống/SoCong rồi mở bảng gửi (Zalo, Drive…).
   * Trên trình duyệt: tải file về. Trả câu báo kết quả, '' nếu không làm được.
   */
  var lastExportSaved = false;
  function exportFile(name, mime, bytes, onSaved) {
    lastExportSaved = false;
    var A = window.SoCongAndroid;
    if (A && typeof A.saveFile === 'function') {
      var b64 = toBase64(bytes);
      var where = '';
      try { where = A.saveFile(name, mime, b64) || ''; } catch (e) { where = ''; }
      var shared = false;
      try { shared = !!A.shareFile(name, mime, b64); } catch (e) { shared = false; }
      lastExportSaved = !!where;
      if (where) return 'Đã lưu file ' + name + ' vào ' + where + '.';
      return shared ? 'Đã mở bảng gửi file ' + name + '.' : '';
    }
    // Điện thoại (iPhone, app web): mở bảng Chia sẻ của máy để lưu vào Tệp, gửi Zalo, Drive…
    var touch = window.matchMedia && window.matchMedia('(pointer: coarse)').matches;
    if (touch && navigator.share && typeof File === 'function') {
      try {
        var file = new File([bytes], name, { type: mime });
        if (navigator.canShare && navigator.canShare({ files: [file] })) {
          navigator.share({ files: [file] }).then(function () {
            if (typeof onSaved === 'function') onSaved();
          }, function () { /* Sếp bấm Huỷ: không ghi gì */ });
          return 'Đã mở bảng Chia sẻ: chọn "Lưu vào Tệp" hoặc gửi Zalo, Drive.';
        }
      } catch (e) { /* thử tải file như máy tính */ }
    }
    try {
      var url = URL.createObjectURL(new Blob([bytes], { type: mime }));
      var a = document.createElement('a');
      a.href = url;
      a.download = name;
      document.body.appendChild(a);
      a.click();
      setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 2000);
      lastExportSaved = true;
      return 'Đã tải file ' + name + '.';
    } catch (e) {
      return '';
    }
  }

  function doBackup() {
    var payload = { app: 'SoCong', format: 1, appVersion: APP_VERSION, exportedAt: new Date().toISOString(), data: state };
    var name = 'SoCong_SaoLuu_' + stamp() + '.json';
    var msg = exportFile(name, 'application/json', utf8Bytes(JSON.stringify(payload)), function () {
      state.meta.lastBackup = dateKey(today);
      saveState();
    });
    closeSettings();
    if (msg) {
      if (lastExportSaved) { state.meta.lastBackup = dateKey(today); saveState(); }
      showNotice(msg + ' Nên gửi file này vào Zalo của mình hoặc Google Drive để giữ an toàn. Chạm để ẩn.');
    } else showNotice('Không lưu được file sao lưu. Chạm để ẩn.');
  }

  function stashCurrent(tag) {
    try {
      // Đang khoá lưu (dữ liệu của bản mới hơn / bị hỏng): cất nguyên văn bản trong máy, không cất bản đã làm sạch.
      var raw = saveBlocked ? localStorage.getItem(STORE_KEY) : null;
      localStorage.setItem('socong.v1.' + tag + '-' + stamp(), raw !== null ? raw : JSON.stringify(state));
      pruneStash('socong.v1.' + tag + '-');
      return true;
    } catch (e) { return false; }
  }

  function summarize(st) {
    var keys = Object.keys(st.cells || {}).filter(function (k) {
      var row = st.cells[k];
      return Object.keys(row).some(function (c) { return row[c] !== ''; });
    }).sort();
    var notes = Object.keys(st.notes || {}).length;
    return { days: keys.length, from: keys[0], to: keys[keys.length - 1], notes: notes };
  }
  function viDate(k) { var p = k.split('-'); return +p[2] + '/' + +p[1] + '/' + p[0]; }

  /** Đọc file sao lưu, kiểm dạng. Trả { obj, res } hoặc null (đã báo lỗi). */
  function readBackup(text) {
    var obj;
    try { obj = JSON.parse(text); } catch (e) { alert('File này không phải file sao lưu của Sổ Công.'); return null; }
    if (!isPlainObject(obj) || obj.app !== 'SoCong' || !isPlainObject(obj.data)) { alert('File này không phải file sao lưu của Sổ Công.'); return null; }
    if (typeof obj.data.version === 'number' && obj.data.version > DATA_VERSION) {
      alert('File được tạo bởi bản Sổ Công mới hơn. Hãy cài bản mới nhất rồi thử lại.');
      return null;
    }
    return { obj: obj, res: sanitizeState(obj.data) };
  }
  function backupDate(obj) { return obj.exportedAt ? new Date(obj.exportedAt).toLocaleString('vi-VN') : '?'; }

  /** Gộp file sao lưu vào máy (1.4.0): theo từng ô, từng ghi chú, lấy lần sửa mới hơn, không xoá gì. */
  function mergeBackupText(text) {
    if (saveBlocked) { alert('App đang tạm không lưu để giữ nguyên dữ liệu cũ (xem dòng màu vàng). Hãy dùng "Nhập lại (thay toàn bộ)" nếu muốn thay bằng file.'); return; }
    var b = readBackup(text);
    if (!b) return;
    var m = mergeDocs(docOf(state), docOf(b.res.state), 'file');
    var nCells = m.cellsAdd + m.cellsUpd;
    if (!nCells && !m.notes && !m.cols && !m.sets.length) {
      alert('File sao lưu ngày ' + backupDate(b.obj) + ' không có gì mới hơn số trên máy. Không cần gộp.');
      return;
    }
    var SET_NAME = { startDay: 'ngày bắt đầu kỳ', payDay: 'ngày nhận lương' };
    var info = 'Gộp file sao lưu ngày ' + backupDate(b.obj) + ' vào máy này:\n' +
      '· ' + m.cellsAdd + ' ô chấm công mới, ' + m.cellsUpd + ' ô được sửa theo file (file sửa sau)\n' +
      '· ' + m.notes + ' ghi chú mới hoặc sửa\n' +
      (m.cols ? '· Cột (tên, màu, thứ tự) theo file\n' : '') +
      (m.sets.length ? '· ' + m.sets.map(function (k) { return SET_NAME[k]; }).join(', ') + ' theo file\n' : '') +
      (b.res.dropped ? '· ' + b.res.dropped + ' mục không đúng dạng sẽ bỏ qua\n' : '') +
      '\nKhông xoá gì trên máy. Gộp?\n(Bản hiện tại được cất riêng trong máy trước khi gộp.)';
    if (!confirm(info)) return;
    if (!stashCurrent('truoc-gop')) { alert('Máy hết chỗ để cất dữ liệu hiện tại, chưa gộp.'); return; }
    var oldDoc = clone(docOf(state));
    var startChanged = m.sets.indexOf('startDay') !== -1;
    applyDoc(m.doc);
    if (!saveState()) { applyDoc(oldDoc); return; }
    syncSoon();
    if (startChanged) current = todayPeriod();
    closeSettings();
    render('today');
    showNotice('Đã gộp file sao lưu: ' + nCells + ' ô, ' + m.notes + ' ghi chú. Chạm để ẩn.');
  }

  /** "Nhập lại (thay toàn bộ)": đang ở nhóm đồng bộ thì đóng giờ mới cho mọi mục để các máy trong nhóm cũng được thay. */
  function stampAllAfterReplace(oldDoc) {
    var nd = docOf(state);
    var ck = {}, nk = {};
    [oldDoc, nd].forEach(function (d) {
      Object.keys(d.cells).forEach(function (day) { Object.keys(d.cells[day]).forEach(function (c) { ck[day + '|' + c] = 1; }); });
      Object.keys(d.mod.c).forEach(function (k) { ck[k] = 1; });
      Object.keys(d.notes).forEach(function (k) { nk[k] = 1; });
      Object.keys(d.mod.n).forEach(function (k) { nk[k] = 1; });
    });
    Object.keys(ck).forEach(function (k) { state.mod.c[k] = nextTs(Math.max(state.mod.c[k] || 0, oldDoc.mod.c[k] || 0)); });
    Object.keys(nk).forEach(function (k) { state.mod.n[k] = nextTs(Math.max(state.mod.n[k] || 0, oldDoc.mod.n[k] || 0)); });
    MOD_SET_KEYS.forEach(function (k) { state.mod.s[k] = nextTs(Math.max(state.mod.s[k] || 0, oldDoc.mod.s[k] || 0)); });
  }

  function importBackupText(text) {
    var b = readBackup(text);
    if (!b) return;
    var obj = b.obj, res = b.res;
    var sum = summarize(res.state);
    var inGroup = syncOn();
    var info = 'File sao lưu ngày ' + backupDate(obj) + ':\n' +
      '· ' + sum.days + ' ngày có chấm công' + (sum.days ? ' (từ ' + viDate(sum.from) + ' đến ' + viDate(sum.to) + ')' : '') + '\n' +
      '· ' + sum.notes + ' ghi chú' + (res.dropped ? '\n· ' + res.dropped + ' mục không đúng dạng sẽ bỏ qua' : '') + '\n\n' +
      'Thay TOÀN BỘ dữ liệu trên máy bằng file này?\n(Dữ liệu hiện tại được cất riêng trong máy trước khi thay.)' +
      (inGroup ? '\n\nMáy đang ở nhóm đồng bộ: các máy cùng nhóm cũng sẽ bị thay theo file này. Chỉ muốn thêm số còn thiếu thì bấm Huỷ và dùng "Gộp từ file sao lưu".' : '');
    if (!confirm(info)) return;
    if (!stashCurrent('truoc-nhap')) { alert('Máy hết chỗ để cất dữ liệu hiện tại, chưa nhập file.'); return; }
    var oldState = state;
    var oldDoc = clone(docOf(state));
    var oldTrial = dateKey(trialStartDate());
    var oldKey = typeof state.meta.licenseKey === 'string' ? state.meta.licenseKey : '';
    state = res.state;
    if (inGroup) stampAllAfterReplace(oldDoc);
    var start = earliestKey2(oldTrial, state.meta.trialStart);
    state.meta.trialStart = start && keyToDate(start) <= today ? start : oldTrial;
    if (oldKey && license.ok && state.meta.licenseKey !== oldKey) state.meta.licenseKey = oldKey;   // giữ khoá đang dùng
    saveBlocked = false;           // bản cũ đã cất riêng ở trên, giờ được phép ghi bản mới
    loadWarning = '';
    if (!saveState()) { state = oldState; return; }
    if (inGroup) syncSoon();
    current = todayPeriod();
    closeSettings();
    render('today');
    showNotice('Đã nhập xong file sao lưu: ' + sum.days + ' ngày chấm công, ' + sum.notes + ' ghi chú. Chạm để ẩn.');
    // kiểm lại khoá theo dữ liệu mới
    var k = typeof state.meta.licenseKey === 'string' ? state.meta.licenseKey : '';
    (k ? verifyKey(k) : Promise.resolve(null)).then(function (pk) { applyLicense(pk); });
  }

  var elFile = document.getElementById('fileInput');
  var fileMode = 'replace';            // 'replace' = Nhập lại (thay toàn bộ) · 'merge' = Gộp từ file
  function pickFile(mode) { fileMode = mode; elFile.click(); }
  elFile.addEventListener('change', function () {
    var f = elFile.files && elFile.files[0];
    elFile.value = '';
    if (!f) return;
    if (f.size > 5 * 1024 * 1024) { alert('File quá lớn, không phải file sao lưu của Sổ Công.'); return; }
    var mode = fileMode;
    var rd = new FileReader();
    rd.onload = function () { if (mode === 'merge') mergeBackupText(String(rd.result)); else importBackupText(String(rd.result)); };
    rd.onerror = function () { alert('Không đọc được file.'); };
    rd.readAsText(f, 'utf-8');
  });

  // --- Xuất Excel (.xlsx) kỳ đang xem: tự ghép file, không cần thư viện ngoài ---
  function crc32(bytes) {
    var t = crc32.t;
    if (!t) {
      t = crc32.t = new Uint32Array(256);
      for (var n = 0; n < 256; n++) {
        var c = n;
        for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
        t[n] = c >>> 0;
      }
    }
    var crc = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) crc = t[(crc ^ bytes[i]) & 0xFF] ^ (crc >>> 8);
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }
  /** Ghép file zip kiểu "store" (không nén). files: [{ name, text }] */
  function makeZip(files) {
    var parts = [], central = [], offset = 0;
    files.forEach(function (f) {
      var name = utf8Bytes(f.name), data = utf8Bytes(f.text), crc = crc32(data);
      var h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true);
      h.setUint16(8, 0, true); h.setUint16(10, 0, true); h.setUint16(12, 0x21, true);
      h.setUint32(14, crc, true); h.setUint32(18, data.length, true); h.setUint32(22, data.length, true);
      h.setUint16(26, name.length, true); h.setUint16(28, 0, true);
      parts.push(new Uint8Array(h.buffer), name, data);
      var c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true);
      c.setUint16(10, 0, true); c.setUint16(12, 0, true); c.setUint16(14, 0x21, true);
      c.setUint32(16, crc, true); c.setUint32(20, data.length, true); c.setUint32(24, data.length, true);
      c.setUint16(28, name.length, true); c.setUint16(30, 0, true); c.setUint16(32, 0, true);
      c.setUint16(34, 0, true); c.setUint16(36, 0, true); c.setUint32(38, 0, true); c.setUint32(42, offset, true);
      central.push(new Uint8Array(c.buffer), name);
      offset += 30 + name.length + data.length;
    });
    var cenSize = central.reduce(function (a, b) { return a + b.length; }, 0);
    var e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true);
    e.setUint32(12, cenSize, true); e.setUint32(16, offset, true);
    var all = parts.concat(central, [new Uint8Array(e.buffer)]);
    var out = new Uint8Array(all.reduce(function (a, b) { return a + b.length; }, 0));
    var pos = 0;
    all.forEach(function (b) { out.set(b, pos); pos += b.length; });
    return out;
  }
  function xmlEsc(v) { return stripCtl(v).replace(/[<>&"]/g, function (c) { return { '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;' }[c]; }); }
  function colLetter(i) { var s2 = ''; i++; while (i > 0) { var m = (i - 1) % 26; s2 = String.fromCharCode(65 + m) + s2; i = Math.floor((i - 1) / 26); } return s2; }

  function periodXlsx(p) {
    var st = state.settings;
    var days = periodDays(p.year, p.month, st.startDay);
    var autoFrom = autoFromDate();
    var rows = [];
    var first = days[0], last = days[days.length - 1];
    rows.push(['Sổ Công – Kỳ ' + p.month + '/' + p.year + ' (' + first.getDate() + '/' + (first.getMonth() + 1) + ' – ' +
      last.getDate() + '/' + (last.getMonth() + 1) + ')' + (st.name ? ' – ' + st.name : '')]);
    var head = ['Ngày', 'Thứ', 'Âm lịch'].concat(state.columns.map(function (c) { return c.name; }), ['Ngày lễ', 'Ghi chú']);
    rows.push(head);
    var totals = state.columns.map(function () { return 0; });
    var hasNum = state.columns.map(function () { return false; });
    days.forEach(function (d) {
      var key = dateKey(d);
      var l = window.SoCongLunar ? window.SoCongLunar.solarToLunar(d.getDate(), d.getMonth() + 1, d.getFullYear()) : null;
      var r = [d.getDate() + '/' + (d.getMonth() + 1) + '/' + d.getFullYear(), DOW[d.getDay()], l ? l.day + '/' + l.month : ''];
      state.columns.forEach(function (c, i) {
        var v = displayValue(key, d, c.id, autoFrom);
        var n = parseNumber(v);
        if (n !== null) { totals[i] += n; hasNum[i] = true; r.push(n); } else r.push(v);
      });
      r.push((holidayOf(d) || '') + (isPayday(d) ? (holidayOf(d) ? ' · ' : '') + 'Nhận lương' : ''));
      r.push(state.notes[key] || '');
      rows.push(r);
    });
    rows.push(['Tổng', '', ''].concat(totals.map(function (t, i) { return hasNum[i] ? Math.round(t * 100) / 100 : ''; }), ['', '']));

    var sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>';
    rows.forEach(function (r, ri) {
      sheet += '<row r="' + (ri + 1) + '">';
      r.forEach(function (v, ci) {
        var ref = colLetter(ci) + (ri + 1);
        var bold = ri === 0 || ri === 1 || ri === rows.length - 1 ? ' s="1"' : '';
        if (typeof v === 'number') sheet += '<c r="' + ref + '"' + bold + '><v>' + v + '</v></c>';
        else if (v !== '') sheet += '<c r="' + ref + '" t="inlineStr"' + bold + '><is><t xml:space="preserve">' + xmlEsc(v) + '</t></is></c>';
      });
      sheet += '</row>';
    });
    sheet += '</sheetData></worksheet>';
    var files = [
      { name: '[Content_Types].xml', text: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>' },
      { name: '_rels/.rels', text: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
      { name: 'xl/workbook.xml', text: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Ky ' + p.month + '-' + p.year + '" sheetId="1" r:id="rId1"/></sheets></workbook>' },
      { name: 'xl/_rels/workbook.xml.rels', text: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' },
      { name: 'xl/styles.xml', text: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>' },
      { name: 'xl/worksheets/sheet1.xml', text: sheet }
    ];
    return makeZip(files);
  }

  function doExportExcel() {
    var name = 'SoCong_Ky-' + current.month + '-' + current.year + '.xlsx';
    var msg = exportFile(name, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', periodXlsx(current));
    closeSettings();
    showNotice(msg ? msg + ' Chạm để ẩn.' : 'Không xuất được file Excel. Chạm để ẩn.');
  }

  function clearPeriod() {
    var days = periodDays(current.year, current.month, state.settings.startDay);
    var n = 0;
    days.forEach(function (d) { var k = dateKey(d); if (state.cells[k] || state.notes[k]) n++; });
    if (!n) { alert('Kỳ ' + current.month + '/' + current.year + ' chưa có số chấm tay hay ghi chú nào để xoá.'); return; }
    if (!confirm('Xoá hết số đã chấm và ghi chú của kỳ ' + current.month + '/' + current.year + ' (' + n + ' ngày)?\nCác ô tự điền (Full, Chủ nhật, ngày lễ) vẫn hiện như mặc định.')) return;
    if (!confirm('Chắc chắn xoá? Bản trước khi xoá được cất riêng trong máy.')) return;
    if (!stashCurrent('truoc-xoa')) { alert('Máy hết chỗ để cất bản cũ, chưa xoá.'); return; }
    days.forEach(function (d) {
      var k = dateKey(d);
      if (state.cells[k]) Object.keys(state.cells[k]).forEach(function (c) { markCell(k, c); });
      if (state.notes[k]) markNote(k);
      delete state.cells[k]; delete state.notes[k];
    });
    saveState();
    closeSettings();
    render('today');
    showNotice('Đã xoá số đã chấm và ghi chú của kỳ ' + current.month + '/' + current.year + '. Chạm để ẩn.');
  }

  // --- Hướng dẫn sử dụng ---
  var elHelp = document.getElementById('helpSheet');
  function helpOpen() { return isShown(elHelp); }
  function openHelp() { setShown(elSettings, false); setShown(elHelp, true); setShown(elBackdrop, true); elHelp.querySelector('.settings-body').scrollTop = 0; }
  function closeHelp() { setShown(elHelp, false); setShown(elBackdrop, false); }
  elHelp.addEventListener('click', function (e) {
    var b = e.target.closest('button');
    if (b && b.getAttribute('data-action') === 'close-help') closeHelp();
  });

  /** Sang kỳ mới mà chưa sao lưu từ đầu kỳ: nhắc một lần mỗi kỳ. */
  function remindBackup() {
    if (loadWarning || !Object.keys(state.cells).length) return;
    var tp = todayPeriod();
    var start = periodStart(tp.year, tp.month, state.settings.startDay);
    var tag = tp.year + '-' + tp.month;
    var last = state.meta.lastBackup ? keyToDate(state.meta.lastBackup) : null;
    if ((last && last >= start) || state.meta.backupReminded === tag) return;
    state.meta.backupReminded = tag;
    saveState();
    showNotice('Nhắc sao lưu: vào Cài đặt → "Sao lưu ra file" rồi gửi file vào Zalo hoặc Drive, phòng khi đổi máy hay lỡ gỡ app. Chạm để ẩn.');
  }

  // ---------- Khoá mở dùng trọn đời (1.2.0; bản 2 gắn mã máy từ 1.3.0) ----------
  var TRIAL_DAYS = 25;          // Sếp chốt 15/11: dùng thử 25 ngày
  var ZALO_SO = '0793366886';
  var ZALO_LINK = 'https://zalo.me/' + ZALO_SO;
  var MA_MAY_RE = /^[0-9A-HJKMNP-TV-Z]{8}$/;

  /** Mã máy: Android lấy từ máy (giữ khi cài lại), bản web tạo ngẫu nhiên một lần, lưu riêng (không theo file sao lưu). */
  var DEVICE_CODE = (function () {
    var A = window.SoCongAndroid;
    if (A && typeof A.deviceCode === 'function') {
      try { var c0 = String(A.deviceCode() || ''); if (MA_MAY_RE.test(c0)) return c0; } catch (e) { /* dùng cách web */ }
    }
    try { var s0 = localStorage.getItem('socong.may'); if (s0 && MA_MAY_RE.test(s0)) return s0; } catch (e) { /* tạo mới */ }
    var CH = '0123456789ABCDEFGHJKMNPQRSTVWXYZ', b = new Uint8Array(8), c = '';
    crypto.getRandomValues(b);
    for (var i = 0; i < 8; i++) c += CH[b[i] & 31];
    try { localStorage.setItem('socong.may', c); } catch (e) { /* không lưu được thì mỗi lần mở một mã */ }
    return c;
  })();
  function deviceCodePretty() { return DEVICE_CODE.slice(0, 4) + '-' + DEVICE_CODE.slice(4); }
  var REVOKED = [];             // mã khoá bị thu hồi (thêm vào đây ở bản sau nếu cần)
  var license = { ok: false, name: '', id: '', pending: false };

  function b64uToBytes(s) {
    s = String(s).replace(/-/g, '+').replace(/_/g, '/');
    while (s.length % 4) s += '=';
    var bin = atob(s), u = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    return u;
  }
  function parseKey(k) {
    var RE = /SC1\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]{86}(?![A-Za-z0-9_-])/;   // chữ ký luôn 86 ký tự
    k = String(k || '');
    var m = k.match(RE) || k.replace(/\s+/g, '').match(/SC1\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]{86}/);   // dán cả câu Zalo, hay khoá bị xuống dòng, vẫn tìm ra
    if (!m) return null;
    var p = m[0].split('.');
    try {
      var info = JSON.parse(new TextDecoder().decode(b64uToBytes(p[1])));
      var sig = b64uToBytes(p[2]);
      if (!isPlainObject(info) || (info.v !== 1 && info.v !== 2) || typeof info.n !== 'string' || typeof info.i !== 'string' || sig.length !== 64) return null;
      if (info.v === 2 && (typeof info.m !== 'string' || typeof info.e !== 'number')) return null;
      return { raw: m[0], msg: p[0] + '.' + p[1], sig: sig, info: info };
    } catch (e) { return null; }
  }
  /** Khoá đã nhập (lưu trong máy): đúng chữ ký, bản 2, đúng mã máy này. Hạn 1 giờ chỉ xét lúc nhập. */
  function verifyKey(k) {
    return verifySig(k).then(function (pk) {
      return pk && pk.info.v === 2 && pk.info.m === DEVICE_CODE ? pk : null;
    });
  }
  /** Kiểm chữ ký khoá. Trả Promise: thông tin khoá nếu chữ ký đúng, null nếu sai. */
  function verifySig(k) {
    var pk = parseKey(k);
    var K = window.SoCongKhoaCongKhai;
    if (!pk || !K || REVOKED.indexOf(pk.info.i) !== -1) return Promise.resolve(null);
    function viaAndroid() {
      var A = window.SoCongAndroid;
      if (!A || typeof A.verifySig !== 'function') return null;
      var bin = '';
      for (var i = 0; i < pk.sig.length; i++) bin += String.fromCharCode(pk.sig[i]);
      try { return A.verifySig(K.spki, pk.msg, btoa(bin)) ? pk : null; } catch (e) { return null; }
    }
    if (window.crypto && crypto.subtle) {
      return crypto.subtle.importKey('jwk', K.jwk, { name: 'ECDSA', namedCurve: 'P-256' }, false, ['verify'])
        .then(function (key) { return crypto.subtle.verify({ name: 'ECDSA', hash: 'SHA-256' }, key, pk.sig, utf8Bytes(pk.msg)); })
        .then(function (ok) { return ok ? pk : null; }, function () { return viaAndroid(); });
    }
    return Promise.resolve(viaAndroid());
  }

  var TRIAL_KEY = 'socong.thu';   // bản ghi riêng trong máy, không nằm trong file sao lưu
  function readTrialMark() {
    try { var v = localStorage.getItem(TRIAL_KEY); return v && DATE_KEY_RE.test(v) ? v : ''; } catch (e) { return ''; }
  }
  function earliestKey2(a, b) {
    var ok = function (k) { return typeof k === 'string' && DATE_KEY_RE.test(k); };
    if (ok(a) && ok(b)) return a < b ? a : b;
    return ok(a) ? a : (ok(b) ? b : '');
  }
  function trialStartDate() {
    var k = earliestKey2(state.meta.trialStart, readTrialMark());
    var d = k ? keyToDate(k) : today;
    return d > today ? today : d;      // ngày tương lai (sửa tay) tính như hôm nay
  }
  function trialDaysLeft() {
    var end = addDays(trialStartDate(), TRIAL_DAYS);
    return Math.max(0, Math.round((end - today) / 86400000));
  }
  function isLocked() { return !license.ok && trialDaysLeft() <= 0; }

  function licenseText() {
    if (license.ok) return 'Bản quyền: ' + license.name;
    if (license.pending) return 'Đang kiểm khoá…';
    var left = trialDaysLeft();
    return left > 0 ? 'Đang dùng thử: còn ' + left + ' ngày' : 'Hết thời gian dùng thử';
  }

  var elKey = document.getElementById('keySheet');
  var elKeyInput = document.getElementById('keyInput');
  var elKeyMsg = document.getElementById('keyMsg');
  function keyOpen() { return isShown(elKey); }
  function openKey(reason) {
    closePicker();
    if (settingsOpen()) setShown(elSettings, false);
    if (helpOpen()) setShown(elHelp, false);
    elKeyMsg.textContent = reason || '';
    elKeyInput.value = '';
    document.getElementById('keyDevice').textContent = deviceCodePretty();
    preloadQr();
    elKey.classList.toggle('licensed', license.ok);
    document.getElementById('keyTitle').textContent = license.ok ? 'Sổ Công · ' + licenseText() : 'Nhập khoá mở Sổ Công';
    setShown(elKey, true);
    setShown(elBackdrop, true);
  }
  function closeKey() {
    setShown(elKey, false);
    setShown(elBackdrop, false);
    if (document.activeElement === elKeyInput) elKeyInput.blur();
  }

  function applyLicense(pk) {
    license.ok = !!pk;
    license.name = pk ? pk.info.n : '';
    license.id = pk ? pk.info.i : '';
    var st = document.getElementById('licStatus');
    if (st) st.textContent = licenseText();
    var btn = document.getElementById('btnEnterKey');
    if (btn) btn.hidden = license.ok;
    showLockNotice();
  }
  function showLockNotice() {
    if (loadWarning) return;
    if (isLocked()) showNotice('Hết 25 ngày dùng thử. Vào Cài đặt → "Nhập khoá mở", nhắn Zalo ' + ZALO_SO + ' để xin khoá. Số đã chấm vẫn xem, sao lưu, xuất Excel được.', true);
    else if (noticeSticky && /dùng thử/.test(elNotice.textContent)) showNotice('');
  }

  /** Nhắn Zalo xin khoá: chép sẵn câu có mã máy rồi mở Zalo của Sếp. */
  function askKeyOnZalo() {
    var msg = 'Cho em xin khoá mở Sổ Công.' + (state.settings.name ? ' Tên: ' + state.settings.name + '.' : '') + ' Mã máy: ' + deviceCodePretty();
    // Chép chữ và mở Zalo NGAY trong lúc chạm (Safari iPhone chặn nếu làm sau một bước chờ)
    copyText(msg);
    elKeyMsg.textContent = 'Câu xin khoá đã chép sẵn, mở Zalo và dán để gửi: "' + msg + '". Có khoá thì nhập trong vòng 1 giờ.';
    var A = window.SoCongAndroid;
    if (A && typeof A.openZalo === 'function') { try { A.openZalo(); return; } catch (e) { /* mở kiểu web */ } }
    var w = window.open(ZALO_LINK, '_blank');
    if (!w) location.href = ZALO_LINK;      // bị chặn cửa sổ mới thì đi thẳng
  }
  /** Bản web: tải sẵn ảnh QR khi mở bảng khoá, để lúc chạm "Lưu ảnh" chia sẻ được ngay. */
  var qrBytes = null;
  function preloadQr() {
    if (qrBytes || (window.SoCongAndroid && window.SoCongAndroid.saveQr)) return;
    try {
      fetch('icons/qr-momo.png').then(function (r) { return r.ok ? r.arrayBuffer() : null; })
        .then(function (b) { if (b) qrBytes = new Uint8Array(b); }, function () { /* không sao */ });
    } catch (e) { /* không sao */ }
  }
  /** Lưu / gửi ảnh QR MoMo mời cà phê. */
  function saveQr() {
    var A = window.SoCongAndroid;
    if (A && typeof A.saveQr === 'function') {
      var where = '';
      try { where = A.saveQr() || ''; } catch (e) { where = ''; }
      elKeyMsg.textContent = where ? 'Đã lưu ảnh QR vào ' + where + '. Mở MoMo → Quét mã → chọn ảnh.' : 'Đã mở bảng gửi ảnh QR.';
      return;
    }
    var tip = ' Trên iPhone cũng có thể chạm giữ ảnh QR → Lưu vào Ảnh, rồi mở MoMo → Quét mã → chọn ảnh.';
    if (!qrBytes) { elKeyMsg.textContent = 'Chạm giữ ảnh QR → Lưu vào Ảnh (hoặc chụp màn hình), rồi mở MoMo → Quét mã → chọn ảnh.'; preloadQr(); return; }
    var msg = exportFile('SoCong_MoMo_QR.png', 'image/png', qrBytes);
    elKeyMsg.textContent = (msg || 'Không lưu được ảnh.') + tip;
  }

  /** Thử mở khoá bằng chuỗi khoá. fromLink: đến từ link #khoa= */
  function tryUnlock(text, fromLink) {
    elKeyMsg.textContent = 'Đang kiểm khoá…';
    return verifySig(text).then(function (pk) {
      var why = '';
      if (!pk) why = 'Khoá không đúng. Hãy dán đúng nguyên tin nhắn có đoạn khoá bắt đầu bằng SC1.';
      else if (pk.info.v !== 2) why = 'Đây là khoá kiểu cũ. Hãy nhắn Zalo xin khoá mới.';
      else if (pk.info.m !== DEVICE_CODE) why = 'Khoá này cấp cho máy khác (' + pk.info.m.slice(0, 4) + '-' + pk.info.m.slice(4) + '). Mã máy này là ' + deviceCodePretty() + '.';
      else if (Date.now() / 1000 > pk.info.e) why = 'Khoá đã quá 1 giờ kể từ lúc cấp nên hết hiệu lực. Hãy nhắn Zalo xin khoá mới.';
      if (why) {
        if (fromLink) showNotice(why + ' Chạm để ẩn.');
        else elKeyMsg.textContent = why;
        return false;
      }
      elKeyMsg.textContent = '';
      state.meta.licenseKey = pk.raw;
      saveState();
      applyLicense(pk);
      if (keyOpen()) closeKey();
      showNotice('Đã mở khoá Sổ Công cho ' + pk.info.n + '. Dùng trọn đời. Chạm để ẩn.');
      return true;
    });
  }

  document.getElementById('keySheet').addEventListener('click', function (e) {
    var b = e.target.closest('button');
    if (!b) return;
    var a = b.getAttribute('data-action');
    if (a === 'key-ok') tryUnlock(elKeyInput.value, false);
    else if (a === 'key-zalo') askKeyOnZalo();
    else if (a === 'save-qr') saveQr();
    else if (a === 'key-close') closeKey();
  });

  // Ngày bắt đầu dùng thử: lần đầu mở bản có khoá. Ghi cả vào dữ liệu và vào bản ghi riêng trong máy.
  (function initTrial() {
    var start = dateKey(trialStartDate());
    try { if (readTrialMark() !== start) localStorage.setItem(TRIAL_KEY, start); } catch (e) { /* không sao */ }
    if (state.meta.trialStart !== start && !saveBlocked) {
      state.meta.trialStart = start;
      saveState();
    }
  })();

  /** Kiểm khoá đã lưu (mỗi lần mở app) và link #khoa= (bản web). */
  function initLicense() {
    var linkKey = '';
    var h = location.hash || '';
    if (h.indexOf('#khoa=') === 0) {
      try { linkKey = decodeURIComponent(h.slice(6)); } catch (e) { linkKey = h.slice(6); }
      try { history.replaceState(null, '', location.pathname + location.search); } catch (e) { /* không sao */ }
    }
    var saved = typeof state.meta.licenseKey === 'string' ? state.meta.licenseKey : '';
    license.pending = !!saved;
    var first = saved ? verifyKey(saved) : Promise.resolve(null);
    return first.then(function (pk) {
      license.pending = false;
      applyLicense(pk);
      if (linkKey && !license.ok) return tryUnlock(linkKey, true);
      if (!license.ok && !isLocked()) {
        var left = trialDaysLeft();
        if (left <= 5 && state.meta.trialWarned !== dateKey(today) && !loadWarning) {
          state.meta.trialWarned = dateKey(today);
          saveState();
          showNotice('Còn ' + left + ' ngày dùng thử. Có khoá mở thì vào Cài đặt → "Nhập khoá mở". Chạm để ẩn.');
        }
      }
    });
  }

  // ---------- Cài đặt (lát 3) ----------
  var elSettings = document.getElementById('settings');
  var elStart = document.getElementById('setStartDay');
  var elPayDay = document.getElementById('setPayDay');
  var elFontVal = document.getElementById('fontValue');
  var elStartHint = document.getElementById('startHint');
  var TOGGLES = ['fitScreen', 'showLunar', 'hideEmpty', 'weekTotals', 'showName', 'autoFull', 'feelVibrate', 'feelSound'];

  function settingsOpen() { return isShown(elSettings); }

  (function fillDaySelects() {
    var o = '';
    for (var i = 1; i <= 31; i++) o += '<option value="' + i + '">' + i + '</option>';
    elStart.innerHTML = o;
    elPayDay.innerHTML = o;
  })();

  var elCols = document.getElementById('colList');
  var PALETTE = ['var(--c1)', 'var(--c2)', 'var(--c3)', 'var(--c4)', 'var(--c5)', 'var(--c6)', 'var(--c7)', 'var(--c8)', 'var(--c9)'];

  function renderColEditor() {
    var h = '';
    state.columns.forEach(function (c, i) {
      h += '<div class="col-row" data-i="' + i + '">' +
        '<button type="button" class="swatch" data-action="col-color" style="background:' + esc(c.color) + '" aria-label="Đổi màu cột"></button>' +
        '<input type="text" class="col-name" maxlength="8" value="' + esc(c.name) + '" aria-label="Tên cột">' +
        '<button type="button" class="mini" data-action="col-up" aria-label="Lên"' + (i === 0 ? ' disabled' : '') + '>▲</button>' +
        '<button type="button" class="mini" data-action="col-down" aria-label="Xuống"' + (i === state.columns.length - 1 ? ' disabled' : '') + '>▼</button>' +
        '<button type="button" class="mini danger" data-action="col-del" aria-label="Bỏ cột"' + (state.columns.length <= 1 ? ' disabled' : '') + '>✕</button>' +
        '</div>';
    });
    var removed = Array.isArray(state.meta.removedColumns) ? state.meta.removedColumns : [];
    var full = state.columns.length >= MAX_COLS;
    removed.forEach(function (c, r) {
      if (!c || typeof c.id !== 'string' || !COL_ID_RE.test(c.id)) return;
      h += '<div class="col-row removed">' +
        '<span class="swatch" style="background:' + esc(COLOR_RE.test(c.color) ? c.color : 'var(--c9)') + '"></span>' +
        '<span class="col-removed-name">' + esc(c.name || c.id) + ' <small>(đã bỏ)</small></span>' +
        '<button type="button" class="mini wide" data-action="col-restore" data-r="' + r + '"' + (full ? ' disabled' : '') + '>Lấy lại</button>' +
        '</div>';
    });
    elCols.innerHTML = h;
    document.getElementById('btnAddCol').disabled = full;
  }

  function colsChanged() {
    markSet('cols');
    saveState();
    renderColEditor();
    render();
  }

  elCols.addEventListener('input', function (e) {
    var inp = e.target.closest('input.col-name');
    if (!inp || isLocked()) return;
    var i = +inp.closest('.col-row').getAttribute('data-i');
    var name = stripCtl(inp.value).trim().slice(0, 8);
    if (!name || name === state.columns[i].name) return;
    state.columns[i].name = name;
    markSet('cols');
    saveState();
  });
  elCols.addEventListener('change', function (e) {
    var inp = e.target.closest('input.col-name');
    if (!inp) return;
    var i = +inp.closest('.col-row').getAttribute('data-i');
    if (isLocked()) { inp.value = state.columns[i].name; return openKey('Hết ' + TRIAL_DAYS + ' ngày dùng thử. Nhập khoá mở để sửa cột.'); }
    var name = stripCtl(inp.value).trim().slice(0, 8);
    if (!name) { inp.value = state.columns[i].name; return; }
    state.columns[i].name = name;
    colsChanged();
  });
  elCols.addEventListener('click', function (e) {
    var btn = e.target.closest('button');
    if (!btn) return;
    if (isLocked()) return openKey('Hết ' + TRIAL_DAYS + ' ngày dùng thử. Nhập khoá mở để sửa cột.');
    var cols = state.columns;
    var a = btn.getAttribute('data-action');
    if (a === 'col-restore') {
      var list = Array.isArray(state.meta.removedColumns) ? state.meta.removedColumns : [];
      var r = +btn.getAttribute('data-r');
      var col = list[r];
      if (!col || cols.length >= MAX_COLS) return;
      list.splice(r, 1);
      if (!cols.some(function (c) { return c.id === col.id; })) {
        cols.push({ id: col.id, name: col.name || col.id, color: COLOR_RE.test(col.color) ? col.color : 'var(--c9)' });
      }
      return colsChanged();
    }
    var i = +btn.closest('.col-row').getAttribute('data-i');
    if (a === 'col-color') {
      var idx = PALETTE.indexOf(cols[i].color);
      cols[i].color = PALETTE[(idx + 1) % PALETTE.length];
    } else if (a === 'col-up' && i > 0) {
      cols.splice(i - 1, 0, cols.splice(i, 1)[0]);
    } else if (a === 'col-down' && i < cols.length - 1) {
      cols.splice(i + 1, 0, cols.splice(i, 1)[0]);
    } else if (a === 'col-del' && cols.length > 1) {
      if (!confirm('Bỏ cột "' + cols[i].name + '" khỏi bảng?\nSố đã chấm ở cột này vẫn cất trong máy, không bị xoá.')) return;
      if (!Array.isArray(state.meta.removedColumns)) state.meta.removedColumns = [];
      state.meta.removedColumns.push(cols[i]);
      cols.splice(i, 1);
    } else return;
    colsChanged();
  });
  document.getElementById('btnAddCol').addEventListener('click', function () {
    if (isLocked()) return openKey('Hết ' + TRIAL_DAYS + ' ngày dùng thử. Nhập khoá mở để thêm cột.');
    if (state.columns.length >= MAX_COLS) return;
    var removed = Array.isArray(state.meta.removedColumns) ? state.meta.removedColumns : [];
    var id;
    do { id = 'c' + Date.now().toString(36) + Math.floor(Math.random() * 36).toString(36); } while (
      state.columns.some(function (c) { return c.id === id; }) || removed.some(function (c) { return c && c.id === id; }));
    state.columns.push({ id: id, name: 'Cột ' + (state.columns.length + 1), color: PALETTE[state.columns.length % PALETTE.length] });
    colsChanged();
  });

  function syncSettingsForm() {
    renderColEditor();
    updateSyncUi();
    applyLicense(license.ok ? { info: { n: license.name, i: license.id } } : null);
    var btnOld = elSettings.querySelector('[data-action="import-old"]');
    if (btnOld) btnOld.hidden = !window.SoCongAnhMau;
    var st = state.settings;
    TOGGLES.forEach(function (k) { document.getElementById('set_' + k).checked = !!st[k]; });
    elStart.value = String(st.startDay);
    elPayDay.value = String(st.payDay);
    elFontVal.textContent = Math.round(st.fontScale * 100) + '%';
    elStartHint.hidden = st.startDay < 29;
  }

  function openSettings(openSync) {
    closePicker();
    syncSettingsForm();
    // 1.4.7: mở sẵn mục Đồng bộ khi vừa làm việc đồng bộ xong, hoặc đồng bộ đang báo lỗi (để thấy nút Nhập mật khẩu mới)
    var fold = document.getElementById('syncFold');
    if (fold && (openSync === true || (syncCfg && syncCfg.err))) fold.open = true;
    setShown(elSettings, true);
    setShown(elBackdrop, true);
  }
  function closeSettings() {
    if (!settingsOpen()) return;
    setShown(elSettings, false);
    setShown(elBackdrop, false);
    if (document.activeElement && document.activeElement.blur) document.activeElement.blur();
    render();
  }

  function changeSetting(k, v) {
    var wasOnToday = samePeriod(current, todayPeriod());
    if ((k === 'startDay' || k === 'payDay') && state.settings[k] !== v) markSet(k);
    state.settings[k] = v;
    saveState();
    if (k === 'startDay' && wasOnToday) current = todayPeriod();
    syncSettingsForm();
    render(k === 'startDay' ? 'today' : undefined);
  }

  document.getElementById('btnSettings').addEventListener('click', openSettings);
  elSettings.addEventListener('change', function (e) {
    var t = e.target;
    if (t.type === 'checkbox' && t.id.indexOf('set_') === 0) changeSetting(t.id.slice(4), t.checked);
    else if (t === elStart) changeSetting('startDay', validDay(t.value, 21));
    else if (t === elPayDay) changeSetting('payDay', validDay(t.value, 21));
  });
  elSettings.addEventListener('click', function (e) {
    var btn = e.target.closest('button');
    if (!btn) return;
    var a = btn.getAttribute('data-action');
    if (a === 'close-settings') closeSettings();
    else if (a === 'enter-key') openKey('');
    else if (a === 'coffee') openKey('');
    else if (a === 'backup') doBackup();
    else if (a === 'restore') pickFile('replace');
    else if (a === 'merge-file') pickFile('merge');
    else if (a === 'export-excel') doExportExcel();
    else if (a === 'clear-period') clearPeriod();
    else if (a === 'help') openHelp();
    else if (a === 'sync-create') openSyncSheet('create');
    else if (a === 'sync-join') openSyncSheet('join');
    else if (a === 'sync-now') syncNow(true);
    else if (a === 'sync-copy') copyGroup(false);
    else if (a === 'sync-leave') leaveGroup();
    else if (a === 'sync-repw') openAdminSheet('repw');
    else if (a === 'sync-showpw') openAdminSheet('showpw');
    else if (a === 'sync-changepw') openAdminSheet('changepw');
    else if (a === 'sync-copyadm') copyAdmin();
    else if (a === 'sync-claim') claimAdmin();
    else if (a === 'sync-admcode') openAdminSheet('admcode');
    else if (a === 'copy-notes') {
      var txt = notesText(current);
      if (!txt) { alert('Kỳ ' + current.month + '/' + current.year + ' chưa có ghi chú nào.'); return; }
      copyText(txt).then(function (ok) {
        closeSettings();
        showNotice(ok ? 'Đã chép ghi chú kỳ ' + current.month + '/' + current.year + '. Mở Zalo và dán ra. Chạm để ẩn.'
          : 'Máy không cho chép tự động. Giữ ngón tay lên đoạn dưới để chép (dòng này tự ẩn khi đổi kỳ):\n' + txt, !ok);
      });
    }
    else if (a === 'import-old') {
      var src = window.SoCongAnhMau;
      if (!src) return;
      if (!confirm('Nạp số đã chấm ở app cũ (' + src.label + ') vào Sổ Công?\nChỉ điền vào ô còn trống, không đè số đã chấm.')) return;
      var n = importOldData();
      if (n < 0) return;
      closeSettings();
      render('today');
      showNotice(n ? 'Đã nạp ' + n + ' ô từ app cũ. Lưu ý: app cũ tính ngày 21/10 vào cả kỳ 10 và kỳ 11, Sổ Công chỉ tính vào kỳ 11, nên Tổng kỳ 10 thấp hơn ảnh cũ đúng bằng số của ngày 21/10. Chạm để ẩn.'
        : 'Số từ app cũ đã có đủ trên máy, không cần nạp thêm. Chạm để ẩn.');
    }
    else if (a === 'font-down') changeSetting('fontScale', Math.max(0.8, Math.round((state.settings.fontScale - 0.1) * 100) / 100));
    else if (a === 'font-up') changeSetting('fontScale', Math.min(1.4, Math.round((state.settings.fontScale + 0.1) * 100) / 100));
    else if (a === 'reset') {
      if (!confirm('Đưa mọi cài đặt về mặc định? Số đã chấm vẫn giữ nguyên.')) return;
      var keepName = state.settings.name;
      var extra = {};
      Object.keys(state.settings).forEach(function (k) { if (!(k in DEFAULT_SETTINGS)) extra[k] = state.settings[k]; });
      var oldS = state.settings;
      state.settings = Object.assign(clone(DEFAULT_SETTINGS), extra, { name: keepName });
      ['startDay', 'payDay'].forEach(function (k) { if (oldS[k] !== state.settings[k]) markSet(k); });
      saveState();
      current = todayPeriod();
      syncSettingsForm();
      render('today');
    }
  });

  var nameTimer = null;
  elName.addEventListener('input', function () {
    clearTimeout(nameTimer);
    nameTimer = setTimeout(function () {
      state.settings.name = stripCtl(elName.value).slice(0, 40);
      saveState();
    }, 400);
  });

  // ---------- 1.4.8: máy cá nhân cài từ trước tự chuyển sang 3 cột (một lần) ----------
  // Chỉ khi: không ở nhóm đồng bộ, bộ cột y nguyên mặc định (Q, P, Full, Q-, P-, chưa đổi tên, chưa bỏ cột),
  // cột P và P- chưa có số nào, chưa từng chuyển. Cột P, P- đưa vào "cột đã bỏ" (Lấy lại được), không xoá gì.
  function shortName(v) { v = String(v || '').replace(/\s+/g, ' ').trim(); return (v.length > 8 ? v.split(' ').pop() : v).slice(0, 8); }
  function migrateTo3() {
    if (syncCfg || saveBlocked || state.meta.cols3) return false;
    var want = [['q', 'Q'], ['p', 'P'], ['full', 'Full'], ['qm', 'Q-'], ['pm', 'P-']];
    var cols = state.columns;
    if (cols.length !== 5) return false;
    for (var i = 0; i < 5; i++) if (cols[i].id !== want[i][0] || cols[i].name !== want[i][1]) return false;
    if (Array.isArray(state.meta.removedColumns) && state.meta.removedColumns.length) return false;
    for (var k in state.cells) {
      var row = state.cells[k];
      if ((row.p !== undefined && row.p !== '') || (row.pm !== undefined && row.pm !== '')) return false;
    }
    var nm = shortName(state.settings.name);
    state.columns = [cols[0], cols[2], cols[3]];
    state.columns[0].name = nm || 'Tôi';
    state.columns[2].name = 'Nghỉ';
    state.meta.removedColumns = [cols[1], cols[4]];
    state.meta.cols3 = Date.now();
    saveState();
    return true;
  }

  // ---------- 1.4.4: hỏi tên lần đầu ----------
  var elNameSheet = document.getElementById('nameSheet');
  var elFirstName = document.getElementById('firstNameInput');
  function nameOpen() { return isShown(elNameSheet); }
  function openName() {
    elFirstName.value = '';
    setShown(elNameSheet, true);
    setShown(elBackdrop, true);
  }
  function closeName(save) {
    var v = stripCtl(elFirstName.value || '').replace(/\s+/g, ' ').trim();
    if (save && v) {
      // Không đánh dấu giờ sửa bộ cột: máy này vào nhóm đồng bộ sau thì bộ cột của nhóm vẫn thắng.
      var col = state.columns.filter(function (c) { return c.id === 'q'; })[0];
      // Tên dài quá 8 chữ thì lấy tên gọi (chữ cuối), ví dụ "Nguyễn Văn Hùng" → "Hùng"
      if (col) col.name = shortName(v);
      state.settings.name = v.slice(0, 40);
      saveState();
      render();
    }
    if (document.activeElement === elFirstName) elFirstName.blur();
    setShown(elNameSheet, false);
    setShown(elBackdrop, false);
  }
  document.getElementById('nameOk').addEventListener('click', function () { closeName(true); });
  document.getElementById('nameSkip').addEventListener('click', function () { closeName(false); });
  elFirstName.addEventListener('keydown', function (e) { if (e.key === 'Enter') closeName(true); });

  elBackdrop.addEventListener('click', function () {
    if (nameOpen()) closeName(false);
    else if (syncSheetOpen()) closeSyncSheet();
    else if (keyOpen()) closeKey();
    else if (editing) closePicker();
    else if (helpOpen()) closeHelp();
    else if (dayOpen()) closeDay();
    else if (settingsOpen()) closeSettings();
  });
  document.addEventListener('keydown', function (e) {
    if (e.key !== 'Escape') return;
    if (syncSheetOpen()) closeSyncSheet();
    else if (editing) closePicker();
    else if (helpOpen()) closeHelp();
    else if (dayOpen()) closeDay();
    else if (settingsOpen()) closeSettings();
  });

  // ---------- Sang ngày mới ----------
  function checkNewDay() {
    var now = startOfToday();
    if (now.getTime() === today.getTime()) return;
    var wasOnTodayPeriod = samePeriod(current, todayPeriod());
    today = now;
    if (wasOnTodayPeriod) current = todayPeriod();
    if (!editing && !dayOpen()) render(wasOnTodayPeriod ? 'today' : undefined);
  }
  document.addEventListener('visibilitychange', function () { if (!document.hidden) checkNewDay(); });
  setInterval(checkNewDay, 60 * 1000);

  // Nút / cử chỉ Quay lại của Android. Trả true nếu app đã xử lý (không thoát app).
  window.SoCongBack = function () {
    if (nameOpen()) { closeName(false); return true; }
    if (editing) { closePicker(); return true; }
    if (syncSheetOpen()) { closeSyncSheet(); return true; }
    if (keyOpen()) { closeKey(); return true; }
    if (helpOpen()) { closeHelp(); return true; }
    if (dayOpen()) { closeDay(); return true; }
    if (settingsOpen()) { closeSettings(); return true; }
    if (!samePeriod(current, todayPeriod())) { goTo(todayPeriod()); return true; }
    return false;
  };

  // ---------- Đồng bộ nhiều máy (1.4.0, lát 2–3) ----------
  // Dữ liệu được khoá bằng AES-GCM ngay trên máy; khoá suy từ mật khẩu nhóm bằng PBKDF2. Máy chủ chỉ thấy chuỗi đã mã hoá.
  // Cấu hình nhóm lưu riêng ở "socong.dongbo" (không nằm trong file sao lưu): mã nhóm, khoá AES, thẻ vào nhóm, rev, giờ.
  var SYNC_KEY = 'socong.dongbo';
  var SYNC_CH = '0123456789ABCDEFGHJKMNPQRSTVWXYZ';
  var GROUP_RE = /^[0-9A-HJKMNP-TV-Z]{16}$/;
  var PBKDF2_ITER = 250000;
  var SYNC_DELAY = 4000;              // gửi sau lần sửa cuối 4 giây
  var webSameSite = location.protocol === 'https:' || location.hostname === 'localhost' || location.hostname === '127.0.0.1';
  var API_URL = webSameSite ? 'api/dongbo.php' : 'https://quyenbds.com/socong/api/dongbo.php';
  var SYNC_ERR = {
    mang: 'Chưa có mạng hoặc không gọi được máy chủ. Có mạng app tự gửi.',
    'khong-vao-duoc': 'Sai mã nhóm hoặc mật khẩu, hoặc mật khẩu nhóm vừa được đổi: vào Cài đặt → Đồng bộ nhiều máy → Nhập mật khẩu mới.',
    'qua-nhieu': 'Gọi máy chủ quá nhiều lần, thử lại sau ít phút.',
    'qua-lon': 'Dữ liệu nhóm vượt giới hạn 1 MB.',
    'xung-dot': 'Máy khác đang gửi cùng lúc, app sẽ tự thử lại sau ít giây.',
    'het-cho': 'Máy chủ đồng bộ đang hết chỗ, báo tác giả giúp.',
    'giai-ma': 'Không mở được dữ liệu nhóm bằng mật khẩu này.',
    'ban-moi': 'Dữ liệu nhóm do bản Sổ Công mới hơn gửi lên. Hãy cập nhật app.',
    'khong-ma-hoa': 'Máy này không hỗ trợ mã hoá nên chưa đồng bộ được.',
    'may-chu': 'Máy chủ đồng bộ đang trục trặc, thử lại sau.'
  };
  function syncErrText(e) { var c = e && e.message; return SYNC_ERR[c] || SYNC_ERR['may-chu']; }

  function loadSyncCfg() {
    try {
      var c = JSON.parse(localStorage.getItem(SYNC_KEY) || 'null');
      if (isPlainObject(c) && GROUP_RE.test(c.g) && /^[0-9a-f]{64}$/.test(c.t) && typeof c.k === 'string') {
        c.rev = typeof c.rev === 'number' ? c.rev : 0;
        return c;
      }
    } catch (e) { /* coi như chưa vào nhóm */ }
    return null;
  }
  var syncCfg = loadSyncCfg();
  function saveSyncCfg() {
    try {
      if (syncCfg) localStorage.setItem(SYNC_KEY, JSON.stringify(syncCfg)); else localStorage.removeItem(SYNC_KEY);
    } catch (e) { /* lần sau đồng bộ lại từ đầu, không mất gì */ }
  }
  function syncOn() { return !!syncCfg; }
  function groupPretty(g) { return g.replace(/(.{4})(?=.)/g, '$1-'); }
  /** Mã nhóm gõ tay: bỏ gạch, khoảng trắng; chữ thường thành hoa; O→0, I/L→1 cho khỏi nhầm. */
  function normGroup(s) {
    var g = String(s || '').toUpperCase().replace(/[^0-9A-Z]/g, '').replace(/O/g, '0').replace(/[IL]/g, '1');
    return GROUP_RE.test(g) ? g : '';
  }
  function randomGroup() {
    var b = new Uint8Array(16), g = '';
    crypto.getRandomValues(b);
    for (var i = 0; i < 16; i++) g += SYNC_CH[b[i] & 31];
    return g;
  }
  function hex(bytes) { var h = ''; for (var i = 0; i < bytes.length; i++) h += (bytes[i] < 16 ? '0' : '') + bytes[i].toString(16); return h; }
  function hasCrypto() { return !!(window.crypto && crypto.subtle && crypto.getRandomValues); }

  /** Từ mật khẩu + mã nhóm suy ra 512 bit: nửa đầu là khoá AES (không rời máy), nửa sau là thẻ vào nhóm. */
  function deriveGroup(g, pw) {
    if (!hasCrypto()) return Promise.reject(new Error('khong-ma-hoa'));
    return crypto.subtle.importKey('raw', utf8Bytes(pw), 'PBKDF2', false, ['deriveBits']).then(function (k) {
      return crypto.subtle.deriveBits({ name: 'PBKDF2', hash: 'SHA-256', salt: utf8Bytes('SoCong-dongbo-v1|' + g), iterations: PBKDF2_ITER }, k, 512);
    }).then(function (bits) {
      var b = new Uint8Array(bits);
      return { k: toBase64(b.subarray(0, 32)), t: hex(b.subarray(32, 64)) };
    });
  }
  var aesCache = { k: '', key: null };
  function aesKey(cfg) {
    if (aesCache.k === cfg.k && aesCache.key) return Promise.resolve(aesCache.key);
    return crypto.subtle.importKey('raw', b64uToBytes(cfg.k), { name: 'AES-GCM' }, false, ['encrypt', 'decrypt']).then(function (key) {
      aesCache = { k: cfg.k, key: key };
      return key;
    });
  }
  /** Khoá dữ liệu: 12 byte ngẫu nhiên + AES-GCM, gắn mã nhóm làm dữ liệu kèm. JSON lấy ngay lúc gọi (trước mọi bước chờ). */
  function encryptDoc(cfg, doc) {
    var plain = utf8Bytes(JSON.stringify({
      v: 1, cells: doc.cells, notes: doc.notes, columns: doc.cols.columns, removed: doc.cols.removed,
      startDay: doc.startDay, payDay: doc.payDay, mod: doc.mod
    }));
    var iv = new Uint8Array(12);
    crypto.getRandomValues(iv);
    return aesKey(cfg).then(function (key) {
      return crypto.subtle.encrypt({ name: 'AES-GCM', iv: iv, additionalData: utf8Bytes('SoCong|' + cfg.g) }, key, plain);
    }).then(function (ct) {
      var c = new Uint8Array(ct), all = new Uint8Array(12 + c.length);
      all.set(iv, 0);
      all.set(c, 12);
      return toBase64(all);
    });
  }
  function decryptDoc(cfg, b64) {
    var all;
    try { all = b64uToBytes(b64); } catch (e) { return Promise.reject(new Error('giai-ma')); }
    if (all.length < 29) return Promise.reject(new Error('giai-ma'));
    return aesKey(cfg).then(function (key) {
      return crypto.subtle.decrypt({ name: 'AES-GCM', iv: all.subarray(0, 12), additionalData: utf8Bytes('SoCong|' + cfg.g) }, key, all.subarray(12));
    }).then(function (plain) {
      var o = JSON.parse(new TextDecoder().decode(plain));
      if (!isPlainObject(o)) throw new Error('giai-ma');
      if (typeof o.v === 'number' && o.v > 1) throw new Error('ban-moi');
      return docFromNet(o);
    }, function () { throw new Error('giai-ma'); });
  }
  /** Dữ liệu nhận về vẫn kiểm từng trường như dữ liệu trong máy. */
  function docFromNet(o) {
    var removed = (Array.isArray(o.removed) ? o.removed : []).filter(function (c) {
      return isPlainObject(c) && typeof c.id === 'string' && COL_ID_RE.test(c.id);
    }).map(function (c) {
      return { id: c.id, name: typeof c.name === 'string' ? stripCtl(c.name).slice(0, 12) : c.id, color: COLOR_RE.test(c.color) ? c.color : 'var(--c9)' };
    });
    var res = sanitizeState({
      version: DATA_VERSION, columns: o.columns, cells: o.cells, notes: o.notes,
      settings: { startDay: o.startDay, payDay: o.payDay }, meta: { removedColumns: removed }, mod: o.mod
    });
    return docOf(res.state);
  }

  // Gọi API: app Android qua cầu nối Java (địa chỉ cố định trong Java), bản web bằng fetch cùng nguồn quyenbds.com.
  var apiSeq = 0, apiWait = {};
  window.SoCongApiTraLoi = function (id, text, status) {
    var f = apiWait[id];
    if (f) { delete apiWait[id]; f(String(text || ''), Number(status) || 0); }
  };
  function parseApi(text, status) {
    var j = null;
    try { j = JSON.parse(text); } catch (e) { j = null; }
    if (!isPlainObject(j)) throw new Error(status ? 'may-chu' : 'mang');
    return j;
  }
  function api(body) {
    var payload = JSON.stringify(body);
    var A = window.SoCongAndroid;
    if (A && typeof A.goiApi === 'function') {
      return new Promise(function (resolve, reject) {
        var id = 'a' + (++apiSeq);
        var timer = setTimeout(function () { delete apiWait[id]; reject(new Error('mang')); }, 30000);
        apiWait[id] = function (text, status) {
          clearTimeout(timer);
          try { resolve(parseApi(text, status)); } catch (e) { reject(e); }
        };
        try { A.goiApi(id, payload); } catch (e) { clearTimeout(timer); delete apiWait[id]; reject(new Error('mang')); }
      });
    }
    if (typeof fetch !== 'function') return Promise.reject(new Error('mang'));
    return fetch(API_URL, {
      method: 'POST', headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, body: payload,
      cache: 'no-store', credentials: 'omit'
    }).then(function (r) {
      return r.text().then(function (t) { return parseApi(t, r.status); });
    }, function () { throw new Error('mang'); });
  }
  function apiOk(r) { if (!r.ok) throw new Error(r.loi || 'may-chu'); return r; }

  var syncBusy = null, syncAgain = false, syncTimer = null, syncDirty = false, syncWarned = '', syncFails = 0;
  // 1.4.1: app đang hiện thì hỏi nhẹ số lần sửa (rev) của nhóm: 20 giây; không ai sửa quá 5 phút thì 60 giây; ẩn app thì dừng
  var POLL_FAST = 20000, POLL_SLOW = 60000, POLL_IDLE = 300000;
  var pollTimer = null, lastActivity = Date.now();
  function schedulePoll(ms) {
    clearTimeout(pollTimer);
    pollTimer = null;
    if (!syncCfg || document.hidden) return;
    if (ms === undefined) ms = (syncCfg.err || Date.now() - lastActivity > POLL_IDLE) ? POLL_SLOW : POLL_FAST;   // đang lỗi thì cũng hỏi chậm
    pollTimer = setTimeout(pollRev, ms);
  }
  function pollRev() {
    pollTimer = null;
    if (!syncCfg || document.hidden) return;
    if (syncBusy || syncTimer || syncDirty) return syncDirty && !syncBusy && !syncTimer ? syncNow(false) : schedulePoll();
    var cfg = syncCfg;
    api({ a: 'rev', g: cfg.g, t: cfg.t }).then(function (r) {
      if (syncCfg !== cfg) return;
      if (r.ok && r.rev !== cfg.rev) { lastActivity = Date.now(); return syncNow(false); }
      if (!r.ok && r.loi === 'khong-vao-duoc' && cfg.err !== 'khong-vao-duoc') return syncNow(false);     // để báo lỗi như thường (một lần)
      schedulePoll(r.ok ? undefined : POLL_SLOW);
    }, function () { schedulePoll(POLL_SLOW); });   // rớt mạng: hỏi lại chậm hơn, không báo
  }
  /** Có sửa: hẹn gửi sau 4 giây (sửa liên tục thì dồn lại một lần). */
  function syncSoon() {
    if (!syncCfg) return;
    lastActivity = Date.now();
    syncDirty = true;
    clearTimeout(syncTimer);
    syncTimer = setTimeout(function () { syncTimer = null; syncNow(false); }, SYNC_DELAY);
    updateSyncUi();
  }
  /** Nhận số máy khác: vẽ lại, giữ chỗ đang xem. */
  function afterRemoteChange(oldStart) {
    if (state.settings.startDay !== oldStart && samePeriod(current, periodOf(today, oldStart))) current = todayPeriod();
    if (settingsOpen() && !(document.activeElement && elCols.contains(document.activeElement))) syncSettingsForm();
    var keep = elWrap.scrollTop;
    render();
    elWrap.scrollTop = keep;
    if (editing) { var td = elBody.querySelector('tr[data-key="' + editing.key + '"] td.cell[data-col="' + editing.col + '"]'); if (td) td.classList.add('selected'); }
  }
  /**
   * Một lần đồng bộ: lấy về → giải mã → gộp vào máy → máy có phần mới hơn thì mã hoá, gửi lên.
   * Máy khác vừa gửi chen (xung đột rev) thì làm lại, tối đa 3 lần.
   */
  function syncNow(manual) {
    if (!syncCfg) return Promise.resolve(false);
    if (saveBlocked) { if (manual) showNotice('App đang tạm không lưu (xem dòng màu vàng) nên chưa đồng bộ. Chạm để ẩn.'); return Promise.resolve(false); }
    if (syncBusy) { syncAgain = true; return syncBusy; }
    clearTimeout(syncTimer);
    syncTimer = null;
    var cfg = syncCfg, tries = 0, pushed = false;
    syncDirty = false;
    function round() {
      tries++;
      return api({ a: 'lay', g: cfg.g, t: cfg.t }).then(apiOk).then(function (r) {
        if (typeof r.qt === 'boolean') cfg.qt = r.qt;      // 1.4.3: nhóm đã có máy quản trị chưa
        return (r.data ? decryptDoc(cfg, r.data) : Promise.resolve(null)).then(function (remote) {
          if (syncCfg !== cfg) throw new Error('roi');
          // Gộp ngay (không chờ giữa lúc đọc và lúc ghi), nên không mất lần sửa nào đang diễn ra
          var m = remote ? mergeDocs(docOf(state), remote, 'sync') : { changed: false, push: true };
          if (m.changed) {
            lastActivity = Date.now();
            var oldStart = state.settings.startDay;
            applyDoc(m.doc);
            if (!saveState()) throw new Error('luu');
            afterRemoteChange(oldStart);
          }
          if (!m.push) return r.rev;
          return encryptDoc(cfg, docOf(state)).then(function (data) {
            return api({ a: 'gui', g: cfg.g, t: cfg.t, rev: r.rev, data: data });
          }).then(function (w) {
            if (w.ok) { pushed = true; return w.rev; }
            if (w.loi === 'xung-dot' && tries < 3) return round();
            throw new Error(w.loi || 'may-chu');
          });
        });
      });
    }
    updateSyncUi('Đang đồng bộ…');
    syncBusy = round().then(function (rev) {
      if (syncCfg !== cfg) return false;
      cfg.rev = rev;
      cfg.at = Date.now();
      cfg.err = '';
      syncWarned = '';
      syncFails = 0;
      delete cfg.doiTam;          // thẻ cũ còn dùng được nghĩa là lần đổi mật khẩu dở dang chưa tới máy chủ
      saveSyncCfg();
      if (manual) showNotice('Đã đồng bộ lúc ' + timeText(cfg.at) + '. Chạm để ẩn.', false, 5000);   // 1.4.6: tự ẩn sau 5 giây (1.4.5: 10 giây)
      return true;
    }, function (e) {
      if (syncCfg !== cfg || (e && e.message === 'roi')) return false;
      syncDirty = true;
      cfg.err = (e && e.message) || 'may-chu';
      saveSyncCfg();
      if (cfg.err === 'khong-vao-duoc' && cfg.doiTam) { adoptPendingPw(cfg); return false; }
      // Tự đồng bộ hỏng vì lý do không phải mạng: báo một lần; bấm "Đồng bộ ngay" thì luôn báo
      if (cfg.err === 'xung-dot') { setTimeout(function () { syncNow(false); }, 5000); if (!manual) return false; }
      // Rớt mạng / máy chủ chập chờn mà app vẫn mở: tự thử lại, giãn dần 30 giây → tối đa 5 phút
      if (cfg.err === 'mang' || cfg.err === 'may-chu' || cfg.err === 'qua-nhieu') {
        syncFails++;
        if (!syncTimer) syncTimer = setTimeout(function () { syncTimer = null; syncNow(false); }, Math.min(300000, 30000 * syncFails));
      }
      if (manual || (cfg.err !== 'mang' && syncWarned !== cfg.err)) {
        syncWarned = cfg.err;
        showNotice('Đồng bộ chưa được: ' + syncErrText(e) + ' Chạm để ẩn.');
      }
      return false;
    }).then(function (ok) {
      syncBusy = null;
      updateSyncUi();
      if (syncAgain) { syncAgain = false; setTimeout(function () { syncNow(false); }, 300); }
      else schedulePoll(pushed ? 3000 : undefined);    // vừa gửi xong: 3 giây sau hỏi lại một lần
      return ok;
    });
    return syncBusy;
  }
  function timeText(ms) {
    var d = new Date(ms), t = pad(d.getHours()) + ':' + pad(d.getMinutes());
    return dateKey(d) === dateKey(new Date()) ? t : t + ' ngày ' + d.getDate() + '/' + (d.getMonth() + 1);
  }

  // --- Giao diện đồng bộ trong Cài đặt ---
  var elSyncSheet = document.getElementById('syncSheet');
  var elSyncCode = document.getElementById('syncCode');
  var elSyncPw = document.getElementById('syncPw');
  var elSyncPw2 = document.getElementById('syncPw2');
  var elSyncMsg = document.getElementById('syncMsg');
  var elSyncOk = document.getElementById('syncOk');
  var syncMode = '';                  // 'create' | 'join' | 'done'
  var syncFormBusy = false;
  function syncSheetOpen() { return isShown(elSyncSheet); }

  function updateSyncUi(busyText) {
    var dot = document.getElementById('syncDot');
    if (dot) {
      dot.hidden = !syncCfg;
      var kind = !syncCfg ? '' : (busyText || syncBusy) ? 'busy' : syncCfg.err ? 'warn' : syncDirty ? 'busy' : 'ok';
      dot.className = 'sync-dot ' + kind;
      dot.textContent = kind === 'warn' ? '⚠' : kind === 'busy' ? '⟳' : '✓';
      dot.setAttribute('aria-label', kind === 'warn' ? 'Chưa đồng bộ được' : kind === 'busy' ? 'Đang đồng bộ' : 'Đã đồng bộ');
    }
    var st = document.getElementById('syncStatus');
    if (!st) return;
    document.getElementById('syncOff').hidden = !!syncCfg;
    document.getElementById('syncOnBox').hidden = !syncCfg;
    var isAdm = !!(syncCfg && syncCfg.adm);
    document.getElementById('syncAdminBox').hidden = !isAdm;
    document.getElementById('syncClaim').hidden = !(syncCfg && !isAdm && syncCfg.qt === false);
    document.getElementById('syncAdmCode').hidden = !(syncCfg && !isAdm && syncCfg.qt !== false);
    document.getElementById('syncRepw').hidden = !(syncCfg && syncCfg.err === 'khong-vao-duoc');
    if (!syncCfg) { st.textContent = 'Chưa vào nhóm. Tạo nhóm trên máy có số đầy đủ nhất, rồi các máy khác Vào nhóm.'; return; }
    var s = 'Nhóm ' + groupPretty(syncCfg.g) + (isAdm ? ' (máy quản trị)' : '') + ' · ';
    if (busyText || syncBusy) s += busyText || 'Đang đồng bộ…';
    else if (syncCfg.err) s += 'Chưa đồng bộ được: ' + (SYNC_ERR[syncCfg.err] || SYNC_ERR['may-chu']);
    else if (syncDirty) s += 'Có thay đổi, sắp gửi…';
    else s += syncCfg.at ? 'Đã đồng bộ lúc ' + timeText(syncCfg.at) : 'Chưa đồng bộ lần nào';
    st.textContent = s;
  }

  function openSyncSheet(mode) {
    // Sếp chốt 15/11 "Cho dùng thử": đang dùng thử cũng tạo / vào nhóm được. Hết hạn mà chưa có khoá thì chỉ vào nhóm để nhận số.
    if (mode === 'create' && isLocked()) {
      return openKey('Hết ' + TRIAL_DAYS + ' ngày dùng thử. Nhập khoá mở để tạo nhóm. Máy hết hạn vẫn Vào nhóm được để xem số.');
    }
    if (!hasCrypto()) { alert(SYNC_ERR['khong-ma-hoa']); return; }
    syncMode = mode;
    setShown(elSettings, false);
    document.getElementById('syncTitle').textContent = mode === 'create' ? 'Tạo nhóm đồng bộ' : 'Vào nhóm đồng bộ';
    document.getElementById('syncIntro').innerHTML = mode === 'create'
      ? 'Đặt <b>mật khẩu nhóm</b> (ít nhất 8 ký tự, nên dài và khó đoán, ví dụ một câu 3–4 từ). Mọi máy trong nhóm dùng chung mật khẩu này. Số trên máy này sẽ được gửi lên nhóm, đã khoá bằng mật khẩu: máy chủ không đọc được. <b>Quên mật khẩu thì không ai mở lại được nhóm</b>, nên ghi lại cẩn thận.'
      : 'Nhập <b>mã nhóm</b> (16 ký tự, có gạch hay không đều được) và <b>mật khẩu nhóm</b> lấy từ máy đã tạo nhóm. Số trên máy này sẽ được gộp với số của nhóm, không mất ô nào chỉ máy này có.';
    elSyncCode.hidden = mode !== 'join';
    elSyncPw2.hidden = mode !== 'create';
    elSyncPw.hidden = false;
    document.getElementById('syncShowRow').hidden = false;
    elSyncPw.setAttribute('autocomplete', mode === 'create' ? 'new-password' : 'current-password');
    elSyncCode.value = '';
    elSyncPw.value = '';
    elSyncPw2.value = '';
    elSyncMsg.textContent = '';
    document.getElementById('syncResult').hidden = true;
    document.getElementById('syncResultCopy').hidden = false;
    elSyncCode.placeholder = 'Mã nhóm, ví dụ ABCD-EFGH-JKMN-PQRS';
    elSyncCode.maxLength = 24;
    elSyncPw.placeholder = 'Mật khẩu nhóm (ít nhất 8 ký tự)';
    elSyncOk.hidden = false;
    elSyncOk.textContent = mode === 'create' ? 'Tạo nhóm' : 'Vào nhóm';
    setShown(elSyncSheet, true);
    setShown(elBackdrop, true);
  }
  function closeSyncSheet() {
    if (syncFormBusy) return;
    setShown(elSyncSheet, false);
    if (document.activeElement && elSyncSheet.contains(document.activeElement)) document.activeElement.blur();
    elSyncPw.value = '';
    elSyncPw2.value = '';
    openSettings(true);
  }
  function syncFormDone(text, showCode) {
    syncMode = 'done';
    elSyncCode.hidden = elSyncPw.hidden = elSyncPw2.hidden = true;
    document.getElementById('syncShowRow').hidden = true;
    elSyncPw.value = '';
    elSyncPw2.value = '';
    elSyncOk.hidden = true;
    document.getElementById('syncIntro').innerHTML = text;
    document.getElementById('syncResultCode').textContent = groupPretty(syncCfg.g);
    document.getElementById('syncResult').hidden = showCode === false;
    document.getElementById('syncResultCopy').hidden = false;
  }

  function submitSync() {
    if (syncFormBusy || (syncMode !== 'create' && syncMode !== 'join') || syncCfg) return;
    var pw = elSyncPw.value.normalize ? elSyncPw.value.normalize('NFC').trim() : elSyncPw.value.trim();
    if (pw.length < 8) { elSyncMsg.textContent = 'Mật khẩu nhóm cần ít nhất 8 ký tự.'; return; }
    var g;
    if (syncMode === 'create') {
      var pw2 = elSyncPw2.value.normalize ? elSyncPw2.value.normalize('NFC').trim() : elSyncPw2.value.trim();
      if (pw !== pw2) { elSyncMsg.textContent = 'Hai lần nhập mật khẩu chưa giống nhau.'; return; }
      g = randomGroup();
    } else {
      g = normGroup(elSyncCode.value);
      if (!g) { elSyncMsg.textContent = 'Mã nhóm gồm 16 chữ và số, ví dụ ABCD-EFGH-JKMN-PQRS.'; return; }
    }
    if (saveBlocked) { elSyncMsg.textContent = 'App đang tạm không lưu (xem dòng màu vàng) nên chưa đồng bộ được.'; return; }
    syncFormBusy = true;
    elSyncOk.disabled = true;
    elSyncMsg.textContent = 'Đang tạo khoá từ mật khẩu…';
    var kt, adm = syncMode === 'create' ? randomCode(26) : '';
    deriveGroup(g, pw).then(function (r) {
      kt = r;
      elSyncMsg.textContent = 'Đang nói chuyện với máy chủ…';
      if (syncMode === 'create') return api({ a: 'tao', g: g, t: kt.t, adm: adm }).then(apiOk);
      // Vào nhóm: lấy thử để biết đúng mã + mật khẩu và giải mã được, rồi mới bật
      return api({ a: 'lay', g: g, t: kt.t }).then(apiOk).then(function (r2) {
        return r2.data ? decryptDoc({ g: g, k: kt.k }, r2.data) : null;
      });
    }).then(function () {
      if (syncMode === 'join' && !stashCurrent('truoc-dongbo')) throw new Error('cat');
      syncCfg = { g: g, k: kt.k, t: kt.t, rev: 0, at: 0, err: '' };
      if (adm) { syncCfg.adm = adm; syncCfg.pw = pw; syncCfg.qt = true; }   // 1.4.3: máy tạo nhóm là máy quản trị
      saveSyncCfg();
      elSyncMsg.textContent = 'Đang đồng bộ lần đầu…';
      return syncNow(false);
    }).then(function (ok) {
      syncFormBusy = false;
      elSyncOk.disabled = false;
      elSyncMsg.textContent = ok ? '' : 'Đã vào nhóm nhưng lần đồng bộ đầu chưa xong: ' + (SYNC_ERR[syncCfg.err] || '') + ' App sẽ tự thử lại.';
      syncFormDone(syncMode === 'create'
        ? '<b>Đã tạo nhóm.</b> Ghi lại mã nhóm dưới đây và mật khẩu. Trên máy khác: Cài đặt → Đồng bộ nhiều máy → <b>Vào nhóm</b> → nhập mã và mật khẩu.'
        : '<b>Đã vào nhóm.</b> Từ giờ chấm ở máy nào, các máy cùng nhóm đều thấy (khi có mạng).');
    }, function (e) {
      syncFormBusy = false;
      elSyncOk.disabled = false;
      var c = e && e.message;
      if (c === 'cat') elSyncMsg.textContent = 'Máy hết chỗ để cất bản hiện tại, chưa vào nhóm.';
      else if (c === 'ton-tai') elSyncMsg.textContent = 'Trùng mã nhóm (rất hiếm). Bấm Tạo nhóm lần nữa.';
      else if (c === 'khong-vao-duoc') elSyncMsg.textContent = 'Sai mã nhóm hoặc mật khẩu.';
      else elSyncMsg.textContent = syncErrText(e);
    });
  }

  elSyncSheet.addEventListener('click', function (e) {
    var b = e.target.closest('button');
    if (!b) return;
    var a = b.getAttribute('data-action');
    if (a === 'sync-ok') { if (ADMIN_MODES[syncMode]) submitAdmin(); else submitSync(); }
    else if (a === 'sync-cancel') closeSyncSheet();
    else if (a === 'sync-copy') copyGroup(true);
  });
  document.getElementById('syncShowPw').addEventListener('change', function (e) {
    elSyncPw.type = elSyncPw2.type = e.target.checked ? 'text' : 'password';
  });
  function copyGroup(inSheet) {
    if (!syncCfg) return;
    var txt = 'Mã nhóm Sổ Công: ' + groupPretty(syncCfg.g);
    copyText(txt).then(function (ok) {
      var msg = ok ? 'Đã chép mã nhóm. Mật khẩu nhóm thì nói trực tiếp hoặc gõ tay, đừng gửi chung với mã.' : 'Máy không cho chép tự động. Mã nhóm: ' + groupPretty(syncCfg.g);
      if (inSheet) elSyncMsg.textContent = msg; else { closeSettings(); showNotice(msg + ' Chạm để ẩn.'); }
    });
  }
  function leaveGroup() {
    if (!syncCfg) return;
    if (!confirm('Rời nhóm ' + groupPretty(syncCfg.g) + '?\nSố trên máy này vẫn giữ nguyên, chỉ thôi tự đồng bộ. Các máy khác trong nhóm không bị ảnh hưởng.\nMuốn vào lại thì cần mã nhóm và mật khẩu.' +
      (syncCfg.adm ? '\n\nMáy này đang là MÁY QUẢN TRỊ: rời nhóm thì mất mã quản trị trên máy này. Nếu chưa có máy quản trị khác, hãy bấm Huỷ và "Chép mã quản trị" cất đi trước.' : ''))) return;
    syncCfg = null;
    clearTimeout(syncTimer);
    syncTimer = null;
    clearTimeout(pollTimer);
    pollTimer = null;
    syncDirty = false;
    aesCache = { k: '', key: null };
    saveSyncCfg();
    updateSyncUi();
  }


  // ---------- 1.4.3: quản trị nhóm — chỉ máy quản trị xem / đổi mật khẩu nhóm ----------
  var ADM_RE = /^[0-9A-HJKMNP-TV-Z]{26}$/;
  function randomCode(n) {
    var b = new Uint8Array(n), c = '';
    crypto.getRandomValues(b);
    for (var i = 0; i < n; i++) c += SYNC_CH[b[i] & 31];
    return c;
  }
  function normAdm(s) {
    var c = String(s || '').toUpperCase().replace(/[^0-9A-Z]/g, '').replace(/O/g, '0').replace(/[IL]/g, '1');
    if (c.length > 26 && ADM_RE.test(c.slice(-26))) c = c.slice(-26);   // dán cả câu "Mã quản trị …: XXXX-…" vẫn được
    return ADM_RE.test(c) ? c : '';
  }
  function normPw(v) { return v.normalize ? v.normalize('NFC').trim() : v.trim(); }
  var ADMIN_MODES = { repw: 1, savepw: 1, changepw: 1, admcode: 1, showpw: 1 };
  var ADMIN_TEXT = {
    repw: ['Nhập mật khẩu mới của nhóm', 'Mật khẩu nhóm có thể vừa được đổi. Nhập <b>mật khẩu hiện tại của nhóm</b> (hỏi người quản trị nhóm). Số trên máy này vẫn nguyên, nhập đúng là đồng bộ tiếp.', 'Đồng bộ lại'],
    savepw: ['Lưu mật khẩu trên máy này', 'Máy này chưa lưu mật khẩu nhóm. Nhập mật khẩu nhóm một lần: app kiểm đúng rồi lưu ngay trên máy này (không gửi đi đâu). Sau đó bấm <b>Xem mật khẩu</b> là thấy.', 'Lưu'],
    changepw: ['Đổi mật khẩu nhóm', 'Nhập <b>mật khẩu mới</b> 2 lần (ít nhất 8 ký tự, nên là một câu 3–4 từ). Mã nhóm giữ nguyên. Các máy khác trong nhóm sẽ báo cần <b>Nhập mật khẩu mới</b> rồi đồng bộ tiếp, không mất số.', 'Đổi mật khẩu'],
    admcode: ['Nhập mã quản trị', 'Dán <b>mã quản trị</b> 26 ký tự chép từ máy quản trị của bạn (Cài đặt → Đồng bộ nhiều máy → Chép mã quản trị). Máy có mã này xem và đổi được mật khẩu nhóm.', 'Nhận mã'],
    showpw: ['Mật khẩu nhóm', 'Mật khẩu nhóm đang lưu trên máy này. Chỉ đưa cho người được vào nhóm; nói trực tiếp, đừng gửi chung với mã nhóm.', '']
  };

  function openAdminSheet(mode) {
    if (!syncCfg) return;
    if (mode !== 'repw' && mode !== 'admcode' && !syncCfg.adm) return;
    if (mode === 'showpw' && !syncCfg.pw) mode = 'savepw';
    if (!hasCrypto()) { alert(SYNC_ERR['khong-ma-hoa']); return; }
    syncMode = mode;
    setShown(elSettings, false);
    var T = ADMIN_TEXT[mode];
    document.getElementById('syncTitle').textContent = T[0];
    document.getElementById('syncIntro').innerHTML = T[1];
    elSyncCode.hidden = mode !== 'admcode';
    elSyncCode.placeholder = mode === 'admcode' ? 'Mã quản trị 26 ký tự' : 'Mã nhóm, ví dụ ABCD-EFGH-JKMN-PQRS';
    elSyncCode.maxLength = mode === 'admcode' ? 200 : 24;
    elSyncPw.hidden = !(mode === 'repw' || mode === 'savepw' || mode === 'changepw');
    elSyncPw2.hidden = mode !== 'changepw';
    elSyncPw.placeholder = mode === 'changepw' ? 'Mật khẩu mới (ít nhất 8 ký tự)' : 'Mật khẩu nhóm (ít nhất 8 ký tự)';
    elSyncPw.setAttribute('autocomplete', mode === 'changepw' ? 'new-password' : 'current-password');
    document.getElementById('syncShowRow').hidden = elSyncPw.hidden;
    elSyncCode.value = elSyncPw.value = elSyncPw2.value = '';
    elSyncPw.type = elSyncPw2.type = 'password';
    document.getElementById('syncShowPw').checked = false;
    elSyncMsg.textContent = '';
    var showPw = mode === 'showpw';
    document.getElementById('syncResult').hidden = !showPw;
    document.getElementById('syncResultCopy').hidden = showPw;
    if (showPw) document.getElementById('syncResultCode').textContent = syncCfg.pw;
    elSyncOk.hidden = showPw;
    elSyncOk.textContent = T[2];
    setShown(elSyncSheet, true);
    setShown(elBackdrop, true);
  }

  /** Đổi mật khẩu: khớp số với nhóm → khoá lại cả sổ bằng khoá mới → máy chủ thay băm thẻ (kiểm mã quản trị, rev). */
  function changePassword(cfg, kt2, pw) {
    if (syncBusy) return Promise.reject(new Error('dang-ban'));
    clearTimeout(syncTimer); syncTimer = null;
    clearTimeout(pollTimer); pollTimer = null;
    var tries = 0;
    function round() {
      tries++;
      return api({ a: 'lay', g: cfg.g, t: cfg.t }).then(apiOk).then(function (r) {
        return (r.data ? decryptDoc(cfg, r.data) : Promise.resolve(null)).then(function (remote) {
          if (syncCfg !== cfg) throw new Error('roi');
          if (remote) {
            var m = mergeDocs(docOf(state), remote, 'sync');
            if (m.changed) {
              var oldStart = state.settings.startDay;
              applyDoc(m.doc);
              if (!saveState()) throw new Error('luu');
              afterRemoteChange(oldStart);
            }
          }
          return encryptDoc({ g: cfg.g, k: kt2.k }, docOf(state)).then(function (data) {
            // Rà 1.4.3 N1: nhớ khoá mới TRƯỚC khi gửi; máy chủ đã đổi mà mất phản hồi thì lần sau dùng khoá này
            cfg.doiTam = { k: kt2.k, t: kt2.t, pw: pw };
            saveSyncCfg();
            return api({ a: 'doi', g: cfg.g, t: cfg.t, adm: cfg.adm, t2: kt2.t, rev: r.rev, data: data });
          }).then(function (w) {
            if (w.ok) return w.rev;
            if (w.loi === 'xung-dot' && tries < 3) return round();
            throw new Error(w.loi || 'may-chu');
          });
        });
      });
    }
    var p = round().then(function (rev) {
      cfg.k = kt2.k; cfg.t = kt2.t; cfg.pw = pw; cfg.rev = rev; cfg.at = Date.now(); cfg.err = '';
      delete cfg.doiTam;
      aesCache = { k: '', key: null };
      saveSyncCfg();
      return true;
    });
    syncBusy = p.then(function () { return true; }, function () { return false; });   // giữ chỗ: không lần đồng bộ nào chen vào
    function after() {
      syncBusy = null;
      updateSyncUi();
      if (syncAgain) { syncAgain = false; syncNow(false); } else schedulePoll();
    }
    return p.then(function (v) { after(); return v; }, function (e) { after(); throw e; });
  }

  function submitAdmin() {
    if (syncFormBusy || !syncCfg) return;
    var cfg = syncCfg, mode = syncMode;
    function busy(on, text) { syncFormBusy = on; elSyncOk.disabled = on; if (text !== undefined) elSyncMsg.textContent = text; }
    function fail(e) {
      busy(false);
      var c = e && e.message;
      elSyncMsg.textContent = c === 'sai-mk' ? 'Mật khẩu chưa đúng.'
        : c === 'sai-ma-quan-tri' ? 'Mã quản trị không đúng.'
        : c === 'dang-ban' ? 'Đang đồng bộ, thử lại sau vài giây.'
        : c === 'khong-vao-duoc' ? 'Mật khẩu chưa đúng (hoặc nhóm không còn trên máy chủ).'
        : (mode === 'changepw' && (c === 'mang' || c === 'may-chu')) ? 'Mất kết nối giữa chừng: mật khẩu có thể đã đổi hoặc chưa. App đã nhớ mật khẩu mới và sẽ tự kiểm khi có mạng; bấm Đồng bộ ngay để kiểm.'
        : syncErrText(e);
    }
    if (mode === 'admcode') {
      var code = normAdm(elSyncCode.value);
      if (!code) { elSyncMsg.textContent = 'Mã quản trị gồm 26 chữ và số.'; return; }
      busy(true, 'Đang kiểm mã…');
      api({ a: 'qtri-kiem', g: cfg.g, t: cfg.t, adm: code }).then(apiOk).then(function () {
        if (syncCfg !== cfg) return;
        cfg.adm = code; cfg.qt = true;
        saveSyncCfg();
        busy(false, '');
        updateSyncUi();
        syncFormDone('<b>Máy này đã là máy quản trị.</b> Bấm <b>Xem mật khẩu</b> (nhập mật khẩu một lần để lưu) hoặc <b>Đổi mật khẩu</b> trong mục Đồng bộ nhiều máy.', false);
      }, fail);
      return;
    }
    var pw = normPw(elSyncPw.value);
    if (pw.length < 8) { elSyncMsg.textContent = 'Mật khẩu nhóm cần ít nhất 8 ký tự.'; return; }
    if (mode === 'changepw') {
      if (pw !== normPw(elSyncPw2.value)) { elSyncMsg.textContent = 'Hai lần nhập mật khẩu chưa giống nhau.'; return; }
      if (cfg.pw && pw === cfg.pw) { elSyncMsg.textContent = 'Mật khẩu mới trùng mật khẩu đang dùng.'; return; }
    }
    busy(true, 'Đang tạo khoá từ mật khẩu…');
    deriveGroup(cfg.g, pw).then(function (kt) {
      if (syncCfg !== cfg) throw new Error('roi');
      if (mode === 'savepw') {
        if (kt.t !== cfg.t) throw new Error('sai-mk');
        cfg.pw = pw;
        saveSyncCfg();
        return true;
      }
      if (mode === 'repw') {
        elSyncMsg.textContent = 'Đang thử mật khẩu với máy chủ…';
        return api({ a: 'lay', g: cfg.g, t: kt.t }).then(apiOk).then(function (r) {
          return r.data ? decryptDoc({ g: cfg.g, k: kt.k }, r.data) : null;
        }).then(function () {
          if (syncCfg !== cfg) throw new Error('roi');
          cfg.k = kt.k; cfg.t = kt.t; cfg.err = '';
          if (cfg.adm) cfg.pw = pw;
          aesCache = { k: '', key: null };
          saveSyncCfg();
          syncWarned = '';
          return syncNow(false);
        });
      }
      elSyncMsg.textContent = 'Đang khoá lại sổ bằng mật khẩu mới…';
      return changePassword(cfg, kt, pw);
    }).then(function () {
      busy(false, '');
      if (mode === 'savepw') return openAdminSheet('showpw');
      syncFormDone(mode === 'repw'
        ? '<b>Đã nhận mật khẩu mới.</b> Máy này đồng bộ tiếp như thường.'
        : '<b>Đã đổi mật khẩu nhóm.</b> Báo các máy khác trong nhóm: vào Cài đặt → Đồng bộ nhiều máy → <b>Nhập mật khẩu mới</b>. Bấm Xem mật khẩu để xem lại.', false);
    }, fail);
  }

  /** Đổi mật khẩu dở dang (mất phản hồi): thẻ cũ bị từ chối thì thử thẻ mới đã nhớ, đúng thì dùng luôn. */
  function adoptPendingPw(cfg) {
    var nt = cfg.doiTam;
    api({ a: 'lay', g: cfg.g, t: nt.t }).then(apiOk).then(function () {
      if (syncCfg !== cfg) return;
      cfg.k = nt.k; cfg.t = nt.t; cfg.pw = nt.pw; cfg.err = '';
      delete cfg.doiTam;
      aesCache = { k: '', key: null };
      saveSyncCfg();
      syncWarned = '';
      showNotice('Mật khẩu nhóm đã đổi xong (lần trước mất mạng lúc báo kết quả). Máy này dùng mật khẩu mới. Chạm để ẩn.');
      syncNow(false);
    }, function () {
      if (syncCfg !== cfg) return;
      if (syncWarned !== 'khong-vao-duoc') { syncWarned = 'khong-vao-duoc'; showNotice('Đồng bộ chưa được: ' + SYNC_ERR['khong-vao-duoc'] + ' Chạm để ẩn.'); }
      updateSyncUi();
    });
  }

  function claimAdmin() {
    var cfg = syncCfg;
    if (!cfg || cfg.adm) return;
    if (!confirm('Nhận quyền quản trị nhóm cho máy này?\nMáy quản trị xem và đổi được mật khẩu nhóm. Mỗi nhóm chỉ nhận được một lần; máy khác của bạn nhận bằng "Chép mã quản trị" từ máy này.')) return;
    // Rà 1.4.3 N1: nhớ mã TRƯỚC khi gửi; máy chủ đã ghi mà mất phản hồi thì lần bấm sau dùng lại đúng mã này
    var code = cfg.admTam && ADM_RE.test(cfg.admTam) ? cfg.admTam : randomCode(26);
    cfg.admTam = code;
    saveSyncCfg();
    function adopt() {
      if (syncCfg !== cfg) return;
      cfg.adm = code; cfg.qt = true;
      delete cfg.admTam;
      saveSyncCfg();
      updateSyncUi();
      showNotice('Máy này đã là máy quản trị nhóm: xem và đổi được mật khẩu. Máy khác của bạn: bấm "Chép mã quản trị" ở đây rồi dán vào máy kia. Chạm để ẩn.');
    }
    api({ a: 'qtri-nhan', g: cfg.g, t: cfg.t, adm: code }).then(apiOk).then(function () {
      if (syncCfg !== cfg) return;
      cfg.adm = code; cfg.qt = true;
      delete cfg.admTam;
      saveSyncCfg();
      updateSyncUi();
      showNotice('Máy này đã là máy quản trị nhóm: xem và đổi được mật khẩu. Máy khác của bạn: bấm "Chép mã quản trị" ở đây rồi dán vào máy kia. Chạm để ẩn.');
    }, function (e) {
      if (e && e.message === 'da-co-quan-tri') {
        // Có thể chính lần bấm trước của máy này đã nhận được (mất phản hồi): thử mã đã nhớ
        api({ a: 'qtri-kiem', g: cfg.g, t: cfg.t, adm: code }).then(apiOk).then(adopt, function (e2) {
          if (!e2 || e2.message !== 'sai-ma-quan-tri') { alert(syncErrText(e2) + ' Bấm lại "Nhận quyền quản trị nhóm" khi có mạng.'); return; }   // rớt mạng: giữ mã đã nhớ
          delete cfg.admTam;
          cfg.qt = true;
          saveSyncCfg();
          updateSyncUi();
          alert('Nhóm đã có máy quản trị. Lấy mã quản trị từ máy đó rồi bấm "Nhập mã quản trị".');
        });
      } else alert(syncErrText(e) + ' Bấm lại "Nhận quyền quản trị nhóm" khi có mạng.');
    });
  }
  function copyAdmin() {
    if (!syncCfg || !syncCfg.adm) return;
    copyText('Mã quản trị nhóm Sổ Công (giữ kín, chỉ dán vào máy của bạn): ' + groupPretty(syncCfg.adm)).then(function (ok) {
      closeSettings();
      showNotice((ok ? 'Đã chép mã quản trị. Gửi Zalo cho chính mình rồi trên máy kia: Cài đặt → Đồng bộ nhiều máy → Nhập mã quản trị. Đừng gửi cho người khác.'
        : 'Máy không cho chép tự động. Mã quản trị: ' + groupPretty(syncCfg.adm)) + ' Chạm để ẩn.', !ok);
    });
  }

  // Tự đồng bộ: mở app, quay lại app, có mạng lại; rời app thì gửi ngay phần đang chờ
  document.addEventListener('visibilitychange', function () {
    if (!syncCfg) return;
    if (document.hidden) { clearTimeout(pollTimer); pollTimer = null; if (syncTimer || syncDirty) syncNow(false); }
    else if (syncDirty || !syncCfg.at) syncNow(false);
    else pollRev();                                    // mở lại app: hỏi ngay có số mới không
  });
  window.addEventListener('online', function () { if (syncCfg) syncNow(false); });
  document.getElementById('syncDot').addEventListener('click', function () {
    openSettings();
    var fold = document.getElementById('syncFold');
    if (fold) fold.open = true;   // 1.4.6: mục Đồng bộ gập lại, chạm dấu ✓ thì mở ra
    var st = document.getElementById('syncStatus');
    if (st && st.scrollIntoView) st.scrollIntoView({ block: 'center' });
  });
  updateSyncUi();

  // Cho phiên rà soát / kiểm thử gọi được
  window.SoCong = {
    version: APP_VERSION,
    periodDays: periodDays, periodOf: periodOf, parseNumber: parseNumber,
    sanitizeState: sanitizeState, goTo: goTo, current: function () { return current; },
    notesText: function () { return notesText(current); },
    mergeDocs: mergeDocs, docOf: function () { return clone(docOf(state)); },
    syncNow: function () { return syncNow(true); }
  };

  document.getElementById('appVersion').textContent = APP_VERSION;
  render('today');
  remindBackup();
  initLicense();
  // ---------- 1.4.9: app Android báo có bản mới ----------
  function verNum(v) { var p = String(v || '').split('.'); return (+p[0] || 0) * 1e6 + (+p[1] || 0) * 1e3 + (+p[2] || 0); }
  window.SoCongBanMoi = function (txt) {
    try {
      var o = JSON.parse(txt);
      if (!o || typeof o.ver !== 'string' || !/^\d+\.\d+\.\d+$/.test(o.ver)) return;
      if (verNum(o.ver) <= verNum(APP_VERSION)) return;
      if (!elNotice.hidden) return;                 // đang có dòng báo khác: lần mở sau báo
      showNotice('Có Sổ Công ' + o.ver + ' mới (máy đang dùng ' + APP_VERSION + '). Chạm để tải, rồi cài đè lên, không gỡ app cũ.');
      noticeAction = function () {
        var A = window.SoCongAndroid;
        if (A && typeof A.moTrangTai === 'function') A.moTrangTai();
      };
    } catch (e) { /* bỏ qua */ }
  };
  if (window.SoCongAndroid && typeof window.SoCongAndroid.kiemBanMoi === 'function') {
    setTimeout(function () { try { window.SoCongAndroid.kiemBanMoi(); } catch (e) { /* bỏ qua */ } }, 2500);
  }

  if (firstRun && !syncCfg && !saveBlocked) openName();
  else if (!firstRun && migrateTo3()) {
    render('today');
    showNotice('Sổ Công đã gọn lại còn 3 cột: cột tên bạn, Full, Nghỉ. Cột P, P- nằm trong Cài đặt → Cột chấm công, bấm Lấy lại nếu cần. Chạm để ẩn.');
    if (!state.settings.name) openName();
  }
  if (syncCfg) setTimeout(function () { syncNow(false); }, 800);   // mở app: lấy số mới của nhóm

  // Hướng dẫn theo nơi đang chạy: app Android hay app web (iPhone)
  var isAndroidApp = !!(window.SoCongAndroid && window.SoCongAndroid.saveFile);
  document.documentElement.classList.add(isAndroidApp ? 'is-android-app' : 'is-web');
  // 1.4.4: iPad mở từ màn hình chính (iPadOS 26 trở lên): hệ thống làm mờ một dải phía trên, đè lên dòng
  // "Sổ Công · Kỳ …". Đẩy dòng này xuống dưới dải mờ.
  (function () {
    var ua = navigator.userAgent || '';
    var isIPad = /iPad/.test(ua) || (/Macintosh/.test(ua) && navigator.maxTouchPoints > 1);
    var standalone = navigator.standalone === true ||
      (window.matchMedia && window.matchMedia('(display-mode: standalone)').matches);
    if (!isAndroidApp && isIPad && standalone) document.documentElement.classList.add('ipad-pwa');
  })();

  // iPhone mở bằng tab Safari (chưa thêm vào màn hình chính): nhắc, vì dữ liệu tab Safari có thể bị xoá
  var standalone = (window.matchMedia && window.matchMedia('(display-mode: standalone)').matches) || window.navigator.standalone === true;
  var isIOS = /iPhone|iPad|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
  if (!isAndroidApp && isIOS && !standalone && !loadWarning) {
    showNotice('Bạn đang mở trong Safari. Bấm nút Chia sẻ → "Thêm vào MH chính" rồi mở Sổ Công từ màn hình chính. Dữ liệu trong tab Safari là kho riêng và có thể bị xoá sau 7 ngày không mở. Chạm để ẩn.');
  }
})();
