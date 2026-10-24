/**
 * Web Worker: đọc file xlsx (ArrayBuffer) → model gọn cho dashboard.
 * Chạy ngoài main thread để giao diện không bị khựng khi parse file lớn.
 */
importScripts('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');

self.onmessage = function (e) {
  try {
    var wb = XLSX.read(e.data, { type: 'array', cellDates: true });
    self.postMessage({ ok: true, model: build(wb) });
  } catch (err) {
    self.postMessage({ ok: false, error: String((err && err.message) || err) });
  }
};

/* ---------------- helpers ---------------- */
function sheetRows(wb, name) {
  var ws = wb.Sheets[name];
  if (!ws) return null;
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null })
    .slice(1)
    .filter(function (r) { return r[0] != null && r[0] !== ''; });
}

// Số có thể là number, chuỗi "1,234", "12.5%", hoặc "-" (= trống)
function num(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  if (v == null || v instanceof Date) return 0;
  var s = String(v).trim();
  if (!s || s === '-') return 0;
  var pct = s.charAt(s.length - 1) === '%';
  var n = parseFloat(s.replace(/[%\s,]/g, ''));
  return isNaN(n) ? 0 : (pct ? n / 100 : n);
}

function pad(n) { return (n < 10 ? '0' : '') + n; }

// Chuẩn hoá ngày về "YYYY-MM-DD" (chấp nhận chuỗi, Date, hoặc serial Excel)
function day(v) {
  if (v instanceof Date) {
    var d = new Date(v.getTime() + 30000);
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
  }
  if (typeof v === 'number') return new Date(Math.round((v - 25569) * 864e5)).toISOString().slice(0, 10);
  return v == null ? '' : String(v).trim().slice(0, 10);
}

function countBy(list, keyFn) {
  var m = {};
  list.forEach(function (x) { var k = keyFn(x); if (k !== '' && k != null) m[k] = (m[k] || 0) + 1; });
  return m;
}

function topEntries(m, n) {
  return Object.keys(m).map(function (k) { return [k, m[k]]; })
    .sort(function (a, b) { return b[1] - a[1]; }).slice(0, n);
}

/* ---------------- builders ---------------- */
// Cột: A Nhóm, B Số đơn, C Số khách, D Quay lại, E Tỉ lệ QL, F Đơn có kg&tiền, G Kg, H Tiền, I %DT, J TB/đơn, K TB/kg
function dimension(rs) {
  return rs.map(function (r) {
    return { name: String(r[0]), orders: num(r[1]), customers: num(r[2]), returning: num(r[3]), rr: num(r[4]),
      paid: num(r[5]), kg: num(r[6]), rev: num(r[7]), pct: num(r[8]), avgOrder: num(r[9]), avgKg: num(r[10]) };
  });
}

function monthly(rs) {
  return dimension(rs).map(function (r, i) { r.name = day(rs[i][0]).slice(0, 7); return r; });
}

// Cột: A Nhóm, B Số khách, C Quay lại, D Tỉ lệ QL, E Số đơn, F Kg, G Tiền, H %DT
function segment(rs) {
  return rs.map(function (r) {
    return { name: String(r[0]), customers: num(r[1]), returning: num(r[2]), rr: num(r[3]),
      orders: num(r[4]), kg: num(r[5]), rev: num(r[6]), pct: num(r[7]) };
  });
}

// B ID, F Tên, P Số đơn, S Tổng tiền, U % lũy kế DT, V Đơn đầu, X Số ngày từ đơn gần nhất
function customers(rs) {
  var n = rs.length;
  var orders = rs.map(function (r) { return num(r[15]); });
  var rec = rs.map(function (r) { return r[23] == null || r[23] === '' ? null : num(r[23]); });
  var revs = rs.map(function (r) { return num(r[18]); });
  var cum = rs.map(function (r) { return num(r[20]); });
  var paid = revs.filter(function (v) { return v > 0; });
  var cumAt = function (p) { return n ? cum[Math.max(1, Math.round(n * p)) - 1] : 0; };
  var inRange = function (arr, lo, hi) { return arr.filter(function (v) { return v != null && v >= lo && v <= hi; }).length; };

  var newByMonth = countBy(rs, function (r) { return day(r[21]).slice(0, 7); });
  var top = rs.map(function (r, i) { return { name: String(r[5] || '?') + ' (' + r[1] + ')', rev: revs[i] }; })
    .filter(function (x) { return x.rev > 0; })
    .sort(function (a, b) { return b.rev - a.rev; }).slice(0, 15);

  return {
    total: n,
    returning: orders.filter(function (v) { return v >= 2; }).length,
    active90: inRange(rec, 0, 90),
    revPerPaid: paid.length ? paid.reduce(function (a, b) { return a + b; }, 0) / paid.length : 0,
    top20Share: cumAt(0.2),
    freq: [['1 đơn', inRange(orders, 1, 1)], ['2 đơn', inRange(orders, 2, 2)], ['3 đơn', inRange(orders, 3, 3)],
      ['4 đơn', inRange(orders, 4, 4)], ['5–9 đơn', inRange(orders, 5, 9)], ['10+ đơn', inRange(orders, 10, Infinity)]],
    recency: [['0–30', inRange(rec, 0, 30)], ['31–60', inRange(rec, 30.0001, 60)], ['61–90', inRange(rec, 60.0001, 90)],
      ['91–180', inRange(rec, 90.0001, 180)], ['181–365', inRange(rec, 180.0001, 365)], ['>365', inRange(rec, 365.0001, Infinity)]],
    pareto: [0.01, 0.05, 0.1, 0.2, 0.5].map(function (p) { return ['Top ' + (p * 100) + '%', cumAt(p)]; }),
    newByMonth: Object.keys(newByMonth).sort().map(function (k) { return [k, newByMonth[k]]; }),
    top: top
  };
}

// A Ngày, B Tháng, E Line, J Quốc gia, N Kg, O Tiền, P Trạng thái
function orders(rs) {
  var days = rs.map(function (r) { return day(r[0]); });
  var ref = days.reduce(function (a, b) { return b > a ? b : a; }, '');
  var refT = Date.parse(ref + 'T00:00:00Z');
  var ago = function (d) { return (refT - Date.parse(d + 'T00:00:00Z')) / 864e5; };

  // 60 ngày liên tục kết thúc tại ngày dữ liệu mới nhất (ngày không có đơn = 0)
  var perDay = countBy(days, function (d) { return d; });
  var daily = [];
  for (var i = 59; i >= 0; i--) {
    var k = new Date(refT - i * 864e5).toISOString().slice(0, 10);
    daily.push([k, perDay[k] || 0]);
  }

  var weekday = [0, 0, 0, 0, 0, 0, 0]; // T2..CN
  days.forEach(function (d) {
    var t = Date.parse(d + 'T00:00:00Z');
    if (!isNaN(t)) weekday[(new Date(t).getUTCDay() + 6) % 7]++;
  });

  var byMonthLine = {}, lineTotals = {};
  rs.forEach(function (r, i) {
    var m = day(r[1]).slice(0, 7) || days[i].slice(0, 7), l = r[4] == null ? '' : String(r[4]);
    if (!m || !l) return;
    (byMonthLine[m] = byMonthLine[m] || {})[l] = (byMonthLine[m][l] || 0) + 1;
    lineTotals[l] = (lineTotals[l] || 0) + 1;
  });

  var kg = 0, paid = 0, noStatus = 0;
  rs.forEach(function (r) {
    kg += num(r[13]);
    if (num(r[14]) > 0) paid++;
    if (r[15] == null || r[15] === '') noStatus++;
  });

  return {
    total: rs.length,
    refDate: ref,
    last7: days.filter(function (d) { var a = ago(d); return a >= 0 && a <= 6; }).length,
    last30: days.filter(function (d) { var a = ago(d); return a >= 0 && a <= 29; }).length,
    paidPct: rs.length ? paid / rs.length : 0,
    noStatusPct: rs.length ? noStatus / rs.length : 0,
    kg: kg,
    daily: daily,
    weekday: ['T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'CN'].map(function (t, i) { return [t, weekday[i]]; }),
    months: Object.keys(byMonthLine).sort(),
    byMonthLine: byMonthLine,
    lines: topEntries(lineTotals, 1e9).map(function (e) { return e[0]; }),
    status: topEntries(countBy(rs, function (r) { return r[15] == null ? '' : String(r[15]); }), 10),
    countries: topEntries(countBy(rs, function (r) { return r[9] == null ? '' : String(r[9]); }), 10)
  };
}

// Dữ liệu gốc của từng sheet để hiện bảng (bỏ các tab "📊 …" và cột trống ở cuối)
var TABLE_SHEETS = ['Khách hàng', 'Chi tiết đơn', 'Theo line', 'Theo quốc gia', 'Theo thành phố',
  'Theo mã nguồn', 'Theo tháng', 'Phân loại khách', 'Theo kênh'];

function tables(wb) {
  var out = {};
  TABLE_SHEETS.forEach(function (name) {
    var ws = wb.Sheets[name];
    if (!ws) return;
    var all = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    var header = (all[0] || []).map(function (h) { return h == null ? '' : String(h); });
    while (header.length && !header[header.length - 1]) header.pop();
    var rows = all.slice(1).filter(function (r) { return r[0] != null && r[0] !== ''; }).map(function (r) {
      return header.map(function (_, i) { var v = r[i]; return v instanceof Date ? day(v) : v; });
    });
    out[name] = { header: header, rows: rows };
  });
  return out;
}

function build(wb) {
  var get = function (name, fn) { var rs = sheetRows(wb, name); return rs ? fn(rs) : null; };
  var notes = wb.Sheets['Ghi chú'] ? XLSX.utils.sheet_to_json(wb.Sheets['Ghi chú'], { header: 1, raw: false })
    .map(function (r) { return r[0]; }).filter(Boolean) : [];
  return {
    builtAt: new Date().toISOString(),
    customers: get('Khách hàng', customers),
    orders: get('Chi tiết đơn', orders),
    line: get('Theo line', dimension),
    country: get('Theo quốc gia', dimension),
    city: get('Theo thành phố', dimension),
    source: get('Theo mã nguồn', dimension),
    monthly: get('Theo tháng', monthly),
    segment: get('Phân loại khách', segment),
    channel: get('Theo kênh', segment),
    notes: notes,
    tables: tables(wb)
  };
}
