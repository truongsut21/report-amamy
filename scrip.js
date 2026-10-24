/**
 * Amamy – Dashboard tự động cho từng tab.
 * Chạy: buildAllDashboards()  (hoặc menu "📊 Dashboard" sau khi reload sheet)
 * Mỗi tab nguồn sẽ có 1 tab "📊 <tên tab>" ngay bên cạnh. Chạy lại = xoá & tạo lại.
 */
var PREFIX = '📊 ';
var COLORS = ['#2E6BD8', '#F2994A', '#27AE60', '#EB5757', '#9B51E0', '#56CCF2', '#F2C94C',
  '#6FCF97', '#BB6BD9', '#828282', '#2D9CDB', '#219653', '#F2A0A0', '#4F4F4F', '#A0C4FF', '#FFD6A5'];
var H_ROW = 3;      // hàng bắt đầu bảng dữ liệu phụ
var CH = 400;       // chiều cao chart (px)
var ROWS = [8, 29, 50]; // vị trí hàng của các dải chart

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 Dashboard')
    .addItem('Tạo / cập nhật tất cả dashboard', 'buildAllDashboards')
    .addToUi();
}

function buildAllDashboards() {
  var ss = SpreadsheetApp.getActive();
  var builders = [
    ['Khách hàng', buildCustomers],
    ['Chi tiết đơn', buildOrders],
    ['Theo line', function (d, s) { buildDimension(d, s, 'line'); }],
    ['Theo quốc gia', function (d, s) { buildDimension(d, s, 'quốc gia'); }],
    ['Theo thành phố', function (d, s) { buildDimension(d, s, 'thành phố'); }],
    ['Theo mã nguồn', function (d, s) { buildDimension(d, s, 'mã nguồn / NV'); }],
    ['Theo tháng', buildMonthly],
    ['Phân loại khách', function (d, s) { buildSegment(d, s, 'nhóm phân loại'); }],
    ['Theo kênh', function (d, s) { buildSegment(d, s, 'kênh'); }]
  ];
  builders.forEach(function (b) {
    var src = ss.getSheetByName(b[0]);
    if (!src) return;
    var old = ss.getSheetByName(PREFIX + b[0]);
    if (old) ss.deleteSheet(old);
    var dash = ss.insertSheet(PREFIX + b[0], src.getIndex());
    dash.setHiddenGridlines(true).setTabColor('#2E6BD8');
    dash.setColumnWidths(1, 12, 100);
    b[1](dash, src);
    SpreadsheetApp.flush();
  });
}

/* ---------------- helpers ---------------- */
function q(s) { return "'" + s.getName().replace(/'/g, "''") + "'!"; }
function col(s, c) { return q(s) + c + '2:' + c; }

function header(d, s, sub) {
  d.getRange('A1').setValue('Dashboard – ' + s.getName())
    .setFontSize(18).setFontWeight('bold').setFontColor('#1F2937');
  d.getRange('A2').setValue(sub || ('Tổng hợp tự động từ tab "' + s.getName() + '" · số liệu cập nhật theo dữ liệu nguồn'))
    .setFontColor('#6B7280').setFontStyle('italic');
  d.getRange(1, 15).setValue('DỮ LIỆU CHO BIỂU ĐỒ (công thức tự tính – không sửa tay)')
    .setFontWeight('bold').setFontColor('#9CA3AF');
}

function kpis(d, items) {
  items.forEach(function (it, i) {
    var c = 1 + i * 2;
    d.getRange(4, c, 1, 2).merge().setValue(it[0]).setFontColor('#6B7280').setFontSize(9);
    d.getRange(5, c, 1, 2).merge().setFormula(it[1]).setNumberFormat(it[2] || '#,##0')
      .setFontSize(15).setFontWeight('bold').setFontColor('#111827').setHorizontalAlignment('left');
    d.getRange(4, c, 2, 2).setBackground('#F3F4F6')
      .setBorder(true, true, true, true, null, null, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID_THICK);
  });
}

function note(d, text) {
  d.getRange('A6').setValue('💡 ' + text).setFontColor('#B45309').setFontSize(9);
}

function setF(d, c, formula, r) { d.getRange(r || H_ROW, c).setFormula(formula); }

function chart(d, type, range, row, colPos, o) {
  o = o || {};
  var b = d.newChart().setChartType(type).addRange(range)
    .setNumHeaders(1)
    .setPosition(row, colPos, 5, 0)
    .setOption('title', o.title || '')
    .setOption('titleTextStyle', { fontSize: 13, bold: true, color: '#1F2937' })
    .setOption('width', o.w || 590)
    .setOption('height', o.h || CH)
    .setOption('legend', { position: o.legend || 'none' })
    .setOption('colors', o.colors || COLORS);
  var opt = o.opt || {};
  for (var k in opt) b = b.setOption(k, opt[k]);
  d.insertChart(b.build());
}

function distinctCount(values) {
  var m = {};
  values.forEach(function (v) { if (v !== '') m[v] = 1; });
  return Object.keys(m).length;
}

/* ---------------- Theo line / quốc gia / thành phố / mã nguồn ---------------- */
// Cột nguồn: A Nhóm, B Số đơn, C Số khách, D Quay lại, E Tỉ lệ QL, F Đơn có kg&tiền, G Kg, H Tiền, I %DT, J TB/đơn, K TB/kg
function buildDimension(d, s, label) {
  var A = col(s, 'A'), B = col(s, 'B'), C = col(s, 'C'), E = col(s, 'E'), F = col(s, 'F'),
    H = col(s, 'H'), J = col(s, 'J'), K = col(s, 'K');
  var groups = Math.max(1, s.getLastRow() - 1);
  var n = Math.min(15, groups);
  header(d, s);
  kpis(d, [
    ['Số ' + label, '=COUNTA(' + A + ')'],
    ['Tổng đơn', '=SUM(' + B + ')'],
    ['Tổng doanh thu (VNĐ)', '=SUM(' + H + ')'],
    ['TB / đơn có tiền (VNĐ)', '=IFERROR(SUM(' + H + ')/SUM(' + F + '),0)'],
    ['Dẫn đầu doanh thu', '=IFERROR(INDEX(SORTN(FILTER({' + A + ',' + H + '},ISNUMBER(' + H + ')),1,0,2,FALSE),1,1),"-")', '@'],
    ['Top 3 chiếm % doanh thu', '=IFERROR(SUM(SORTN(FILTER(' + H + ',ISNUMBER(' + H + ')),3,0,1,FALSE))/SUM(' + H + '),0)', '0.0%']
  ]);
  note(d, 'Doanh thu, kg, TB/đơn chỉ tính trên các đơn có đủ kg & tiền – không phải toàn bộ đơn. Tỉ lệ quay lại lọc theo top ' + n + ' nhóm nhiều khách nhất để tránh nhiễu nhóm nhỏ.');

  setF(d, 15, '={"Nhóm","Doanh thu (VNĐ)";SORTN(FILTER({' + A + ',' + H + '},ISNUMBER(' + H + ')),' + n + ',0,2,FALSE)}');
  setF(d, 18, '={"Nhóm","Số đơn","Số khách";SORTN(FILTER({' + A + ',' + B + ',' + C + '},ISNUMBER(' + B + ')),' + n + ',0,2,FALSE)}');
  setF(d, 22, '={"Nhóm","Tỉ lệ quay lại";SORT(CHOOSECOLS(SORTN(FILTER({' + A + ',' + E + ',' + C + '},ISNUMBER(' + E + ')),' + n + ',0,3,FALSE),1,2),2,FALSE)}');
  var top = Math.min(7, groups);
  setF(d, 25, '={"Nhóm","Doanh thu";SORTN(FILTER({' + A + ',' + H + '},ISNUMBER(' + H + ')),' + top + ',0,2,FALSE)}');
  var otherRow = H_ROW + 1 + top;
  d.getRange(otherRow, 25).setValue('Khác');
  d.getRange(otherRow, 26).setFormula('=MAX(0,SUM(' + H + ')-SUM(' + d.getRange(H_ROW + 1, 26, top, 1).getA1Notation() + '))');
  setF(d, 28, '={"Nhóm","TB / đơn (VNĐ)","TB / kg (VNĐ)";CHOOSECOLS(SORTN(FILTER({' + A + ',' + J + ',' + K + ',' + H + '},ISNUMBER(' + J + ')),' + n + ',0,4,FALSE),1,2,3)}');

  chart(d, Charts.ChartType.BAR, d.getRange(H_ROW, 15, n + 1, 2), ROWS[0], 1,
    { title: 'Top ' + n + ' ' + label + ' theo doanh thu (VNĐ)', opt: { hAxis: { format: 'short' } } });
  chart(d, Charts.ChartType.COLUMN, d.getRange(H_ROW, 18, n + 1, 3), ROWS[0], 7,
    { title: 'Số đơn & số khách – top ' + n + ' ' + label, legend: 'top', opt: { vAxis: { format: 'short' } } });
  chart(d, Charts.ChartType.BAR, d.getRange(H_ROW, 22, n + 1, 2), ROWS[1], 1,
    { title: 'Tỉ lệ khách quay lại (≥2 đơn)', colors: ['#27AE60'], opt: { hAxis: { format: 'percent', minValue: 0 } } });
  chart(d, Charts.ChartType.PIE, d.getRange(H_ROW, 25, top + 2, 2), ROWS[1], 7,
    { title: 'Cơ cấu doanh thu theo ' + label, legend: 'right', opt: { pieHole: 0.45, pieSliceText: 'percentage' } });
  chart(d, Charts.ChartType.COMBO, d.getRange(H_ROW, 28, n + 1, 3), ROWS[2], 1,
    { title: 'Giá trị TB / đơn (cột) & TB / kg (đường) – top ' + n + ' theo doanh thu', legend: 'top', w: 1190,
      opt: { seriesType: 'bars', series: { 1: { type: 'line', targetAxisIndex: 1, pointSize: 6 } },
        vAxes: { 0: { title: 'TB / đơn', format: 'short' }, 1: { title: 'TB / kg', format: 'short' } } } });
}

/* ---------------- Theo tháng ---------------- */
function buildMonthly(d, s) {
  var A = col(s, 'A'), B = col(s, 'B'), C = col(s, 'C'), E = col(s, 'E'), H = col(s, 'H'), J = col(s, 'J'), K = col(s, 'K');
  var rows = Math.max(2, s.getLastRow());
  header(d, s);
  kpis(d, [
    ['Số tháng có dữ liệu', '=COUNTA(' + A + ')'],
    ['Tổng đơn', '=SUM(' + B + ')'],
    ['Đơn tháng gần nhất', '=INDEX(' + B + ',COUNTA(' + A + '))'],
    ['Đơn tháng trước', '=IFERROR(INDEX(' + B + ',COUNTA(' + A + ')-1),0)'],
    ['Tăng trưởng MoM', '=IFERROR(INDEX(' + B + ',COUNTA(' + A + '))/INDEX(' + B + ',COUNTA(' + A + ')-1)-1,0)', '+0.0%;-0.0%'],
    ['Đỉnh đơn / tháng', '=MAX(' + B + ')']
  ]);
  note(d, 'Tháng hiện tại có thể chưa đủ ngày → MoM sẽ thấp hơn thực tế. Ô "-" (chưa có kg/tiền) được bỏ trống trên biểu đồ.');

  setF(d, 15, '=ARRAYFORMULA({"Tháng","Số đơn","Số khách";FILTER({' + A + ',' + B + ',' + C + '},' + A + '<>"")})');
  setF(d, 19, '=ARRAYFORMULA({"Tháng","Doanh thu (VNĐ)";FILTER({' + A + ',IF(ISNUMBER(' + H + '),' + H + ',)},' + A + '<>"")})');
  setF(d, 22, '=ARRAYFORMULA({"Tháng","Tỉ lệ quay lại";FILTER({' + A + ',IF(ISNUMBER(' + E + '),' + E + ',)},' + A + '<>"")})');
  setF(d, 25, '=ARRAYFORMULA({"Tháng","TB / đơn (VNĐ)","TB / kg (VNĐ)";FILTER({' + A + ',IF(ISNUMBER(' + J + '),' + J + ',),IF(ISNUMBER(' + K + '),' + K + ',)},' + A + '<>"")})');

  chart(d, Charts.ChartType.COMBO, d.getRange(H_ROW, 15, rows, 3), ROWS[0], 1,
    { title: 'Số đơn (cột) & số khách (đường) theo tháng', legend: 'top',
      opt: { seriesType: 'bars', series: { 1: { type: 'line', pointSize: 6 } } } });
  chart(d, Charts.ChartType.COLUMN, d.getRange(H_ROW, 19, rows, 2), ROWS[0], 7,
    { title: 'Doanh thu theo tháng (VNĐ)', colors: ['#27AE60'], opt: { vAxis: { format: 'short' } } });
  chart(d, Charts.ChartType.LINE, d.getRange(H_ROW, 22, rows, 2), ROWS[1], 1,
    { title: 'Tỉ lệ khách quay lại theo tháng', colors: ['#9B51E0'],
      opt: { vAxis: { format: 'percent', minValue: 0 }, pointSize: 6, interpolateNulls: true } });
  chart(d, Charts.ChartType.LINE, d.getRange(H_ROW, 25, rows, 3), ROWS[1], 7,
    { title: 'Giá trị TB / đơn & TB / kg theo tháng', legend: 'top',
      opt: { vAxis: { format: 'short' }, pointSize: 6, interpolateNulls: true } });
}

/* ---------------- Phân loại khách / Theo kênh ---------------- */
// Cột nguồn: A Nhóm, B Số khách, C Quay lại, D Tỉ lệ QL, E Số đơn, F Kg, G Tiền, H %DT
function buildSegment(d, s, label) {
  var A = col(s, 'A'), B = col(s, 'B'), D = col(s, 'D'), E = col(s, 'E'), G = col(s, 'G');
  var rows = Math.max(2, s.getLastRow());
  header(d, s);
  kpis(d, [
    ['Số ' + label, '=COUNTA(' + A + ')'],
    ['Tổng khách', '=SUM(' + B + ')'],
    ['Tổng đơn', '=SUM(' + E + ')'],
    ['Đơn / khách', '=IFERROR(SUM(' + E + ')/SUM(' + B + '),0)', '0.00'],
    ['Tổng doanh thu (VNĐ)', '=SUM(' + G + ')'],
    ['DT cao nhất', '=IFERROR(INDEX(SORTN(FILTER({' + A + ',' + G + '},ISNUMBER(' + G + ')),1,0,2,FALSE),1,1),"-")', '@']
  ]);
  note(d, 'Doanh thu chỉ tính trên các đơn có đủ kg & tiền. So sánh cơ cấu số khách với cơ cấu doanh thu để thấy nhóm nào "đáng giá" nhất.');

  setF(d, 15, '={"Nhóm","Số khách";FILTER({' + A + ',' + B + '},' + A + '<>"")}');
  setF(d, 18, '=ARRAYFORMULA({"Nhóm","Doanh thu";FILTER({' + A + ',IF(ISNUMBER(' + G + '),' + G + ',)},' + A + '<>"")})');
  setF(d, 21, '={"Nhóm","Số khách","Số đơn";FILTER({' + A + ',' + B + ',' + E + '},' + A + '<>"")}');
  setF(d, 25, '=ARRAYFORMULA({"Nhóm","Doanh thu / khách (VNĐ)";FILTER({' + A + ',IF(ISNUMBER(' + G + '),' + G + '/' + B + ',)},' + A + '<>"")})');
  setF(d, 28, '=ARRAYFORMULA({"Nhóm","Tỉ lệ quay lại";FILTER({' + A + ',IF(ISNUMBER(' + D + '),' + D + ',)},' + A + '<>"")})');

  chart(d, Charts.ChartType.PIE, d.getRange(H_ROW, 15, rows, 2), ROWS[0], 1,
    { title: 'Cơ cấu số khách theo ' + label, legend: 'right', opt: { pieHole: 0.45 } });
  chart(d, Charts.ChartType.PIE, d.getRange(H_ROW, 18, rows, 2), ROWS[0], 7,
    { title: 'Cơ cấu doanh thu theo ' + label, legend: 'right', opt: { pieHole: 0.45 } });
  chart(d, Charts.ChartType.COLUMN, d.getRange(H_ROW, 21, rows, 3), ROWS[1], 1,
    { title: 'Số khách & số đơn theo ' + label, legend: 'top', opt: { vAxis: { format: 'short' } } });
  chart(d, Charts.ChartType.BAR, d.getRange(H_ROW, 25, rows, 2), ROWS[1], 7,
    { title: 'Doanh thu bình quân / khách (VNĐ)', colors: ['#F2994A'], opt: { hAxis: { format: 'short' } } });
  chart(d, Charts.ChartType.BAR, d.getRange(H_ROW, 28, rows, 2), ROWS[2], 1,
    { title: 'Tỉ lệ khách quay lại theo ' + label, colors: ['#27AE60'], opt: { hAxis: { format: 'percent', minValue: 0 } } });
}

/* ---------------- Khách hàng ---------------- */
// B ID, F Tên, P Số đơn, S Tổng tiền, U % lũy kế DT, V Đơn đầu, X Số ngày từ đơn gần nhất
function buildCustomers(d, s) {
  var B = col(s, 'B'), F = col(s, 'F'), P = col(s, 'P'), S = col(s, 'S'), U = col(s, 'U'), V = col(s, 'V'), X = col(s, 'X');
  header(d, s);
  kpis(d, [
    ['Tổng khách', '=COUNTA(' + B + ')'],
    ['Khách quay lại (≥2 đơn)', '=COUNTIF(' + P + ',">=2")'],
    ['Tỉ lệ quay lại', '=IFERROR(COUNTIF(' + P + ',">=2")/COUNTA(' + B + '),0)', '0.0%'],
    ['Khách hoạt động ≤90 ngày', '=COUNTIF(' + X + ',"<=90")'],
    ['DT TB / khách có tiền (VNĐ)', '=IFERROR(SUM(' + S + ')/COUNT(' + S + '),0)'],
    ['Top 20% khách chiếm % DT', '=IFERROR(INDEX(' + U + ',MAX(1,ROUND(COUNTA(' + B + ')*0.2))),0)', '0.0%']
  ]);
  note(d, 'Pareto & DT/khách dựa trên khách có dữ liệu tiền. "Khách hoạt động" = có đơn trong 90 ngày gần nhất.');

  var freq = [['Số đơn / khách', 'Số khách'],
    ['1 đơn', '=COUNTIF(' + P + ',1)'], ['2 đơn', '=COUNTIF(' + P + ',2)'], ['3 đơn', '=COUNTIF(' + P + ',3)'],
    ['4 đơn', '=COUNTIF(' + P + ',4)'], ['5–9 đơn', '=COUNTIFS(' + P + ',">=5",' + P + ',"<=9")'], ['10+ đơn', '=COUNTIF(' + P + ',">=10")']];
  d.getRange(H_ROW, 15, freq.length, 2).setValues(freq);
  var rec = [['Ngày từ đơn gần nhất', 'Số khách'],
    ['0–30', '=COUNTIFS(' + X + ',">=0",' + X + ',"<=30")'], ['31–60', '=COUNTIFS(' + X + ',">30",' + X + ',"<=60")'],
    ['61–90', '=COUNTIFS(' + X + ',">60",' + X + ',"<=90")'], ['91–180', '=COUNTIFS(' + X + ',">90",' + X + ',"<=180")'],
    ['181–365', '=COUNTIFS(' + X + ',">180",' + X + ',"<=365")'], ['>365', '=COUNTIF(' + X + ',">365")']];
  d.getRange(H_ROW, 18, rec.length, 2).setValues(rec);
  var par = [['Nhóm khách', '% doanh thu lũy kế']];
  [0.01, 0.05, 0.1, 0.2, 0.5].forEach(function (p) {
    par.push(['Top ' + (p * 100) + '%', '=IFERROR(INDEX(' + U + ',MAX(1,ROUND(COUNTA(' + B + ')*' + p + '))),0)']);
  });
  d.getRange(H_ROW, 21, par.length, 2).setValues(par);
  d.getRange(H_ROW + 1, 22, par.length - 1, 1).setNumberFormat('0.0%');
  var lastRow = s.getLastRow();
  var months = lastRow > 1 ? distinctCount(s.getRange(2, 22, lastRow - 1, 1).getDisplayValues()
    .map(function (r) { return String(r[0]).slice(0, 7); })) : 1;
  setF(d, 24, '=QUERY(ARRAYFORMULA({LEFT(TO_TEXT(' + V + '),7),' + B + '}),"select Col1, count(Col2) where Col1 <> \'\' group by Col1 order by Col1 label Col1 \'Tháng\', count(Col2) \'Khách mới\'",0)');
  setF(d, 27, '=ARRAYFORMULA({"Khách hàng","Doanh thu (VNĐ)";SORTN(FILTER({' + F + '&" ("&' + B + '&")",' + S + '},ISNUMBER(' + S + ')),15,0,2,FALSE)})');

  chart(d, Charts.ChartType.COLUMN, d.getRange(H_ROW, 15, 7, 2), ROWS[0], 1,
    { title: 'Phân bổ khách theo số đơn đã đặt', colors: ['#2E6BD8'], opt: { vAxis: { format: 'short' } } });
  chart(d, Charts.ChartType.COLUMN, d.getRange(H_ROW, 18, 7, 2), ROWS[0], 7,
    { title: 'Recency – số ngày từ đơn gần nhất', colors: ['#EB5757'], opt: { vAxis: { format: 'short' } } });
  chart(d, Charts.ChartType.COLUMN, d.getRange(H_ROW, 21, 6, 2), ROWS[1], 1,
    { title: 'Pareto – % doanh thu do top khách tạo ra', colors: ['#9B51E0'], opt: { vAxis: { format: 'percent', minValue: 0, maxValue: 1 } } });
  chart(d, Charts.ChartType.COLUMN, d.getRange(H_ROW, 24, months + 1, 2), ROWS[1], 7,
    { title: 'Khách mới theo tháng (theo đơn đầu tiên)', colors: ['#27AE60'] });
  chart(d, Charts.ChartType.BAR, d.getRange(H_ROW, 27, 16, 2), ROWS[2], 1,
    { title: 'Top 15 khách hàng theo doanh thu (VNĐ)', colors: ['#F2994A'], w: 1190, h: 450, opt: { hAxis: { format: 'short' } } });
}

/* ---------------- Chi tiết đơn ---------------- */
// A Ngày, B Tháng, C ID khách, E Line, J Quốc gia, N Kg, O Tiền, P Trạng thái
function buildOrders(d, s) {
  var Q = q(s);
  var A = col(s, 'A'), Bm = col(s, 'B'), E = col(s, 'E'), N = col(s, 'N'), O = col(s, 'O'), P = col(s, 'P');
  var dateArr = 'IFERROR(DATEVALUE(LEFT(TO_TEXT(' + A + '),10)),0)';
  header(d, s);
  kpis(d, [
    ['Tổng đơn', '=COUNTA(' + A + ')'],
    ['Đơn 7 ngày qua', '=SUMPRODUCT(--(' + dateArr + '>=TODAY()-6))'],
    ['Đơn 30 ngày qua', '=SUMPRODUCT(--(' + dateArr + '>=TODAY()-29))'],
    ['% đơn có đủ kg & tiền', '=IFERROR(COUNT(' + O + ')/COUNTA(' + A + '),0)', '0.0%'],
    ['% đơn chưa có trạng thái', '=IFERROR(1-COUNTA(' + P + ')/COUNTA(' + A + '),0)', '0.0%'],
    ['Tổng kg (đơn có dữ liệu)', '=SUM(' + N + ')', '#,##0.0']
  ]);
  note(d, 'Data quality: tỉ lệ đơn thiếu kg/tiền và trạng thái cao → nên chuẩn hoá nhập liệu. Trạng thái đang lẫn mã tiếng Anh & tiếng Việt.');

  setF(d, 15, '={"Ngày","Số đơn";SORT(QUERY(ARRAYFORMULA(LEFT(TO_TEXT(' + A + '),10)),"select Col1, count(Col1) where Col1 <> \'\' group by Col1 order by Col1 desc limit 60 label Col1 \'\', count(Col1) \'\'",0),1,TRUE)}');
  var lastRow = s.getLastRow();
  var vals = lastRow > 1 ? s.getRange(2, 1, lastRow - 1, 5).getDisplayValues() : [];
  var nMonths = distinctCount(vals.map(function (r) { return String(r[1]).slice(0, 7); }));
  var nLines = distinctCount(vals.map(function (r) { return r[4]; }));
  setF(d, 18, '=QUERY(ARRAYFORMULA({LEFT(TO_TEXT(' + Bm + '),7),' + E + '}),"select Col1, count(Col2) where Col1 <> \'\' and Col2 <> \'\' group by Col1 pivot Col2 label Col1 \'Tháng\'",0)');
  var wcol = 18 + nLines + 2;
  var wd = [['Thứ', 'Số đơn']];
  ['T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'CN'].forEach(function (t, i) {
    wd.push([t, '=SUMPRODUCT(--(IFERROR(WEEKDAY(DATEVALUE(LEFT(TO_TEXT(' + A + '),10)),2),0)=' + (i + 1) + '))']);
  });
  d.getRange(H_ROW, wcol, 8, 2).setValues(wd);
  setF(d, wcol + 3, '=QUERY(' + Q + 'P1:P,"select P, count(P) where P <> \'\' group by P order by count(P) desc limit 10 label count(P) \'Số đơn\'",1)');
  setF(d, wcol + 6, '=QUERY(' + Q + 'J1:J,"select J, count(J) where J <> \'\' group by J order by count(J) desc limit 10 label count(J) \'Số đơn\'",1)');

  chart(d, Charts.ChartType.COLUMN, d.getRange(H_ROW, 15, 61, 2), ROWS[0], 1,
    { title: 'Số đơn theo ngày – 60 ngày gần nhất', colors: ['#2E6BD8'], w: 1190 });
  chart(d, Charts.ChartType.COLUMN, d.getRange(H_ROW, 18, nMonths + 1, nLines + 1), ROWS[1], 1,
    { title: 'Số đơn theo tháng, chia theo line', legend: 'right', opt: { isStacked: true } });
  chart(d, Charts.ChartType.COLUMN, d.getRange(H_ROW, wcol, 8, 2), ROWS[1], 7,
    { title: 'Số đơn theo thứ trong tuần', colors: ['#56CCF2'] });
  chart(d, Charts.ChartType.BAR, d.getRange(H_ROW, wcol + 3, 11, 2), ROWS[2], 1,
    { title: 'Top 10 trạng thái đơn', colors: ['#F2994A'] });
  chart(d, Charts.ChartType.BAR, d.getRange(H_ROW, wcol + 6, 11, 2), ROWS[2], 7,
    { title: 'Top 10 quốc gia nhận theo số đơn', colors: ['#27AE60'] });
}
