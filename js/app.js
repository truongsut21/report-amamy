/**
 * Amamy dashboard – đọc data/amamy.xlsx, dựng các tab dashboard tương ứng với scrip.js (Apps Script).
 */
(function () {
  var DATA_URL = 'data/amamy.xlsx';
  var CACHE_KEY = 'amamy-model-v2';
  var f = C.fmt, S = C.series;

  var lastTab = null, model = null, sourceName = 'amamy.xlsx', charts = [], observer = null, worker = null;
  var $view = document.getElementById('view'), $tabs = document.getElementById('tabs'), $meta = document.getElementById('meta');

  var TABS = [
    { id: 'khach-hang', title: 'Khách hàng', key: 'customers', render: pageCustomers },
    { id: 'chi-tiet-don', title: 'Chi tiết đơn', key: 'orders', render: pageOrders },
    { id: 'line', title: 'Theo line', key: 'line', render: function (m) { return pageDimension(m.line, 'line', 'Theo line'); } },
    { id: 'quoc-gia', title: 'Theo quốc gia', key: 'country', render: function (m) { return pageDimension(m.country, 'quốc gia', 'Theo quốc gia'); } },
    { id: 'thanh-pho', title: 'Theo thành phố', key: 'city', render: function (m) { return pageDimension(m.city, 'thành phố', 'Theo thành phố'); } },
    { id: 'ma-nguon', title: 'Theo mã nguồn', key: 'source', render: function (m) { return pageDimension(m.source, 'mã nguồn / NV', 'Theo mã nguồn'); } },
    { id: 'thang', title: 'Theo tháng', key: 'monthly', render: pageMonthly },
    { id: 'phan-loai', title: 'Phân loại khách', key: 'segment', render: function (m) { return pageSegment(m.segment, 'nhóm phân loại', 'Phân loại khách'); } },
    { id: 'kenh', title: 'Theo kênh', key: 'channel', render: function (m) { return pageSegment(m.channel, 'kênh', 'Theo kênh'); } },
    { id: 'ghi-chu', title: 'Ghi chú', key: 'notes', render: pageNotes }
  ];

  /* ================= DOM helpers ================= */
  function el(tag, attrs, children) {
    var n = document.createElement(tag);
    if (attrs) for (var k in attrs) {
      if (k === 'text') n.textContent = attrs[k];
      else if (k === 'class') n.className = attrs[k];
      else if (attrs[k] != null) n.setAttribute(k, attrs[k]);
    }
    (children || []).forEach(function (c) { if (c) n.appendChild(typeof c === 'string' ? document.createTextNode(c) : c); });
    return n;
  }
  function sum(rows, k) { return rows.reduce(function (a, r) { return a + (r[k] || 0); }, 0); }
  function byDesc(k) { return function (a, b) { return b[k] - a[k]; }; }
  function pluck(rows, k) { return rows.map(function (r) { return r[k]; }); }
  function orNull(v) { return v ? v : null; } // 0 = chưa có dữ liệu → bỏ trống trên biểu đồ

  function kpi(label, value, sub) {
    return el('div', { class: 'kpi' }, [el('div', { class: 'label', text: label }), el('div', { class: 'value', text: value }),
      sub ? el('div', { class: 'sub', text: sub }) : null]);
  }

  // Card có biểu đồ: chỉ tạo Chart khi card cuộn vào màn hình (lazy) → chuyển tab mượt
  function card(o) {
    var canvas = el('canvas', { role: 'img', 'aria-label': o.title });
    var box = el('div', { class: 'chart-box' }, [canvas]);
    if (o.height) box.style.height = o.height + 'px';
    var c = el('section', { class: 'card' + (o.wide ? ' wide' : '') }, [
      el('h3', { text: o.title }), o.desc ? el('p', { class: 'desc', text: o.desc }) : null, box]);
    c._make = function () { charts.push(o.make(canvas)); };
    return c;
  }

  // Chiều cao cho bar ngang theo số nhóm
  function hbarHeight(n, grouped) { return Math.max(240, n * (grouped ? 36 : 24) + 60); }

  /* ---------- Bảng dữ liệu gốc của sheet: tìm kiếm, sắp xếp, phân trang, tải CSV ---------- */
  var PAGE_SIZE = 50;

  // Kiểu hiển thị theo tên cột (số đọc từ sheet là số thô)
  function colFormat(h) {
    h = String(h || '');
    var blank = /tiền|TB \/|Kg|kg|%|Tỉ lệ/.test(h); // 0 ở các cột này = chưa có dữ liệu
    var fn = /%|Tỉ lệ/.test(h) ? f.pct : (/^(Tổng\s+)?kg$/i.test(h) ? f.dec : f.int);
    return function (v) {
      if (v == null || v === '') return '';
      if (typeof v !== 'number') return String(v);
      return blank && v === 0 ? '–' : fn(v);
    };
  }

  function dataTable(name, t) {
    var fmts = t.header.map(colFormat);
    // Cột số căn phải: xét giá trị khác rỗng đầu tiên của cột
    var isNum = t.header.map(function (_, i) {
      for (var k = 0; k < t.rows.length; k++) { var v = t.rows[k][i]; if (v != null && v !== "") return typeof v === "number"; }
      return false;
    });
    var state = { col: -1, dir: 0, q: '', page: 0 }, list = t.rows;
    var tbody = el('tbody'), info = el('span', { class: 'muted' }), pager = el('div', { class: 'pager' });

    var heads = t.header.map(function (h, i) {
      var th = el('th', { class: isNum[i] ? 'num' : null, text: h, scope: 'col', title: 'Sắp xếp theo ' + h });
      th.addEventListener('click', function () {
        state.dir = state.col === i ? (state.dir === -1 ? 1 : -1) : -1;
        state.col = i; state.page = 0;
        heads.forEach(function (x) { x.removeAttribute('aria-sort'); });
        th.setAttribute('aria-sort', state.dir > 0 ? 'ascending' : 'descending');
        update();
      });
      return th;
    });

    function update() {
      var q = state.q.trim().toLowerCase();
      if (q && !t._search) t._search = t.rows.map(function (r) { return r.join('\u0001').toLowerCase(); });
      list = q ? t.rows.filter(function (_, i) { return t._search[i].indexOf(q) >= 0; }) : t.rows;
      if (state.dir) {
        var c = state.col, d = state.dir;
        list = list.slice().sort(function (a, b) {
          var x = a[c], y = b[c];
          if (x == null || x === '') return 1;
          if (y == null || y === '') return -1;
          return (typeof x === 'number' && typeof y === 'number' ? x - y : String(x).localeCompare(String(y), 'vi')) * d;
        });
      }
      draw();
    }

    function draw() {
      var pages = Math.max(1, Math.ceil(list.length / PAGE_SIZE));
      state.page = Math.min(state.page, pages - 1);
      var start = state.page * PAGE_SIZE, slice = list.slice(start, start + PAGE_SIZE);
      var frag = document.createDocumentFragment();
      slice.forEach(function (r) {
        frag.appendChild(el('tr', null, t.header.map(function (_, i) {
          var txt = fmts[i](r[i]);
          return el('td', { class: (isNum[i] ? 'num' : '') + (txt === '–' ? ' empty' : '') || null, text: txt });
        })));
      });
      if (!slice.length) frag.appendChild(el('tr', null, [el('td', { class: 'empty', colspan: t.header.length, text: 'Không có dòng nào khớp.' })]));
      tbody.replaceChildren(frag);
      info.textContent = list.length
        ? f.int(start + 1) + '–' + f.int(start + slice.length) + ' / ' + f.int(list.length) + ' dòng' + (list.length < t.rows.length ? ' (lọc từ ' + f.int(t.rows.length) + ')' : '')
        : '0 dòng';
      var btn = function (label, p, aria) {
        var b = el('button', { class: 'btn', type: 'button', text: label, 'aria-label': aria || label });
        b.disabled = p < 0 || p >= pages || p === state.page;
        b.addEventListener('click', function () { state.page = p; draw(); wrap.scrollTop = 0; });
        return b;
      };
      pager.replaceChildren(btn('«', 0, 'Trang đầu'), btn('‹', state.page - 1, 'Trang trước'),
        el('span', { class: 'muted', text: 'Trang ' + (state.page + 1) + ' / ' + pages }),
        btn('›', state.page + 1, 'Trang sau'), btn('»', pages - 1, 'Trang cuối'));
    }

    function exportCsv() {
      var esc = function (v) { v = v == null ? '' : String(v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
      var csv = [t.header].concat(list).map(function (r) { return r.map(esc).join(','); }).join('\n');
      var a = el('a', { href: URL.createObjectURL(new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' })), download: name + '.csv' });
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
    }

    var timer, input = el('input', { type: 'search', placeholder: 'Tìm trong bảng…', 'aria-label': 'Tìm trong bảng' });
    input.addEventListener('input', function () {
      clearTimeout(timer);
      timer = setTimeout(function () { state.q = input.value; state.page = 0; update(); }, 150);
    });
    var csvBtn = el('button', { class: 'btn', type: 'button', text: 'Tải CSV' });
    csvBtn.addEventListener('click', exportCsv);

    var wrap = el('div', { class: 'table-wrap' }, [el('table', null, [el('thead', null, [el('tr', null, heads)]), tbody])]);
    update();
    return el('section', { class: 'card wide table-card', id: 'bang-du-lieu' }, [
      el('div', { class: 'table-tools' }, [
        el('div', null, [el('h3', { text: 'Bảng dữ liệu – ' + name }), el('p', { class: 'desc', text: 'Dữ liệu gốc của tab "' + name + '" · bấm tiêu đề cột để sắp xếp' })]),
        el('div', { class: 'tools' }, [input, csvBtn])
      ]),
      wrap,
      el('div', { class: 'table-foot' }, [info, pager])
    ]);
  }

  function page(o) {
    return [
      el('div', { class: 'page-head' }, [el('h2', { text: o.title }), el('p', { text: o.sub || ('Tổng hợp từ tab "' + o.title + '" · ' + sourceName) })]),
      el('div', { class: 'kpis' }, o.kpis),
      o.note ? el('p', { class: 'note', text: '💡 ' + o.note }) : null,
      el('div', { class: 'grid' }, o.cards)
    ];
  }

  /* ================= Theo line / quốc gia / thành phố / mã nguồn ================= */
  function pageDimension(rows, label, title) {
    var n = Math.min(15, rows.length);
    var byRev = rows.filter(function (r) { return r.rev > 0; }).sort(byDesc('rev'));
    var totalRev = sum(rows, 'rev');
    var topRev = byRev.slice(0, n);
    var topOrders = rows.slice().sort(byDesc('orders')).slice(0, n);
    var topRR = rows.slice().sort(byDesc('customers')).slice(0, n).sort(byDesc('rr'));
    var pie = byRev.slice(0, 7);
    var other = Math.max(0, totalRev - sum(pie, 'rev'));
    var avg = byRev.filter(function (r) { return r.avgOrder > 0; }).slice(0, n);

    return page({
      title: title,
      kpis: [
        kpi('Số ' + label, f.int(rows.length)),
        kpi('Tổng đơn', f.int(sum(rows, 'orders'))),
        kpi('Tổng doanh thu', f.short(totalRev), f.vnd(totalRev)),
        kpi('TB / đơn có tiền', f.short(totalRev / (sum(rows, 'paid') || 1)), 'VNĐ'),
        kpi('Dẫn đầu doanh thu', byRev[0] ? byRev[0].name : '–', byRev[0] ? f.short(byRev[0].rev) : ''),
        kpi('Top 3 chiếm', f.pct(sum(byRev.slice(0, 3), 'rev') / (totalRev || 1)), 'doanh thu')
      ],
      note: 'Doanh thu, kg, TB/đơn chỉ tính trên các đơn có đủ kg & tiền – không phải toàn bộ đơn. Tỉ lệ quay lại lọc theo top ' + n + ' nhóm nhiều khách nhất để tránh nhiễu nhóm nhỏ.',
      cards: [
        card({ title: 'Top ' + topRev.length + ' ' + label + ' theo doanh thu', desc: 'VNĐ', height: hbarHeight(topRev.length),
          make: function (c) { return C.bar(c, { horizontal: true, labels: pluck(topRev, 'name'), fmt: f.vnd,
            datasets: [{ label: 'Doanh thu', data: pluck(topRev, 'rev'), color: S(5) }] }); } }),
        card({ title: 'Số đơn & số khách – top ' + n + ' ' + label, desc: 'Xếp theo số đơn', height: hbarHeight(n, true),
          make: function (c) { return C.bar(c, { horizontal: true, labels: pluck(topOrders, 'name'),
            datasets: [{ label: 'Số đơn', data: pluck(topOrders, 'orders'), color: S(0) }, { label: 'Số khách', data: pluck(topOrders, 'customers'), color: S(1) }] }); } }),
        card({ title: 'Tỉ lệ khách quay lại (≥2 đơn)', desc: 'Top ' + n + ' ' + label + ' nhiều khách nhất', height: hbarHeight(n),
          make: function (c) { return C.bar(c, { horizontal: true, labels: pluck(topRR, 'name'), fmt: f.pct, axisFmt: function (v) { return f.pct(v, 0); },
            datasets: [{ label: 'Tỉ lệ quay lại', data: pluck(topRR, 'rr'), color: S(6) }] }); } }),
        card({ title: 'Cơ cấu doanh thu theo ' + label, desc: 'Top 7 + Khác', height: 340,
          make: function (c) {
            var labels = pluck(pie, 'name'), data = pluck(pie, 'rev'), colors = pie.map(function (_, i) { return S(i); });
            if (other > 0) { labels.push('Khác'); data.push(other); colors.push(C.tok('other')); }
            return C.doughnut(c, { labels: labels, data: data, colors: colors, fmt: f.vnd });
          } }),
        card({ title: 'Giá trị TB / đơn', desc: 'VNĐ · top ' + avg.length + ' theo doanh thu', height: hbarHeight(avg.length),
          make: function (c) { return C.bar(c, { horizontal: true, labels: pluck(avg, 'name'), fmt: f.vnd,
            datasets: [{ label: 'TB / đơn', data: pluck(avg, 'avgOrder'), color: S(0) }] }); } }),
        card({ title: 'Giá trị TB / kg', desc: 'VNĐ · cùng nhóm với biểu đồ bên trái', height: hbarHeight(avg.length),
          make: function (c) { return C.bar(c, { horizontal: true, labels: pluck(avg, 'name'), fmt: f.vnd,
            datasets: [{ label: 'TB / kg', data: pluck(avg, 'avgKg'), color: S(1) }] }); } })
      ]
    });
  }

  /* ================= Theo tháng ================= */
  function pageMonthly(m) {
    var rows = m.monthly, n = rows.length;
    var last = rows[n - 1] || {}, prev = rows[n - 2] || {};
    var labels = rows.map(function (r) { return f.month(r.name); });
    var mom = prev.orders ? last.orders / prev.orders - 1 : 0;
    return page({
      title: 'Theo tháng',
      kpis: [
        kpi('Số tháng có dữ liệu', f.int(n)),
        kpi('Tổng đơn', f.int(sum(rows, 'orders'))),
        kpi('Đơn tháng gần nhất', f.int(last.orders), last.name ? f.month(last.name) : ''),
        kpi('Đơn tháng trước', f.int(prev.orders), prev.name ? f.month(prev.name) : ''),
        kpi('Tăng trưởng MoM', (mom >= 0 ? '+' : '') + f.pct(mom)),
        kpi('Đỉnh đơn / tháng', f.int(Math.max.apply(null, pluck(rows, 'orders').concat(0))))
      ],
      note: 'Tháng hiện tại có thể chưa đủ ngày → MoM sẽ thấp hơn thực tế. Tháng chưa có kg/tiền được bỏ trống trên biểu đồ.',
      cards: [
        card({ title: 'Số đơn & số khách theo tháng',
          make: function (c) { return C.bar(c, { labels: labels,
            datasets: [{ label: 'Số đơn', data: pluck(rows, 'orders'), color: S(0) }, { label: 'Số khách', data: pluck(rows, 'customers'), color: S(1) }] }); } }),
        card({ title: 'Doanh thu theo tháng', desc: 'VNĐ',
          make: function (c) { return C.bar(c, { labels: labels, fmt: f.vnd,
            datasets: [{ label: 'Doanh thu', data: rows.map(function (r) { return orNull(r.rev); }), color: S(5) }] }); } }),
        card({ title: 'Tỉ lệ khách quay lại theo tháng',
          make: function (c) { return C.line(c, { labels: labels, fmt: f.pct, axisFmt: function (v) { return f.pct(v, 0); },
            datasets: [{ label: 'Tỉ lệ quay lại', data: rows.map(function (r) { return orNull(r.rr); }), color: S(6) }] }); } }),
        card({ title: 'Giá trị TB / đơn theo tháng', desc: 'VNĐ',
          make: function (c) { return C.line(c, { labels: labels, fmt: f.vnd,
            datasets: [{ label: 'TB / đơn', data: rows.map(function (r) { return orNull(r.avgOrder); }), color: S(0) }] }); } }),
        card({ title: 'Giá trị TB / kg theo tháng', desc: 'VNĐ',
          make: function (c) { return C.line(c, { labels: labels, fmt: f.vnd,
            datasets: [{ label: 'TB / kg', data: rows.map(function (r) { return orNull(r.avgKg); }), color: S(1) }] }); } }),
        card({ title: 'Tổng kg theo tháng', desc: 'Đơn có dữ liệu kg',
          make: function (c) { return C.bar(c, { labels: labels, fmt: function (v) { return f.dec(v) + ' kg'; },
            datasets: [{ label: 'Kg', data: rows.map(function (r) { return orNull(r.kg); }), color: S(2) }] }); } })
      ]
    });
  }

  /* ================= Phân loại khách / Theo kênh ================= */
  function pageSegment(rows, label, title) {
    var totalRev = sum(rows, 'rev');
    var best = rows.slice().sort(byDesc('rev'))[0];
    var names = pluck(rows, 'name');
    var colors = rows.map(function (_, i) { return S(i); }); // màu gắn với nhóm, giống nhau ở mọi biểu đồ
    var perCust = rows.map(function (r) { return r.customers && r.rev ? r.rev / r.customers : null; });
    return page({
      title: title,
      kpis: [
        kpi('Số ' + label, f.int(rows.length)),
        kpi('Tổng khách', f.int(sum(rows, 'customers'))),
        kpi('Tổng đơn', f.int(sum(rows, 'orders'))),
        kpi('Đơn / khách', f.dec(sum(rows, 'orders') / (sum(rows, 'customers') || 1), 2)),
        kpi('Tổng doanh thu', f.short(totalRev), f.vnd(totalRev)),
        kpi('DT cao nhất', best && best.rev ? best.name : '–', best && best.rev ? f.short(best.rev) : '')
      ],
      note: 'Doanh thu chỉ tính trên các đơn có đủ kg & tiền. So sánh cơ cấu số khách với cơ cấu doanh thu để thấy nhóm nào "đáng giá" nhất.',
      cards: [
        card({ title: 'Cơ cấu số khách theo ' + label,
          make: function (c) { return C.doughnut(c, { labels: names, data: pluck(rows, 'customers'), colors: colors }); } }),
        card({ title: 'Cơ cấu doanh thu theo ' + label,
          make: function (c) { return C.doughnut(c, { labels: names, data: pluck(rows, 'rev'), colors: colors, fmt: f.vnd }); } }),
        card({ title: 'Số khách & số đơn theo ' + label, height: hbarHeight(rows.length, true),
          make: function (c) { return C.bar(c, { horizontal: true, labels: names, labelLen: 28,
            datasets: [{ label: 'Số khách', data: pluck(rows, 'customers'), color: S(1) }, { label: 'Số đơn', data: pluck(rows, 'orders'), color: S(0) }] }); } }),
        card({ title: 'Doanh thu bình quân / khách', desc: 'VNĐ', height: hbarHeight(rows.length, true),
          make: function (c) { return C.bar(c, { horizontal: true, labels: names, labelLen: 28, fmt: f.vnd,
            datasets: [{ label: 'DT / khách', data: perCust, color: S(5) }] }); } }),
        card({ title: 'Tỉ lệ khách quay lại theo ' + label, wide: true, height: hbarHeight(rows.length),
          make: function (c) { return C.bar(c, { horizontal: true, labels: names, labelLen: 36, fmt: f.pct, axisFmt: function (v) { return f.pct(v, 0); },
            datasets: [{ label: 'Tỉ lệ quay lại', data: rows.map(function (r) { return orNull(r.rr); }), color: S(6) }] }); } })
      ]
    });
  }

  /* ================= Khách hàng ================= */
  function pageCustomers(m) {
    var c = m.customers;
    var col = function (pairs, label, color, o) {
      return function (cv) { return C.bar(cv, Object.assign({ labels: pairs.map(function (p) { return p[0]; }),
        datasets: [{ label: label, data: pairs.map(function (p) { return p[1]; }), color: color }] }, o)); };
    };
    var pareto = c.pareto.map(function (p) { return [p[0], Math.min(1, p[1])]; });
    return page({
      title: 'Khách hàng',
      kpis: [
        kpi('Tổng khách', f.int(c.total)),
        kpi('Khách quay lại (≥2 đơn)', f.int(c.returning)),
        kpi('Tỉ lệ quay lại', f.pct(c.returning / (c.total || 1))),
        kpi('Khách hoạt động ≤90 ngày', f.int(c.active90)),
        kpi('DT TB / khách có tiền', f.short(c.revPerPaid), 'VNĐ'),
        kpi('Top 20% khách chiếm', f.pct(Math.min(1, c.top20Share)), 'doanh thu')
      ],
      note: 'Pareto & DT/khách dựa trên khách có dữ liệu tiền. "Khách hoạt động" = có đơn trong 90 ngày gần nhất.',
      cards: [
        card({ title: 'Phân bổ khách theo số đơn đã đặt', make: col(c.freq, 'Số khách', S(0)) }),
        card({ title: 'Recency – số ngày từ đơn gần nhất', make: col(c.recency, 'Số khách', S(7)) }),
        card({ title: 'Pareto – % doanh thu do top khách tạo ra', desc: 'Doanh thu lũy kế',
          make: col(pareto, '% doanh thu', S(6), { max: 1, fmt: f.pct, axisFmt: function (v) { return f.pct(v, 0); } }) }),
        card({ title: 'Khách mới theo tháng', desc: 'Theo tháng của đơn đầu tiên',
          make: col(c.newByMonth.map(function (p) { return [f.month(p[0]), p[1]]; }), 'Khách mới', S(2)) }),
        card({ title: 'Top 15 khách hàng theo doanh thu', desc: 'VNĐ', wide: true, height: hbarHeight(c.top.length),
          make: function (cv) { return C.bar(cv, { horizontal: true, labelLen: 32, labels: pluck(c.top, 'name'), fmt: f.vnd,
            datasets: [{ label: 'Doanh thu', data: pluck(c.top, 'rev'), color: S(1) }] }); } })
      ]
    });
  }

  /* ================= Chi tiết đơn ================= */
  function pageOrders(m) {
    var o = m.orders;
    var mainLines = o.lines.slice(0, 7), rest = o.lines.slice(7);
    var stack = mainLines.map(function (l, i) {
      return { label: l, color: S(i), data: o.months.map(function (mo) { return o.byMonthLine[mo][l] || 0; }) };
    });
    if (rest.length) stack.push({ label: 'Khác', color: C.tok('other'), data: o.months.map(function (mo) {
      return rest.reduce(function (a, l) { return a + (o.byMonthLine[mo][l] || 0); }, 0);
    }) });
    var pairs = function (p) { return { labels: p.map(function (x) { return x[0]; }), data: p.map(function (x) { return x[1]; }) }; };
    var st = pairs(o.status), ct = pairs(o.countries), wd = pairs(o.weekday), dl = pairs(o.daily);

    return page({
      title: 'Chi tiết đơn',
      sub: 'Tổng hợp từ tab "Chi tiết đơn" · số liệu tính đến ngày ' + f.date(o.refDate),
      kpis: [
        kpi('Tổng đơn', f.int(o.total)),
        kpi('Đơn 7 ngày qua', f.int(o.last7), 'đến ' + f.date(o.refDate)),
        kpi('Đơn 30 ngày qua', f.int(o.last30)),
        kpi('% đơn có đủ kg & tiền', f.pct(o.paidPct)),
        kpi('% đơn chưa có trạng thái', f.pct(o.noStatusPct)),
        kpi('Tổng kg', f.dec(o.kg), 'đơn có dữ liệu')
      ],
      note: 'Data quality: tỉ lệ đơn thiếu kg/tiền và trạng thái cao → nên chuẩn hoá nhập liệu. Trạng thái đang lẫn mã tiếng Anh & tiếng Việt.',
      cards: [
        card({ title: 'Số đơn theo ngày – 60 ngày gần nhất', wide: true,
          make: function (c) { return C.bar(c, { labels: dl.labels.map(function (d) { return f.date(d).slice(0, 5); }),
            datasets: [{ label: 'Số đơn', data: dl.data, color: S(0) }] }); } }),
        card({ title: 'Số đơn theo tháng, chia theo line', desc: 'Top 7 line + Khác', wide: true, height: 380,
          make: function (c) { return C.bar(c, { stacked: true, labels: o.months.map(f.month), datasets: stack }); } }),
        card({ title: 'Số đơn theo thứ trong tuần',
          make: function (c) { return C.bar(c, { labels: wd.labels, datasets: [{ label: 'Số đơn', data: wd.data, color: S(0) }] }); } }),
        card({ title: 'Top 10 quốc gia nhận theo số đơn',
          make: function (c) { return C.bar(c, { horizontal: true, labels: ct.labels, datasets: [{ label: 'Số đơn', data: ct.data, color: S(2) }] }); } }),
        card({ title: 'Top 10 trạng thái đơn', wide: true, height: hbarHeight(st.labels.length),
          make: function (c) { return C.bar(c, { horizontal: true, labelLen: 44, labels: st.labels, datasets: [{ label: 'Số đơn', data: st.data, color: S(1) }] }); } })
      ]
    });
  }

  /* ================= Ghi chú ================= */
  function pageNotes(m) {
    return [
      el('div', { class: 'page-head' }, [el('h2', { text: 'Ghi chú cách tính' }), el('p', { text: 'Từ tab "Ghi chú" · ' + sourceName })]),
      el('div', { class: 'notes' }, m.notes.slice(1).map(function (t) { return el('p', { text: t }); }))
    ];
  }

  /* ================= Router / render ================= */
  function available() {
    return TABS.filter(function (t) { var v = model[t.key]; return v && (!Array.isArray(v) || v.length); });
  }

  function current() {
    var tabs = available(), id = location.hash.slice(1);
    return tabs.filter(function (t) { return t.id === id; })[0] || tabs[0];
  }

  // Màn rộng: các nút tab (tự xuống dòng). Màn hẹp: dropdown (CSS quyết định cái nào hiện)
  function renderTabs() {
    var tabs = available();
    var select = el('select', { class: 'tab-select', 'aria-label': 'Chọn tab' }, tabs.map(function (t) {
      return el('option', { value: t.id, text: t.title });
    }));
    select.addEventListener('change', function () { location.hash = select.value; });
    var links = el('div', { class: 'tab-links', role: 'tablist' }, tabs.map(function (t) {
      return el('a', { href: '#' + t.id, role: 'tab', 'data-id': t.id, text: t.title });
    }));
    $tabs.replaceChildren(links, select);
  }

  function destroyCharts() {
    charts.forEach(function (ch) { ch && ch.destroy(); });
    charts = [];
    if (observer) observer.disconnect();
  }

  function render() {
    if (!model) return;
    var tab = current();
    if (!tab) { showError('File không có tab nào khớp với cấu trúc dashboard.'); return; }
    $tabs.querySelectorAll('a[data-id]').forEach(function (a) {
      a.setAttribute('aria-selected', a.getAttribute('data-id') === tab.id);
    });
    var select = $tabs.querySelector('select');
    if (select) select.value = tab.id;
    // Đổi tab khi đang cuộn giữa trang → đưa về đầu nội dung tab mới
    if (lastTab && lastTab !== tab.id) {
      var top = $view.getBoundingClientRect().top + window.scrollY - $tabs.offsetHeight;
      if (window.scrollY > top) window.scrollTo(0, top);
    }
    lastTab = tab.id;
    destroyCharts();
    var nodes = tab.render(model);
    var sheet = model.tables && model.tables[tab.title]; // tên tab = tên sheet nguồn
    if (sheet) nodes.push(dataTable(tab.title, sheet));
    $view.classList.remove('enter');
    $view.replaceChildren.apply($view, nodes.filter(Boolean));
    void $view.offsetWidth;
    $view.classList.add('enter');
    document.title = tab.title + ' · Amamy Dashboard';

    observer = new IntersectionObserver(function (entries) {
      entries.forEach(function (e) {
        if (!e.isIntersecting || !e.target._make) return;
        var mk = e.target._make; e.target._make = null; observer.unobserve(e.target);
        requestAnimationFrame(mk);
      });
    }, { rootMargin: '200px 0px' });
    $view.querySelectorAll('.card').forEach(function (c) { if (c._make) observer.observe(c); });
  }

  function showModel(m, name, fromCache) {
    model = m; sourceName = name;
    var o = m.orders, c = m.customers;
    $meta.textContent = [name, o ? f.int(o.total) + ' đơn' : '', c ? f.int(c.total) + ' khách' : '',
      o && o.refDate ? 'dữ liệu đến ' + f.date(o.refDate) : '', fromCache ? 'đã lưu cache' : '']
      .filter(Boolean).join(' · ');
    renderTabs();
    render();
  }

  function showLoading(text) {
    destroyCharts();
    $view.replaceChildren(el('div', { class: 'loading' }, [el('div', { class: 'spinner' }), el('p', { text: text })]));
  }

  function showError(msg) {
    destroyCharts();
    $meta.textContent = 'Chưa có dữ liệu';
    $view.replaceChildren(el('div', { class: 'error' }, [el('div', { class: 'box' }, [
      el('h2', { text: 'Không đọc được dữ liệu' }), el('p', { text: msg }),
      el('p', { class: 'muted', text: 'Nếu mở trực tiếp file index.html (file://), trình duyệt sẽ chặn đọc data/amamy.xlsx. Hãy chạy qua server (vd: npx serve) hoặc bấm "Mở file xlsx" / kéo thả file vào trang.' })
    ])]));
  }

  /* ================= Load data ================= */
  function parse(buffer) {
    return new Promise(function (resolve, reject) {
      worker = worker || new Worker('js/worker.js');
      worker.onmessage = function (e) { e.data.ok ? resolve(e.data.model) : reject(new Error(e.data.error)); };
      worker.onerror = function (e) { reject(new Error(e.message || 'Lỗi worker')); };
      worker.postMessage(buffer, [buffer]);
    });
  }

  // Cache model trong IndexedDB (bảng dữ liệu gốc vài MB – vượt quota của localStorage)
  function idb(mode, fn) {
    return new Promise(function (resolve) {
      try {
        var req = indexedDB.open('amamy-dashboard', 1);
        req.onupgradeneeded = function () { req.result.createObjectStore('cache'); };
        req.onerror = function () { resolve(null); };
        req.onsuccess = function () {
          try {
            var tx = req.result.transaction('cache', mode), r = fn(tx.objectStore('cache'));
            tx.oncomplete = function () { resolve(r && r.result); };
            tx.onerror = tx.onabort = function () { resolve(null); };
          } catch (e) { resolve(null); }
        };
      } catch (e) { resolve(null); }
    });
  }
  function readCache() { return idb('readonly', function (s) { return s.get(CACHE_KEY); }); }
  function writeCache(v) { return idb('readwrite', function (s) { return s.put(v, CACHE_KEY); }); }

  function loadRemote(force) {
    showLoading('Đang tải dữ liệu…');
    var cached = null;
    return readCache().then(function (c) {
      cached = c;
      return fetch(DATA_URL, { cache: force ? 'reload' : 'no-cache' });
    }).then(function (res) {
      if (!res.ok) throw new Error('Không tải được ' + DATA_URL + ' (HTTP ' + res.status + ')');
      // Chữ ký file: nếu không đổi thì dùng model đã parse sẵn → mở trang gần như tức thì
      var sig = [res.headers.get('etag'), res.headers.get('last-modified'), res.headers.get('content-length')].join('|');
      if (!force && cached && cached.sig === sig && sig !== '||') { showModel(cached.model, 'amamy.xlsx', true); return; }
      showLoading('Đang đọc file Excel…');
      return res.arrayBuffer().then(parse).then(function (m) {
        writeCache({ sig: sig, model: m });
        showModel(m, 'amamy.xlsx');
      });
    }).catch(function (err) {
      if (cached && !force) { showModel(cached.model, 'amamy.xlsx', true); return; }
      showError(err.message);
    });
  }

  function loadFile(file) {
    if (!file) return;
    showLoading('Đang đọc ' + file.name + '…');
    file.arrayBuffer().then(parse).then(function (m) { showModel(m, file.name); })
      .catch(function (err) { showError(err.message); });
  }

  /* ================= Events ================= */
  window.addEventListener('hashchange', render);
  document.getElementById('reload').addEventListener('click', function () { loadRemote(true); });
  document.getElementById('file').addEventListener('change', function (e) { loadFile(e.target.files[0]); e.target.value = ''; });

  var $drop = document.getElementById('drop'), dragDepth = 0;
  window.addEventListener('dragenter', function (e) { e.preventDefault(); dragDepth++; $drop.hidden = false; });
  window.addEventListener('dragleave', function () { if (--dragDepth <= 0) { dragDepth = 0; $drop.hidden = true; } });
  window.addEventListener('dragover', function (e) { e.preventDefault(); });
  window.addEventListener('drop', function (e) {
    e.preventDefault(); dragDepth = 0; $drop.hidden = true;
    loadFile(e.dataTransfer.files[0]);
  });

  // Theme: lưu lựa chọn, vẽ lại biểu đồ để lấy màu mới
  function setTheme(t) {
    if (t) document.documentElement.setAttribute('data-theme', t);
    try { localStorage.setItem('amamy-theme', t); } catch (e) { /* bỏ qua */ }
  }
  try { var saved = localStorage.getItem('amamy-theme'); if (saved) document.documentElement.setAttribute('data-theme', saved); } catch (e) { /* bỏ qua */ }
  document.getElementById('theme').addEventListener('click', function () {
    var dark = document.documentElement.getAttribute('data-theme') === 'dark' ||
      (!document.documentElement.getAttribute('data-theme') && matchMedia('(prefers-color-scheme: dark)').matches);
    setTheme(dark ? 'light' : 'dark');
    render();
  });
  matchMedia('(prefers-color-scheme: dark)').addEventListener('change', function () {
    if (!document.documentElement.getAttribute('data-theme')) render();
  });

  try { localStorage.removeItem('amamy-model-v1'); } catch (e) { /* cache cũ */ }
  loadRemote(false);
})();
