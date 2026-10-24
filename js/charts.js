/**
 * Chart.js helpers + định dạng số. Màu đọc từ CSS variables nên tự đổi theo light/dark.
 */
var C = (function () {
  function tok(name) { return getComputedStyle(document.documentElement).getPropertyValue('--' + name).trim(); }
  // Màu theo thứ tự cố định; quá 8 nhóm thì gộp vào "Khác" (màu xám) – không sinh màu mới
  function series(i) { return i < 8 ? tok('series-' + (i + 1)) : tok('other'); }

  var nf = [0, 1, 2].map(function (d) { return new Intl.NumberFormat('vi-VN', { maximumFractionDigits: d }); });
  var fmt = {
    int: function (v) { return nf[0].format(v || 0); },
    dec: function (v, d) { return nf[d == null ? 1 : d].format(v || 0); },
    pct: function (v, d) { return nf[d == null ? 1 : d].format((v || 0) * 100) + '%'; },
    short: function (v) {
      var a = Math.abs(v || 0);
      if (a >= 1e9) return nf[2].format(v / 1e9) + ' tỷ';
      if (a >= 1e6) return nf[1].format(v / 1e6) + ' tr';
      if (a >= 1e3) return nf[1].format(v / 1e3) + 'k';
      return nf[0].format(v || 0);
    },
    vnd: function (v) { return nf[0].format(v || 0) + ' ₫'; },
    date: function (s) { var p = String(s || '').split('-'); return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : s; },
    month: function (s) { var p = String(s || '').split('-'); return p.length >= 2 ? p[1] + '/' + p[0] : s; }
  };

  function trunc(s, n) { s = String(s); return s.length > n ? s.slice(0, n - 1) + '…' : s; }

  Chart.defaults.font.family = '"Be Vietnam Pro", system-ui, sans-serif';
  Chart.defaults.font.size = 12;
  Chart.defaults.animation.duration = 450;
  Chart.defaults.maintainAspectRatio = false;
  Chart.defaults.responsive = true;

  function tooltip(valueFmt) {
    return {
      backgroundColor: tok('surface-1'), titleColor: tok('text-primary'), bodyColor: tok('text-secondary'),
      borderColor: tok('border'), borderWidth: 1, padding: 10, cornerRadius: 8, boxPadding: 4,
      usePointStyle: true, titleFont: { weight: '600' },
      callbacks: {
        label: function (ctx) {
          var v = ctx.parsed && typeof ctx.parsed === 'object' ? (ctx.chart.options.indexAxis === 'y' ? ctx.parsed.x : ctx.parsed.y) : ctx.parsed;
          if (v == null) return ' ' + ctx.dataset.label + ': –';
          return ' ' + (ctx.dataset.label ? ctx.dataset.label + ': ' : '') + valueFmt(v);
        }
      }
    };
  }

  function legend(show) {
    return {
      display: show, position: 'top', align: 'start',
      labels: { color: tok('text-secondary'), usePointStyle: true, pointStyle: 'circle', boxWidth: 8, boxHeight: 8, padding: 14 }
    };
  }

  /**
   * o: { labels, datasets:[{label,data,color}], horizontal, stacked, fmt (tooltip), axisFmt, max, labelLen }
   */
  function bar(canvas, o) {
    var text2 = tok('text-secondary'), grid = tok('grid'), surface = tok('surface-1');
    var multi = o.datasets.length > 1;
    var valueAxis = {
      beginAtZero: true, stacked: !!o.stacked, max: o.max,
      grid: { color: grid }, border: { display: false },
      ticks: { color: text2, maxTicksLimit: 6, callback: function (v) { return (o.axisFmt || fmt.short)(v); } }
    };
    var catAxis = {
      stacked: !!o.stacked, grid: { display: false }, border: { color: grid },
      ticks: {
        color: text2, autoSkip: true, maxRotation: 0, autoSkipPadding: 8,
        callback: function (v) { return trunc(this.getLabelForValue(v), o.labelLen || (o.horizontal ? 24 : 12)); }
      }
    };
    var last = o.datasets.length - 1;
    return new Chart(canvas, {
      type: 'bar',
      data: {
        labels: o.labels,
        datasets: o.datasets.map(function (d, i) {
          return {
            label: d.label, data: d.data, backgroundColor: d.color, hoverBackgroundColor: d.color,
            // Stacked: viền trên 2px màu nền tạo khe giữa các đoạn; chỉ đoạn trên cùng bo góc
            borderColor: surface, borderWidth: o.stacked && i !== last ? { top: 2 } : 0,
            borderRadius: o.stacked ? (i === last ? { topLeft: 4, topRight: 4 } : 0) : 4,
            borderSkipped: o.stacked ? false : 'start',
            maxBarThickness: o.stacked ? 40 : 26, categoryPercentage: multi && !o.stacked ? 0.78 : 0.82, barPercentage: 0.9
          };
        })
      },
      options: {
        indexAxis: o.horizontal ? 'y' : 'x',
        interaction: { mode: 'index', intersect: false, axis: o.horizontal ? 'y' : 'x' },
        scales: o.horizontal ? { x: valueAxis, y: catAxis } : { x: catAxis, y: valueAxis },
        plugins: {
          legend: legend(multi),
          tooltip: Object.assign(tooltip(o.fmt || fmt.int), o.stacked ? {
            itemSort: function (a, b) { return b.datasetIndex - a.datasetIndex; },
            filter: function (it) { return it.raw > 0; },
            callbacks: Object.assign(tooltip(o.fmt || fmt.int).callbacks, {
              footer: function (items) {
                var s = items.reduce(function (a, it) { return a + (it.raw || 0); }, 0);
                return 'Tổng: ' + (o.fmt || fmt.int)(s);
              }
            })
          } : {})
        }
      }
    });
  }

  function line(canvas, o) {
    var text2 = tok('text-secondary'), grid = tok('grid'), surface = tok('surface-1');
    return new Chart(canvas, {
      type: 'line',
      data: {
        labels: o.labels,
        datasets: o.datasets.map(function (d) {
          return {
            label: d.label, data: d.data, borderColor: d.color, backgroundColor: d.color,
            borderWidth: 2, tension: 0.25, spanGaps: true,
            pointRadius: 4, pointHoverRadius: 6, pointBorderColor: surface, pointBorderWidth: 2,
            pointBackgroundColor: d.color
          };
        })
      },
      options: {
        interaction: { mode: 'index', intersect: false },
        scales: {
          x: { grid: { display: false }, border: { color: grid }, ticks: { color: text2, maxRotation: 0, autoSkipPadding: 8 } },
          y: {
            beginAtZero: true, max: o.max, grid: { color: grid }, border: { display: false },
            ticks: { color: text2, maxTicksLimit: 6, callback: function (v) { return (o.axisFmt || fmt.short)(v); } }
          }
        },
        plugins: { legend: legend(o.datasets.length > 1), tooltip: tooltip(o.fmt || fmt.int) }
      }
    });
  }

  // o: { labels, data, colors, fmt }
  function doughnut(canvas, o) {
    var total = o.data.reduce(function (a, b) { return a + b; }, 0) || 1;
    var t = tooltip(o.fmt || fmt.int);
    t.callbacks = {
      label: function (ctx) { return ' ' + (o.fmt || fmt.int)(ctx.parsed) + ' · ' + fmt.pct(ctx.parsed / total); }
    };
    var narrow = canvas.parentNode.clientWidth < 520;
    return new Chart(canvas, {
      type: 'doughnut',
      data: {
        labels: o.labels,
        datasets: [{ data: o.data, backgroundColor: o.colors, hoverBackgroundColor: o.colors, borderColor: tok('surface-1'), borderWidth: 2, hoverOffset: 6 }]
      },
      options: {
        cutout: '62%',
        layout: { padding: 6 },
        plugins: {
          legend: Object.assign(legend(true), {
            position: narrow ? 'bottom' : 'right', align: 'center',
            labels: Object.assign(legend(true).labels, {
              generateLabels: function (chart) {
                return Chart.overrides.doughnut.plugins.legend.labels.generateLabels(chart).map(function (l) {
                  l.text = trunc(l.text, 26) + '  ' + fmt.pct(o.data[l.index] / total);
                  l.fontColor = tok('text-secondary');
                  return l;
                });
              }
            })
          }),
          tooltip: t
        }
      }
    });
  }

  return { tok: tok, series: series, fmt: fmt, bar: bar, line: line, doughnut: doughnut };
})();
