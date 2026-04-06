/**
 * ════════════════════════════════════════════════════════════
 *  pcc-charts.js  —  PCC Weekly Material Monitoring System
 *  Chart Module  (Chart.js 4.x)
 *
 *  Depends on:
 *    • Chart.js loaded before this file
 *    • window.S      — shared app state (rows, week, sheet)
 *    • window.currentTheme — 'dark' | 'light'
 *    • DOM canvas IDs: chart-weekly, chart-donut, chart-dept, chart-profs
 *    • CSS variable --green, --red, etc. (set by the main stylesheet)
 *
 *  Public API (called by pcc-app.js):
 *    PccCharts.render()        — build/refresh all four charts
 *    PccCharts.destroyAll()    — tear down all chart instances
 * ════════════════════════════════════════════════════════════
 */

const PccCharts = (() => {

  /* ── internal chart-instance registry ─────────────────── */
  const _instances = {};

  function _destroy(key) {
    if (_instances[key]) {
      _instances[key].destroy();
      delete _instances[key];
    }
  }

  function destroyAll() {
    Object.keys(_instances).forEach(_destroy);
  }

  /* ── theme-aware colour palette ───────────────────────── */
  function _palette() {
    const dark = (window.currentTheme || 'dark') === 'dark';
    return {
      text:   dark ? '#96B4C8'              : '#2E4A6A',
      grid:   dark ? 'rgba(226,184,90,.07)' : 'rgba(26,86,219,.07)',
      gold:   dark ? '#E2B85A'              : '#1A56DB',
      green:  dark ? '#34D475'              : '#059669',
      red:    dark ? '#F05252'              : '#DC2626',
      blue:   dark ? '#4DA3FF'              : '#1D4ED8',
      orange: dark ? '#F59E0B'              : '#D97706',
      purple: dark ? '#B07DFA'              : '#7C3AED',
      teal:   dark ? '#2DD4BF'              : '#0D9488',
      seq: dark
        ? ['#E2B85A','#4DA3FF','#34D475','#F05252','#F59E0B','#B07DFA','#2DD4BF']
        : ['#1A56DB','#059669','#DC2626','#D97706','#7C3AED','#0D9488','#1D4ED8'],
    };
  }

  function _baseOptions(C) {
    return {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 500 },
      plugins: {
        legend: {
          labels: { color: C.text, font: { size: 13 }, boxWidth: 14, padding: 16 },
        },
      },
    };
  }

  function _ctx(id) {
    const el = document.getElementById(id);
    if (!el) { console.warn('[PccCharts] Canvas not found:', id); return null; }
    return el.getContext('2d');
  }

  function _hasData() {
    return window.S && Array.isArray(window.S.rows) && window.S.rows.length > 0;
  }

  function _totalWeeks() {
    return (window.TOTAL_WEEKS && window.TOTAL_WEEKS > 0) ? window.TOTAL_WEEKS : 18;
  }

  /* ══════════════════════════════════════════════════════════
     CHART 0 — Central Overview Donut  (overview panel)
  ══════════════════════════════════════════════════════════ */
  function _renderOverviewDonut() {
    _destroy('overview');
    const ctx = _ctx('chart-overview-donut');
    if (!ctx) return;

    const C    = _palette();
    const rows = window.S.rows;
    const TW   = _totalWeeks();

    const checked = rows.reduce((a, r) => a + r.done, 0);
    const missed  = rows.reduce((a, r) => a + r.miss, 0);
    const empty   = rows.reduce((a, r) => a + (TW - r.done - r.miss), 0);

    const total = checked + missed;
    const pct   = total > 0 ? Math.round(checked / total * 100) : 0;
    const rateEl = document.getElementById('s-rate');
    if (rateEl) rateEl.textContent = total > 0 ? pct + '%' : '—';

    const deptCenterEl = document.getElementById('s-dept-sub-center');
    if (deptCenterEl) deptCenterEl.textContent = window.S.sheet || '—';

    _instances['overview'] = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels: ['Complete', 'No Upload', 'Empty'],
        datasets: [{
          data: [checked, missed, empty],
          backgroundColor: [C.green + 'CC', C.red + 'CC', 'rgba(128,128,128,.22)'],
          borderColor:     [C.green,         C.red,         'rgba(128,128,128,.3)'],
          borderWidth: 2,
          hoverOffset: 12,
        }],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '70%',
        animation: { duration: 600 },
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: { label: c => c.label + ': ' + c.raw.toLocaleString() },
          },
        },
      },
    });
  }

  /* ══════════════════════════════════════════════════════════
     CHART 1 — Weekly Upload Progress  (line)
  ══════════════════════════════════════════════════════════ */
  function _renderWeekly() {
    _destroy('weekly');
    const ctx = _ctx('chart-weekly');
    if (!ctx) return;

    const C     = _palette();
    const rows  = window.S.rows;
    const total = rows.length;
    const TW    = _totalWeeks();

    const counts = Array.from({ length: TW }, (_, i) =>
      rows.filter(r => r.weeks[i + 1] === 'c').length
    );
    const pcts = counts.map(v => (total > 0 ? Math.round(v / total * 100) : 0));

    const base = _baseOptions(C);
    base.plugins.legend.position = 'top';

    _instances['weekly'] = new Chart(ctx, {
      type: 'line',
      data: {
        labels: Array.from({ length: TW }, (_, i) => 'W' + (i + 1)),
        datasets: [
          {
            label: 'Upload %',
            data: pcts,
            borderColor: C.gold,
            backgroundColor: C.gold + '18',
            pointBackgroundColor: C.gold,
            pointRadius: 4,
            pointHoverRadius: 7,
            tension: 0.4,
            fill: true,
            borderWidth: 2.5,
            yAxisID: 'y',
          },
          {
            label: 'Count',
            data: counts,
            borderColor: C.blue,
            backgroundColor: 'transparent',
            pointBackgroundColor: C.blue,
            pointRadius: 3,
            pointHoverRadius: 6,
            tension: 0.4,
            fill: false,
            borderWidth: 2,
            yAxisID: 'y2',
          },
        ],
      },
      options: {
        ...base,
        plugins: { legend: { display: false } },
        scales: {
          x: {
            ticks: { color: C.text, font: { size: 12 } },
            grid:  { color: C.grid },
          },
          y: {
            min: 0,
            max: 100,
            position: 'left',
            ticks: { color: C.gold, font: { size: 12 }, callback: v => v + '%' },
            grid:  { color: C.grid },
          },
          y2: {
            min: 0,
            position: 'right',
            ticks: { color: C.blue, font: { size: 12 } },
            grid:  { display: false },
          },
        },
      },
    });
  }

  /* ══════════════════════════════════════════════════════════
     CHART 2 — Status Breakdown  (doughnut by Status column)
  ══════════════════════════════════════════════════════════ */
  function _renderDonut() {
    _destroy('donut');
    const ctx = _ctx('chart-donut');
    if (!ctx) return;

    const C    = _palette();
    const rows = window.S.rows;

    // Count each status value
    const STATUS_OPTIONS = ['Complete', 'No Upload', 'Overdue', 'Revision', 'Pending'];
    const counts = {};
    STATUS_OPTIONS.forEach(s => counts[s] = 0);
    let hasStatus = false;
    rows.forEach(r => {
      const s = (r.status || '').trim();
      if (!s) return;
      hasStatus = true;
      if (counts[s] !== undefined) counts[s]++;
      else counts[s] = (counts[s] || 0) + 1;
    });

    // Fallback: if no status column data yet, show empty state
    const labels = STATUS_OPTIONS.filter(s => counts[s] > 0);
    const data   = labels.map(s => counts[s]);

    const COLORS = {
      'Complete':  C.green,
      'No Upload': C.red,
      'Overdue':   C.orange,
      'Revision':  C.purple,
      'Pending':   C.blue,
    };

    _instances['donut'] = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels,
        datasets: [{
          data,
          backgroundColor: labels.map(l => (COLORS[l] || C.text) + 'BB'),
          borderColor:     labels.map(l =>  COLORS[l] || C.text),
          borderWidth: 2,
          hoverOffset: 12,
        }],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        animation: { duration: 600 },
        cutout: '62%',
        plugins: {
          legend: {
            display: true,
            position: 'bottom',
            labels: {
              color: C.text,
              font: { size: 12, weight: '700' },
              padding: 14,
              usePointStyle: true,
              pointStyleWidth: 10,
            },
          },
          tooltip: {
            callbacks: {
              label: ctx => ' ' + ctx.label + ': ' + ctx.raw + ' rows',
            },
          },
        },
      },
    });
  }

  /* ══════════════════════════════════════════════════════════
     CHART 3 — Department Radar / Spider  (radar)
  ══════════════════════════════════════════════════════════ */
  function _renderDept() {
    _destroy('dept');
    const ctx = _ctx('chart-dept');
    if (!ctx) return;

    const C    = _palette();
    const rows = window.S.rows;
    const TW   = _totalWeeks();

    /* aggregate per-dept completion rates */
    const map = {};
    rows.forEach(r => {
      const d = (r.dept || 'OTHER').trim();
      if (!map[d]) map[d] = { done: 0, slots: 0 };
      map[d].done  += r.done;
      map[d].slots += TW;
    });

    const labels = Object.keys(map);
    const rates  = labels.map(d =>
      map[d].slots > 0 ? Math.round(map[d].done / map[d].slots * 100) : 0
    );

    /* need at least 3 points for a meaningful radar */
    const dark = (window.currentTheme || 'dark') === 'dark';

    _instances['dept'] = new Chart(ctx, {
      type: 'radar',
      data: {
        labels,
        datasets: [{
          label: 'Completion %',
          data: rates,
          backgroundColor: C.gold + '22',
          borderColor:     C.gold,
          borderWidth: 2.5,
          pointBackgroundColor: C.gold,
          pointBorderColor:     C.gold,
          pointHoverBackgroundColor: '#fff',
          pointRadius: 5,
          pointHoverRadius: 7,
        }],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        animation: { duration: 600 },
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: ctx => ctx.raw + '%' } },
        },
        scales: {
          r: {
            min: 0,
            max: 100,
            ticks: {
              stepSize: 20,
              color: C.text,
              font: { size: 11 },
              backdropColor: 'transparent',
              callback: v => v + '%',
            },
            grid:         { color: C.grid },
            angleLines:   { color: C.grid },
            pointLabels:  {
              color: C.text,
              font:  { size: 12, weight: '700' },
            },
          },
        },
      },
    });
  }

  /* ══════════════════════════════════════════════════════════
     CHART 4 — Top 10 Professors  (horizontal bar)
  ══════════════════════════════════════════════════════════ */
  function _renderProfs() {
    _destroy('profs');
    const ctx = _ctx('chart-profs');
    if (!ctx) return;

    const C    = _palette();
    const rows = window.S.rows;
    const TW   = _totalWeeks();

    const top10 = [...rows]
      .sort((a, b) => b.done - a.done)
      .slice(0, 10);

    const labels = top10.map(r => {
      const parts = r.professor.trim().split(/\s+/);
      return parts.length > 1
        ? parts[parts.length - 1] + ', ' + parts[0][0] + '.'
        : r.professor;
    });

    const data = top10.map(r => Math.round(r.done / TW * 100));

    const bgColors     = data.map(v => v >= 80 ? C.green + '88' : v >= 50 ? C.gold + '88' : C.red + '88');
    const borderColors = data.map(v => v >= 80 ? C.green          : v >= 50 ? C.gold          : C.red);

    const base = _baseOptions(C);
    base.plugins.legend = { display: false };

    _instances['profs'] = new Chart(ctx, {
      type: 'bar',
      data: {
        labels,
        datasets: [{
          label: 'Completion %',
          data,
          backgroundColor: bgColors,
          borderColor:     borderColors,
          borderWidth: 2,
          borderRadius: 5,
        }],
      },
      options: {
        ...base,
        indexAxis: 'y',
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: ctx => ctx.raw + '%' } },
        },
        scales: {
          x: {
            min: 0,
            max: 100,
            ticks: { color: C.text, font: { size: 12 }, callback: v => v + '%' },
            grid:  { color: C.grid },
          },
          y: {
            ticks: { color: C.text, font: { size: 12 } },
            grid:  { display: false },
          },
        },
      },
    });
  }

  /* ══════════════════════════════════════════════════════════
     PUBLIC: render()
  ══════════════════════════════════════════════════════════ */
  function render() {
    const section = document.getElementById('charts-section');

    if (!_hasData()) {
      if (section) section.style.display = 'none';
      return;
    }

    if (section) section.style.display = 'block';

    destroyAll();

    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        _renderOverviewDonut();
        _renderWeekly();
        _renderDonut();
        _renderProfs();
      });
    });
  }

  return { render, destroyAll };

})();