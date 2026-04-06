/**
 * ════════════════════════════════════════════════════════════
 *  pcc-app.js  —  PCC Weekly Material Monitoring System
 *  Main Application Logic — Mobile-Responsive Edition
 * ════════════════════════════════════════════════════════════
 */

const CLIENT_ID      = '707446801798-i0s4ij075s4tlp9lqbk4emkrdsfml23k.apps.googleusercontent.com';
const SPREADSHEET_ID = '1MTaFaaO5pxFe8fXYrxyXjs4MWRSTSLru2btMYIdcDQw';
const SCOPES         = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/userinfo.email openid email profile';

const WEBAPP_URL = 'https://script.google.com/a/macros/paranaquecitycollege.edu.ph/s/AKfycbys7x3r-ErMn_4-ssxoUwUWQj3lkNDfW5vu_IqmxLOOJJlneygO6OjaQTygsscq_U5u/exec';

const ADMIN_EMAILS = [
  'dlc@paranaquecitycollege.edu.ph',
  'office.accer@paranaquecitycollege.edu.ph',
];



const DEPT_ACCESS = {};

const VIEWER_ACCESS = {
  
  'edrosolan.bhonmark@paranaquecitycollege.edu.ph':     { sheets: ['BSTM DASHBOARD'], deptKey: 'BSTM'  },
  'ocan.gerald@paranaquecitycollege.edu.ph':        { sheets: ['BSE DASHBOARD'],  deptKey: 'BSE'   },
  'salcedo.elcid@paranaquecitycollege.edu.ph':      { sheets: ['BPA DASHBOARD'],  deptKey: 'BPA'   },
  'lubang.philipp@paranaquecitycollege.edu.ph':     { sheets: ['ITVET DASHBOARD'],deptKey: 'ITVET' },
  'ang.jaime@paranaquecitycollege.edu.ph':          { sheets: ['GENED DASHBOARD'],deptKey: 'GENED' },
  'bandal.markanthony@paranaquecitycollege.edu.ph': { sheets: ['BSREM DASHBOARD'],deptKey: 'BSREM' },
};

const ALL_DEPTS = [
  { label: 'BSTM',    sheet: 'BSTM DASHBOARD',  key: 'bstm'   },
  { label: 'BSE',     sheet: 'BSE DASHBOARD',   key: 'bse'    },
  { label: 'BSREM',   sheet: 'BSREM DASHBOARD', key: 'bsrem'  },
  { label: 'BPA',     sheet: 'BPA DASHBOARD',   key: 'bpa'    },
  { label: 'NON-ABM', sheet: 'NON-ABM',         key: 'nonabm' },
  { label: 'ITVET',   sheet: 'ITVET DASHBOARD', key: 'itvet'  },
  { label: 'GENED',   sheet: 'GENED DASHBOARD', key: 'gened'  },
];

const DATA_START  = 55;
const TOTAL_WEEKS = 18;
const I_PROF    = 0;
const I_COURSE  = 1;
const I_SECTION = 2;
const I_URL     = 3;
const I_W1      = 5;

window.S = {
  token: '',
  sheet: '',
  rows: [],
  week: 0,
  role: '',
  allowedSheets: [],
  userEmail: '',
  colMap: {},
};
window.currentTheme = localStorage.getItem('pcc-theme') || 'dark';
window.TOTAL_WEEKS  = TOTAL_WEEKS;

const EXTRA_FIELDS = [
  { key: 'reflection',        labels: ['reflection'] },
  { key: 'quiz',              labels: ['quiz'] },
  { key: 'assignment',        labels: ['assignment'] },
  { key: 'researchTask',      labels: ['research task', 'research'] },
  { key: 'recitation',        labels: ['recitation'] },
  { key: 'caseStudy',         labels: ['case study'] },
  { key: 'activity',          labels: ['activity'] },
  { key: 'videoPresentation', labels: ['video presentation', 'video'] },
  { key: 'oba',               labels: ['oba'] },
  { key: 'withProfessors',    labels: ['with professor', 'with professors', 'w/ professor'] },
  { key: 'withStudents',      labels: ['with student', 'with students', 'w/ student'] },
  { key: 'status',            labels: ['status', 'upload status', 'current status', 'remarks'] },
];

/* ── Mobile helpers ── */
function _isMobile()  { return window.innerWidth <= 480; }
function _isTablet()  { return window.innerWidth <= 768; }

/* ══════════════════════════════════════════════════════════════
   THEME
══════════════════════════════════════════════════════════════ */
function applyTheme(t) {
  window.currentTheme = t;
  document.documentElement.setAttribute('data-theme', t);
  const isDark = t === 'dark';
  const lbl = document.getElementById('theme-lbl');
  if (lbl) lbl.textContent = isDark ? '🌙 Dark' : '☀️ Light';
  const icon = document.getElementById('toggle-icon');
  if (icon) icon.textContent = isDark ? '🌙' : '☀️';
  localStorage.setItem('pcc-theme', t);
  if (window.S.rows.length) PccCharts.render();
}

function toggleTheme() {
  applyTheme(window.currentTheme === 'dark' ? 'light' : 'dark');
}

applyTheme(window.currentTheme);

/* ══════════════════════════════════════════════════════════════
   CLOCK
══════════════════════════════════════════════════════════════ */
function tick() {
  const n = new Date();
  const mob = _isMobile();
  document.getElementById('hdate').textContent = mob
    ? n.toLocaleDateString('en-PH', { month: 'short', day: 'numeric', year: 'numeric' })
    : n.toLocaleDateString('en-PH', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
  document.getElementById('htime').textContent =
    n.toLocaleTimeString('en-PH', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
}
tick();
setInterval(tick, 1000);

// Re-run tick on resize to adjust date format
window.addEventListener('resize', tick);

/* ══════════════════════════════════════════════════════════════
   SMART SCROLL HEADER
   • Scroll DOWN  → header slides up out of view (like Facebook)
   • Scroll UP    → header slides back into view immediately
   • At very top  → header always visible
══════════════════════════════════════════════════════════════ */
(function () {
  const hdr        = document.querySelector('header');
  const THRESHOLD  = 8;   // px — ignore tiny jitter scrolls
  const SHOW_ABOVE = 60;  // px from top — always show header near top

  let lastY    = 0;
  let ticking  = false;

  function onScroll() {
    if (!ticking) {
      requestAnimationFrame(() => {
        const currentY = window.scrollY;
        const diff     = currentY - lastY;

        if (currentY <= SHOW_ABOVE) {
          // Near the very top — always show
          hdr.classList.remove('header-hidden');
          hdr.classList.add('header-visible');
        } else if (diff > THRESHOLD) {
          // Scrolling DOWN — hide
          hdr.classList.add('header-hidden');
          hdr.classList.remove('header-visible');
        } else if (diff < -THRESHOLD) {
          // Scrolling UP — show
          hdr.classList.remove('header-hidden');
          hdr.classList.add('header-visible');
        }

        lastY    = currentY;
        ticking  = false;
      });
      ticking = true;
    }
  }

  window.addEventListener('scroll', onScroll, { passive: true });
})();

/* ══════════════════════════════════════════════════════════════
   ACCESS CONTROL
══════════════════════════════════════════════════════════════ */
function checkAccess(email) {
  const em = email.toLowerCase().trim();
  if (ADMIN_EMAILS.map(e => e.toLowerCase()).includes(em))
    return { role: 'admin',  allowedSheets: ALL_DEPTS.map(d => d.sheet) };
  if (DEPT_ACCESS[em])
    return { role: 'dept',   allowedSheets: DEPT_ACCESS[em].sheets,   deptKey: DEPT_ACCESS[em].deptKey };
  if (VIEWER_ACCESS[em])
    return { role: 'viewer', allowedSheets: VIEWER_ACCESS[em].sheets, deptKey: VIEWER_ACCESS[em].deptKey };
  return { role: 'denied', allowedSheets: [] };
}

/* Shortcut — true if current user can edit */
function _canEdit() {
  return window.S.role === 'admin' || window.S.role === 'dept';
}

/* ══════════════════════════════════════════════════════════════
   HEADER DEPT QUICK-SWITCHER
══════════════════════════════════════════════════════════════ */
const DEPT_COLORS = {
  bstm:   '#E2B85A',
  bse:    '#4DA3FF',
  bsrem:  '#34D475',
  bpa:    '#F05252',
  nonabm: '#F59E0B',
  itvet:  '#B07DFA',
  gened:  '#2DD4BF',
};

function buildDeptSwitcher() {
  const wrap = document.getElementById('dept-switcher-wrap');
  const menu = document.getElementById('dept-switcher-menu');
  const lbl  = document.getElementById('dsb-label');
  if (!wrap || !menu) return;

  // Show the switcher correctly
  wrap.style.display = 'flex';
  wrap.style.alignItems = 'center';

  // Build menu items
  const items = ALL_DEPTS.map(d => {
    const allowed  = window.S.allowedSheets.includes(d.sheet);
    const isActive = d.sheet === window.S.sheet;
    const color    = DEPT_COLORS[d.key] || 'var(--t3)';
    return `
      <button
        class="dsm-item ${isActive ? 'active' : ''} ${!allowed ? 'locked' : ''}"
        onclick="switchDeptFromHeader('${d.sheet}','${d.label}')"
        ${!allowed ? 'disabled' : ''}
        title="${!allowed ? 'Access restricted' : 'Switch to ' + d.label}"
      >
        <span class="dsm-dot" style="background:${color}"></span>
        ${d.label}
        ${!allowed ? ' 🔒' : ''}
        ${isActive ? ' ✓' : ''}
      </button>`;
  }).join('');

  // Find header of existing menu content, replace items only
  const header = menu.querySelector('.dsm-header');
  menu.innerHTML = '';
  if (header) menu.appendChild(header);
  else menu.innerHTML = '<div class="dsm-header">Switch Department</div>';
  menu.querySelector('.dsm-header').insertAdjacentHTML('afterend', items);

  // Update button label to current dept
  const current = ALL_DEPTS.find(d => d.sheet === window.S.sheet);
  if (lbl && current) lbl.textContent = current.label;
}

function toggleDeptSwitcher(e) {
  e.stopPropagation();
  const wrap = document.getElementById('dept-switcher-wrap');
  if (!wrap) return;
  wrap.classList.toggle('open');
}

function switchDeptFromHeader(sheet, label) {
  // Close dropdown
  const wrap = document.getElementById('dept-switcher-wrap');
  if (wrap) wrap.classList.remove('open');

  // Update label
  const lbl = document.getElementById('dsb-label');
  if (lbl) lbl.textContent = label;

  // Sync with the main dept tab buttons
  if (!window.S.allowedSheets.includes(sheet)) { toast('Access denied', 'err'); return; }
  document.querySelectorAll('.dept-tab').forEach(t => {
    t.classList.toggle('active', t.getAttribute('data-sheet') === sheet);
  });
  window.S.sheet = sheet;
  if (window.S.token) loadSheet(sheet);

  // Refresh switcher to show new active state
  buildDeptSwitcher();
}

// Close dropdown when clicking anywhere outside
document.addEventListener('click', () => {
  const wrap = document.getElementById('dept-switcher-wrap');
  if (wrap) wrap.classList.remove('open');
});

/* ══════════════════════════════════════════════════════════════
   DEPT TABS
══════════════════════════════════════════════════════════════ */
function buildDeptTabs() {
  const container = document.getElementById('dept-tabs-container');
  container.innerHTML = '';
  ALL_DEPTS.forEach(d => {
    const btn     = document.createElement('button');
    const allowed = window.S.allowedSheets.includes(d.sheet);
    btn.className = 'dept-tab';
    if (!allowed) {
      btn.classList.add('locked');
      btn.disabled    = true;
      btn.textContent = d.label;
      btn.title       = 'Access restricted';
    } else {
      btn.textContent = d.label;
      btn.setAttribute('data-sheet', d.sheet);
      btn.onclick = () => switchDept(btn);
    }
    if (d.sheet === window.S.sheet) btn.classList.add('active');
    container.appendChild(btn);
  });
}

/* ══════════════════════════════════════════════════════════════
   SIGN IN / OUT
══════════════════════════════════════════════════════════════ */
let _tokenClient;

function doSignIn() {
  _tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: SCOPES,
    callback: async (resp) => {
      if (resp.error) { toast('Sign-in failed: ' + resp.error, 'err'); return; }
      window.S.token = resp.access_token;
      try {
        const uiRes = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
          headers: { Authorization: 'Bearer ' + window.S.token },
        });
        const ud = await uiRes.json();
        let emailRaw = ud.email || '';

        if (!emailRaw) {
          try {
            const tiRes = await fetch('https://oauth2.googleapis.com/tokeninfo?access_token=' + window.S.token);
            const td    = await tiRes.json();
            emailRaw    = td.email || '';
          } catch (_) {}
        }

        if (!emailRaw && resp.id_token) {
          try {
            const payload = JSON.parse(
              atob(resp.id_token.split('.')[1].replace(/-/g, '+').replace(/_/g, '/'))
            );
            emailRaw = payload.email || '';
          } catch (_) {}
        }

        window.S.userEmail = emailRaw.toLowerCase().trim();
        const access       = checkAccess(window.S.userEmail);

        if (access.role === 'denied') {
          document.getElementById('signin-overlay').style.display        = 'none';
          document.getElementById('access-denied-overlay').style.display = 'flex';
          document.getElementById('denied-email-display').textContent    = window.S.userEmail;
          document.title = 'DENIED: ' + window.S.userEmail;
          document.getElementById('allowed-domains-display').innerHTML =
            'Authorized accounts:<br><br>' +
            Object.entries(DEPT_ACCESS).map(([em, info]) =>
              `• <span>${em}</span> — ${info.deptKey}`
            ).join('<br>');
          window.S.token = '';
          return;
        }

        window.S.role          = access.role;
        window.S.allowedSheets = access.allowedSheets;

        // Truncate name/email on mobile
        const displayName = ud.name || window.S.userEmail;
        const displayEmail = ud.email || window.S.userEmail;
        document.getElementById('user-info').textContent =
          _isMobile() ? displayName.split(' ')[0] : displayName + ' (' + displayEmail + ')';

        document.getElementById('role-display').innerHTML =
          window.S.role === 'admin'
            ? '<span class="role-badge role-admin">⭐ Admin</span>'
            : window.S.role === 'viewer'
            ? '<span class="role-badge role-viewer">👁 Viewer — ' + (access.deptKey || 'DEPT') + '</span>'
            : '<span class="role-badge role-dept">🏬 ' + (access.deptKey || 'DEPT') + '</span>';

        window.S.sheet = window.S.allowedSheets[0] || '';
        document.getElementById('signin-overlay').style.display   = 'none';
        document.getElementById('sync-indicator').style.display   = '';

        buildDeptTabs();
        buildDeptSwitcher();
        buildWeekTabs();
        if (window.S.sheet) loadSheet(window.S.sheet);
        toast('✓ Access granted', 'ok');

      } catch (e) {
        toast('Error: ' + e.message, 'err');
        window.S.token = '';
      }
    },
  });
  _tokenClient.requestAccessToken({ prompt: 'consent' });
}

function signOut() {
  if (window.S.token) google.accounts.oauth2.revoke(window.S.token, () => {});
  window.S.token = '';
  window.S.rows  = [];
  window.S.role  = '';
  window.S.allowedSheets = [];
  window.S.userEmail     = '';

  PccCharts.destroyAll();

  document.getElementById('user-info').textContent               = '—';
  document.getElementById('role-display').innerHTML              = '';
  document.getElementById('btn-refresh').style.display           = 'none';
  document.getElementById('sync-indicator').style.display        = 'none';
  document.getElementById('charts-section').style.display        = 'none';
  document.getElementById('btn-announce').style.display         = 'none'; 
  document.getElementById('access-denied-overlay').style.display = 'none';
  document.getElementById('signin-overlay').style.display        = 'flex';
  const dsw = document.getElementById('dept-switcher-wrap');
  if (dsw) dsw.style.display = 'none';
  document.getElementById('table-container').innerHTML =
    `<div class="loader"><div class="spin"></div>Sign in to load data…</div>`;
}

/* ══════════════════════════════════════════════════════════════
   WEEK TABS
══════════════════════════════════════════════════════════════ */
function buildWeekTabs() {
  const container = document.getElementById('week-tabs');
  let h = `<button class="week-tab ${window.S.week === 0 ? 'active' : ''}" onclick="selWeek(0,this)">ALL</button>`;
  for (let w = 1; w <= TOTAL_WEEKS; w++)
    h += `<button class="week-tab ${window.S.week === w ? 'active' : ''}" onclick="selWeek(${w},this)">W${w}</button>`;
  container.innerHTML = h;
  document.getElementById('s-week').textContent = window.S.week === 0 ? 'ALL' : window.S.week;
}

function selWeek(w, el) {
  window.S.week = w;
  document.querySelectorAll('.week-tab').forEach(t => t.classList.remove('active'));
  el.classList.add('active');
  document.getElementById('s-week').textContent = w === 0 ? 'ALL' : w;
  renderTable();
}

/* ══════════════════════════════════════════════════════════════
   DEPT SWITCH / RELOAD
══════════════════════════════════════════════════════════════ */
function switchDept(el) {
  const sheet = el.getAttribute('data-sheet');
  if (!window.S.allowedSheets.includes(sheet)) { toast('Access denied', 'err'); return; }
  document.querySelectorAll('.dept-tab').forEach(t => t.classList.remove('active'));
  el.classList.add('active');
  window.S.sheet = sheet;
  if (window.S.token) loadSheet(sheet);
}

function reloadCurrent() {
  if (window.S.token && window.S.sheet) loadSheet(window.S.sheet);
}

/* ══════════════════════════════════════════════════════════════
   LOAD SHEET
══════════════════════════════════════════════════════════════ */
async function loadSheet(sheetName) {
  setStatus('Loading ' + sheetName + '…');
  document.getElementById('table-container').innerHTML =
    `<div class="loader"><div class="spin"></div>Loading ${sheetName}…</div>`;

  try {
    const hdrRange = encodeURIComponent("'" + sheetName + "'!A54:AZ54");
    const hdrUrl   = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${hdrRange}` +
                     `?valueRenderOption=FORMATTED_VALUE&majorDimension=ROWS`;
    const hdrRes   = await fetch(hdrUrl, { headers: { Authorization: 'Bearer ' + window.S.token } });
    if (hdrRes.status === 401) { signOut(); toast('Session expired', 'warn'); return; }
    const hdrJson  = await hdrRes.json();
    const hdrRow   = (hdrJson.values || [[]])[0] || [];

    window.S.colMap = {};
    hdrRow.forEach((cell, idx) => {
      const norm = String(cell).trim().toLowerCase();
      if (norm) window.S.colMap[norm] = idx;
    });

    window.S.extraIdx = {};
    EXTRA_FIELDS.forEach(({ key, labels }) => {
      for (const lbl of labels) {
        if (window.S.colMap[lbl] !== undefined) {
          window.S.extraIdx[key] = window.S.colMap[lbl];
          break;
        }
      }
    });

    console.log('[PCC] Column map detected:', window.S.colMap);
    console.log('[PCC] Extra column indices:', window.S.extraIdx);

    const range = encodeURIComponent("'" + sheetName + "'!A" + DATA_START + ':AZ');
    const url   = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${range}` +
                  `?valueRenderOption=UNFORMATTED_VALUE&majorDimension=ROWS`;
    const res   = await fetch(url, { headers: { Authorization: 'Bearer ' + window.S.token } });

    if (res.status === 401) { signOut(); toast('Session expired', 'warn'); return; }
    if (!res.ok) {
      const errBody = await res.json();
      throw new Error(errBody.error?.message || 'HTTP ' + res.status);
    }

    const json  = await res.json();
    window.S.rows = (json.values || [])
      .map((r, i) => _parseRow(r, sheetName, DATA_START + i))
      .filter(r => r && r.professor.trim());

    document.getElementById('btn-refresh').style.display    = '';
    document.getElementById('last-updated').textContent     = 'Updated: ' + new Date().toLocaleTimeString('en-PH');
    setStatus('✓ ' + window.S.rows.length + ' rows loaded from "' + sheetName + '"');

    buildWeekTabs();
    renderTable();
    toast('✓ ' + window.S.rows.length + ' rows from ' + sheetName, 'ok');

  } catch (e) {
    setStatus('Error: ' + e.message, true);
    toast('Error: ' + e.message, 'err');
    document.getElementById('table-container').innerHTML =
      `<div class="empty"><span class="eicon">⚠️</span><h3>Failed to Load</h3><p>${esc(e.message)}</p></div>`;
  }
}

/* ══════════════════════════════════════════════════════════════
   PARSE ROW
══════════════════════════════════════════════════════════════ */
function _parseRow(r, sheetName, sheetRow) {
  const prof   = String(r[I_PROF]   || '').trim();
  const course = String(r[I_COURSE] || '').trim();

  // Skip empty rows
  if (!prof) return null;

  // Skip footer/summary rows:
  // 1. Course contains ***NOTHING FOLLOWS*** (any asterisk variant)
  // 2. Professor is a plain number (e.g. "2", "15" — count rows)
  // 3. Professor itself contains NOTHING FOLLOWS
  if (/\*+\s*NOTHING\s+FOLLOWS\s*\*+/i.test(course))  return null;
  if (/\*+\s*NOTHING\s+FOLLOWS\s*\*+/i.test(prof))    return null;
  if (/^\d+$/.test(prof))                               return null;

  const weeks = {};
  for (let w = 1; w <= TOTAL_WEEKS; w++) {
    const cell = r[I_W1 + w - 1];
    const raw  = String(cell === undefined || cell === null || cell === '' ? '' : cell).trim().toUpperCase();
    weeks[w] = (raw === 'TRUE'  || raw === '1' || raw === 'COMPLETE')    ? 'c'
             : (raw === 'FALSE' || raw === '0' || raw === 'NO UPLOAD')   ? 'n'
             : '';
  }

  const done = Object.values(weeks).filter(v => v === 'c').length;
  const miss = Object.values(weeks).filter(v => v === 'n').length;
  const dept = sheetName.replace(/\s*DASHBOARD\s*/i, '').trim();

  const ei = window.S.extraIdx || {};
  const _x = (key) => {
    const idx = ei[key];
    return idx !== undefined ? String(r[idx] || '').trim() : '';
  };

  return {
    professor:         prof,
    course:            String(r[I_COURSE]  || '').trim(),
    section:           String(r[I_SECTION] || '').trim(),
    url:               String(r[I_URL]     || '').trim(),
    dept,
    weeks,
    done,
    miss,
    sheetRow,
    reflection:        _x('reflection'),
    quiz:              _x('quiz'),
    assignment:        _x('assignment'),
    researchTask:      _x('researchTask'),
    recitation:        _x('recitation'),
    caseStudy:         _x('caseStudy'),
    activity:          _x('activity'),
    videoPresentation: _x('videoPresentation'),
    oba:               _x('oba'),
    withProfessors:    _x('withProfessors'),
    withStudents:      _x('withStudents'),
    status:            _x('status'),
  };
}

/* ══════════════════════════════════════════════════════════════
   WRITE CELL — Direct Sheets API v4
══════════════════════════════════════════════════════════════ */
async function _writeCell(sheetName, row1based, col1based, value) {
  if (!window.S.token) throw new Error('Not signed in — no OAuth token available');

  let writeVal;
  if      (value === true  || value === 1  || value === '1')  writeVal = 1;
  else if (value === false || value === 0  || value === '0')  writeVal = 0;
  else                                                         writeVal = value;

  const colLetter   = _colToLetter(col1based);
  const cellA1      = "'" + sheetName + "'!" + colLetter + row1based;
  const inputOption = (typeof writeVal === 'number') ? 'RAW' : 'USER_ENTERED';
  const encoded     = encodeURIComponent(cellA1);
  const url         = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encoded}` +
                      `?valueInputOption=${inputOption}`;

  const res = await fetch(url, {
    method:  'PUT',
    headers: {
      'Authorization': 'Bearer ' + window.S.token,
      'Content-Type':  'application/json',
    },
    body: JSON.stringify({
      range:          cellA1,
      majorDimension: 'ROWS',
      values:         [[writeVal]],
    }),
  });

  if (res.status === 401) { signOut(); throw new Error('Session expired. Please sign in again.'); }
  if (!res.ok) {
    let errMsg = 'HTTP ' + res.status;
    try { const e = await res.json(); errMsg = e.error?.message || errMsg; } catch (_) {}
    throw new Error(errMsg);
  }
  return await res.json();
}

function _colToLetter(col) {
  let s = '';
  while (col > 0) {
    col--;
    s   = String.fromCharCode(65 + (col % 26)) + s;
    col = Math.floor(col / 26);
  }
  return s;
}

/* ══════════════════════════════════════════════════════════════
   TOGGLE WEEK DOT
══════════════════════════════════════════════════════════════ */
async function toggleWeek(dotEl, dataIndex, weekNum) {
  if (!_canEdit()) {
    toast('👁 View-only access — editing is disabled', 'warn');
    return;
  }
  const row = window.S.rows[dataIndex];
  if (!row) return;

  const currentVal = row.weeks[weekNum];
  let newState, sheetValue;
  if      (currentVal === 'c') { newState = 'n'; sheetValue = 0;  }
  else if (currentVal === 'n') { newState = '';  sheetValue = ''; }
  else                          { newState = 'c'; sheetValue = 1;  }

  row.weeks[weekNum] = newState;
  row.done = Object.values(row.weeks).filter(v => v === 'c').length;
  row.miss = Object.values(row.weeks).filter(v => v === 'n').length;

  dotEl.className   = 'wdot ' + (newState || 'e') + ' saving';
  dotEl.textContent = newState === 'c' ? '✓' : newState === 'n' ? '✗' : '·';
  updateStats();
  _setSyncStatus('saving', 'Saving…');

  try {
    const col1based = I_W1 + weekNum;
    await _writeCell(window.S.sheet, row.sheetRow, col1based, sheetValue);

    dotEl.classList.remove('saving');
    _setSyncStatus('saved', 'Saved ✓');
    setTimeout(() => _setSyncStatus('', ''), 2500);

    const msg = newState === 'c' ? '✓ Marked complete'
              : newState === 'n' ? '✗ Marked no upload'
              : '· Cleared';
    toast(msg, 'ok');

  } catch (e) {
    row.weeks[weekNum] = currentVal;
    row.done = Object.values(row.weeks).filter(v => v === 'c').length;
    row.miss = Object.values(row.weeks).filter(v => v === 'n').length;

    dotEl.className   = 'wdot ' + (currentVal || 'e');
    dotEl.textContent = currentVal === 'c' ? '✓' : currentVal === 'n' ? '✗' : '·';
    dotEl.classList.remove('saving');
    updateStats();
    _setSyncStatus('error', 'Save failed');
    setTimeout(() => _setSyncStatus('', ''), 3000);
    toast('Save failed: ' + e.message, 'err');
  }
}

/* ══════════════════════════════════════════════════════════════
   STATUS DROPDOWN
══════════════════════════════════════════════════════════════ */
async function onStatusChange(selectEl, dataIndex) {
  if (!_canEdit()) {
    selectEl.value = window.S.rows[dataIndex]?.status || '';
    toast('👁 View-only access — editing is disabled', 'warn');
    return;
  }
  const row = window.S.rows[dataIndex];
  if (!row) return;

  const newVal = selectEl.value;
  const oldVal = row.status || '';
  const colIdx = (window.S.extraIdx || {})['status'];

  if (colIdx === undefined) {
    const found = Object.keys(window.S.colMap).join(', ');
    toast('❌ Status column not found. Headers: ' + found, 'err');
    return;
  }

  row.status = newVal;
  _applyStatusStyle(selectEl, newVal);
  _setSyncStatus('saving', 'Saving…');
  selectEl.disabled = true;

  try {
    await _writeCell(window.S.sheet, row.sheetRow, colIdx + 1, newVal);
    selectEl.disabled = false;
    _setSyncStatus('saved', 'Saved ✓');
    setTimeout(() => _setSyncStatus('', ''), 2500);
    toast('✅ Status: ' + (newVal || '—'), 'ok');
    PccCharts.render();
  } catch (e) {
    row.status = oldVal;
    selectEl.value = oldVal;
    _applyStatusStyle(selectEl, oldVal);
    selectEl.disabled = false;
    _setSyncStatus('error', 'Save failed');
    setTimeout(() => _setSyncStatus('', ''), 3000);
    toast('❌ Save failed: ' + e.message, 'err');
  }
}

/* ══════════════════════════════════════════════════════════════
   ACTIVITY CELL EDIT
══════════════════════════════════════════════════════════════ */
async function onActivityChange(inputEl, dataIndex, colKey) {
  if (!_canEdit()) {
    inputEl.value = inputEl.getAttribute('data-old-val') || '';
    toast('👁 View-only access — editing is disabled', 'warn');
    return;
  }
  const row = window.S.rows[dataIndex];
  if (!row) return;

  const newVal = inputEl.value.trim();
  const oldVal = inputEl.getAttribute('data-old-val') || '';
  const colIdx = (window.S.extraIdx || {})[colKey];

  if (colIdx === undefined) {
    toast('Column "' + colKey + '" not found', 'err');
    inputEl.value = oldVal;
    return;
  }

  if (newVal === oldVal) return;

  const writeVal = (newVal === '' || newVal === null) ? '' : Number(newVal);

  row[colKey] = newVal;
  inputEl.setAttribute('data-old-val', newVal);
  inputEl.className = 'activity-input ' + (newVal !== '' ? 'has-val' : 'empty-val');

  _setSyncStatus('saving', 'Saving…');
  inputEl.disabled = true;

  try {
    await _writeCell(window.S.sheet, row.sheetRow, colIdx + 1, writeVal);
    inputEl.disabled = false;
    _setSyncStatus('saved', 'Saved ✓');
    setTimeout(() => _setSyncStatus('', ''), 2500);
    toast('✅ ' + colKey + ': ' + (newVal || '—'), 'ok');
  } catch (e) {
    row[colKey] = oldVal;
    inputEl.value = oldVal;
    inputEl.setAttribute('data-old-val', oldVal);
    inputEl.className = 'activity-input ' + (oldVal !== '' ? 'has-val' : 'empty-val');
    inputEl.disabled = false;
    _setSyncStatus('error', 'Save failed');
    setTimeout(() => _setSyncStatus('', ''), 3000);
    toast('❌ Save failed: ' + e.message, 'err');
  }
}

function _applyStatusStyle(sel, val) {
  sel.className = 'status-dd ' + (
    val === 'Complete'  ? 'sd-complete'  :
    val === 'No Upload' ? 'sd-noupload'  :
    val === 'Overdue'   ? 'sd-overdue'   :
    val === 'Revision'  ? 'sd-revision'  :
    val === 'Pending'   ? 'sd-pending'   : 'sd-empty'
  );
}

function _setSyncStatus(type, label) {
  document.getElementById('sync-dot').className     = 'sync-dot' + (type ? ' ' + type : '');
  document.getElementById('sync-label').textContent = label;
  document.getElementById('sync-indicator').style.display = label ? '' : 'none';
}

/* ══════════════════════════════════════════════════════════════
   RENDER TABLE
   On mobile: hides some less-critical columns automatically
══════════════════════════════════════════════════════════════ */
function renderTable() {
  updateStats();

  // Update legend hint based on role
  const hint = document.getElementById('legend-hint');
  if (hint) {
    hint.innerHTML = _canEdit()
      ? '💡 Click any dot to toggle &amp; sync to sheet · ✏️ Click activity numbers to edit'
      : '👁 View-only mode — you can browse but not edit';
  }

  const sf  = document.getElementById('f-status').value;
  const q   = (document.getElementById('f-search').value || '').toLowerCase();
  const w   = window.S.week;
  const mob = _isMobile();
  const tab = _isTablet();

  const filtered = window.S.rows
    .map((r, i) => ({ r, i }))
    .filter(({ r }) => {
      if (!r.professor) return false;
      if (q && !r.professor.toLowerCase().includes(q) &&
               !r.course.toLowerCase().includes(q)    &&
               !r.section.toLowerCase().includes(q))  return false;
      if (sf) {
        if (w > 0) { if (r.weeks[w] !== sf) return false; }
        else       {
          if (sf === 'c' && r.done === 0) return false;
          if (sf === 'n' && r.miss === 0) return false;
        }
      }
      return true;
    });

  document.getElementById('row-count').textContent =
    filtered.length + ' entr' + (filtered.length === 1 ? 'y' : 'ies');
  document.getElementById('table-title').textContent =
    '📋 ' + window.S.sheet + (mob ? '' : ' — ' + (w === 0 ? 'All Weeks' : 'Week ' + w)) +
    ' (' + filtered.length + ')';

  if (!filtered.length) {
    document.getElementById('table-container').innerHTML =
      `<div class="empty"><span class="eicon">🔍</span><h3>No Results</h3><p>No entries match your current filters.</p></div>`;
    PccCharts.render();
    return;
  }

  const wcols = w === 0
    ? Array.from({ length: TOTAL_WEEKS }, (_, i) => i + 1)
    : [w];

  // On mobile, hide "Dept" and "Link" columns; keep essential ones
  const showDept    = !mob;
  const showSection = !mob;
  const showLink    = true;

  const EXTRA_COLS = [
    { key: 'reflection',        label: 'Refl.' },
    { key: 'quiz',              label: 'Quiz' },
    { key: 'assignment',        label: 'Assign.' },
    { key: 'researchTask',      label: 'Research' },
    { key: 'recitation',        label: 'Recit.' },
    { key: 'caseStudy',         label: 'Case Study' },
    { key: 'activity',          label: 'Activity' },
    { key: 'videoPresentation', label: 'Video' },
    { key: 'oba',               label: 'OBA' },
    { key: 'withProfessors',    label: 'W/Prof' },
    { key: 'withStudents',      label: 'W/Stud' },
    { key: 'status',            label: 'Status' },
  ];

  const activeExtra = EXTRA_COLS.filter(col =>
    (window.S.extraIdx || {})[col.key] !== undefined
  );

  // On mobile: only show Status in extra columns to save space
  const visibleExtra = mob
    ? activeExtra.filter(c => c.key === 'status')
    : activeExtra;

  // Build colspan for "extra" group header
  const extraGroupSpan = visibleExtra.length;

  let h = `<div class="tbl-scroll-wrap"><table class="main-tbl"><thead>
    <tr class="thead-group">
      <th class="th-frozen th-num" rowspan="2">#</th>
      <th class="th-frozen th-prof-hdr" rowspan="2">Professor</th>
      ${showDept    ? `<th rowspan="2">Dept</th>` : ''}
      <th rowspan="2">Course</th>
      ${showSection ? `<th rowspan="2">Section</th>` : ''}
      ${showLink    ? `<th rowspan="2">Link</th>` : ''}
      <th colspan="${wcols.length}" class="th-group-lbl">Weekly Uploads</th>
      <th colspan="2" class="th-group-lbl">Summary</th>
      <th rowspan="2">%</th>
      ${extraGroupSpan > 0 ? `<th colspan="${extraGroupSpan}" class="th-group-lbl th-extra-group">Activities</th>` : ''}
    </tr>
    <tr>
      ${wcols.map(wk => `<th class="th-week">W${wk}</th>`).join('')}
      <th style="text-align:center;color:var(--green);font-size:12px;min-width:28px">✓</th>
      <th style="text-align:center;color:var(--red);font-size:12px;min-width:28px">✗</th>
      ${visibleExtra.map(col => `<th class="th-extra">${col.label}</th>`).join('')}
    </tr>
  </thead><tbody>`;

  filtered.forEach(({ r, i }, di) => {
    const dk  = r.dept.toLowerCase();
    const bc  = dk.includes('bstm')  ? 'b-bstm'
              : dk.includes('bse')   ? 'b-bse'
              : dk.includes('bsrem') ? 'b-bsrem'
              : dk.includes('bpa')   ? 'b-bpa'
              : dk.includes('non')   ? 'b-nonabm'
              : dk.includes('itvet') ? 'b-itvet'
              : dk.includes('gened') ? 'b-gened'
              : '';
    const pct     = Math.round(r.done / TOTAL_WEEKS * 100);
    const canEdit = _canEdit();
    const wcs = wcols.map(wk => {
      const v = r.weeks[wk];
      return `<td class="td-week"><span
        class="wdot ${v || 'e'}${canEdit ? '' : ' viewer-dot'}"
        onclick="toggleWeek(this,${i},${wk})"
        title="${canEdit ? 'Toggle W' + wk : '👁 View only — W' + wk + ': ' + esc(r.professor)}"
        style="${canEdit ? '' : 'cursor:default;opacity:.85;'}"
        >${v === 'c' ? '✓' : v === 'n' ? '✗' : '·'}</span></td>`;
    }).join('');

    const STATUS_OPTS = ['', 'Complete', 'No Upload', 'Overdue', 'Revision', 'Pending'];
    const extraCells = visibleExtra.map(col => {
      const val = r[col.key] || '';

      if (col.key === 'status') {
        const statusCls = val === 'Complete'  ? 'sd-complete'
                        : val === 'No Upload' ? 'sd-noupload'
                        : val === 'Overdue'   ? 'sd-overdue'
                        : val === 'Revision'  ? 'sd-revision'
                        : val === 'Pending'   ? 'sd-pending'
                        : 'sd-empty';
        if (!canEdit) {
          // Viewer — show as plain styled badge, no dropdown arrow
          return `<td class="td-extra td-status">
            <span class="status-dd ${statusCls}" style="display:inline-block;pointer-events:none;cursor:default;background-image:none;padding-right:10px;">
              ${val || '—'}
            </span>
          </td>`;
        }
        const opts = STATUS_OPTS.map(o =>
          `<option value="${o}" ${o === val ? 'selected' : ''}>${o || '—'}</option>`
        ).join('');
        return `<td class="td-extra td-status">
          <select class="status-dd ${statusCls}"
            onchange="onStatusChange(this,${i})">
            ${opts}
          </select>
        </td>`;
      }

      const displayVal = (val !== '' && val !== undefined) ? val : '';
      return `<td class="td-extra td-activity">
        <input
          type="number" min="0"
          class="activity-input ${displayVal !== '' ? 'has-val' : 'empty-val'}"
          value="${esc(String(displayVal))}"
          placeholder="—"
          data-old-val="${esc(String(displayVal))}"
          onchange="onActivityChange(this,${i},'${col.key}')"
          onkeydown="if(event.key==='Enter'){this.blur()}"
          title="${canEdit ? esc(col.label) + ': ' + esc(r.professor) : '👁 View only'}"
          ${canEdit ? '' : 'readonly style="cursor:default;pointer-events:none;"'}
        >
      </td>`;
    }).join('');

    // Course: shorten on mobile
    const courseDisplay = mob && r.course.length > 14
      ? r.course.substring(0, 13) + '…'
      : r.course;

    h += `<tr>
      <td class="td-num td-frozen">${di + 1}</td>
      <td class="td-prof td-frozen-prof">${esc(r.professor)}</td>
      ${showDept    ? `<td><span class="badge ${bc}">${esc(r.dept)}</span></td>` : ''}
      <td class="td-course" title="${esc(r.course)}">${esc(courseDisplay)}</td>
      ${showSection ? `<td class="td-section">${esc(r.section)}</td>` : ''}
     ${showLink    ? `<td class="td-link">${r.url ? `<a href="${esc(r.url)}" target="_blank" rel="noopener"><span class="link-icon">🔗</span><span class="link-label">Open<br>Class</span></a>` : '<span style="color:var(--t3);opacity:.3">—</span>'}</td>` : ''}
      ${wcs}
      <td class="td-total" style="color:var(--green)">${r.done}</td>
      <td class="td-total" style="color:${r.miss > 0 ? 'var(--red)' : 'var(--t3)'}">${r.miss}</td>
      <td>
        <div class="prog-wrap"><div class="prog-fill" style="width:${pct}%"></div></div>
        <span class="pct-txt"> ${pct}%</span>
      </td>
      ${extraCells}
    </tr>`;
  });

  h += '</tbody></table></div>';
  document.getElementById('table-container').innerHTML = h;

  PccCharts.render();
}

/* ── Re-render table on resize (debounced) to switch mobile/desktop layout ── */
let _tableResizeTimer;
window.addEventListener('resize', () => {
  clearTimeout(_tableResizeTimer);
  _tableResizeTimer = setTimeout(() => {
    if (window.S.rows.length) renderTable();
  }, 300);
});

/* ══════════════════════════════════════════════════════════════
   STATS
══════════════════════════════════════════════════════════════ */
function updateStats() {
  const rows = window.S.rows;
  const w    = window.S.week;

  document.getElementById('s-total').textContent    = rows.length || '—';
  document.getElementById('s-dept-sub').textContent = window.S.sheet || '—';

  let c = 0, n = 0;
  if (w === 0) {
    c = rows.reduce((a, r) => a + r.done, 0);
    n = rows.reduce((a, r) => a + r.miss, 0);
  } else {
    rows.forEach(r => {
      if (r.weeks[w] === 'c') c++;
      if (r.weeks[w] === 'n') n++;
    });
  }

  document.getElementById('s-complete').textContent     = c || '—';
  document.getElementById('s-noupload').textContent     = n || '—';
  const rateEl = document.getElementById('s-rate');
  if (rateEl && !window.S.rows.length) rateEl.textContent = '—';
  document.getElementById('s-complete-sub').textContent = w === 0 ? 'all weeks' : 'week ' + w;
  document.getElementById('s-noupload-sub').textContent = w === 0 ? 'all weeks' : 'week ' + w;
}

function setStatus(msg, isErr) {
  const el     = document.getElementById('status-bar');
  el.textContent = msg;
  el.className   = 'status-bar ' + (isErr ? 'err' : 'show');
}

/* ══════════════════════════════════════════════════════════════
   EXPORT CSV
══════════════════════════════════════════════════════════════ */
function exportCSV() {
  if (!window.S.rows.length) { toast('No data to export', 'err'); return; }

  const headers = [
    'Professor', 'Dept', 'Course', 'Section', 'URL',
    ...Array.from({ length: TOTAL_WEEKS }, (_, i) => 'Week' + (i + 1)),
    'Done', 'Missing',
  ];
  const body = window.S.rows.map(r => [
    r.professor, r.dept, r.course, r.section, r.url,
    ...Array.from({ length: TOTAL_WEEKS }, (_, i) =>
      r.weeks[i + 1] === 'c' ? 1 : r.weeks[i + 1] === 'n' ? 0 : ''
    ),
    r.done, r.miss,
  ]);

  const csv = [headers, ...body]
    .map(row => row.map(v => `"${String(v).replace(/"/g, '""')}"`).join(','))
    .join('\n');

  const a = Object.assign(document.createElement('a'), {
    href:     URL.createObjectURL(new Blob([csv], { type: 'text/csv' })),
    download: 'PCC_' + window.S.sheet.replace(/\s/g, '_') + '_' + new Date().toISOString().slice(0, 10) + '.csv',
  });
  a.click();
  URL.revokeObjectURL(a.href);
  toast('✓ CSV exported!', 'ok');
}

/* ══════════════════════════════════════════════════════════════
   UTILITIES
══════════════════════════════════════════════════════════════ */
function esc(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

let _toastTimer;
function toast(msg, type) {
  const el   = document.getElementById('toast');
  el.textContent = msg;
  el.className   = 'show ' + (type || '');
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => { el.className = ''; }, 4500);
}