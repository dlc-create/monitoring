/* ════════════════════════════════════════════════════════════
   EMAIL ANNOUNCE PANEL
   Paste this entire block at the BOTTOM of pcc-app.js
   (after the last closing brace of the toast() function)
════════════════════════════════════════════════════════════ */

/* ── State ── */
let _announceScope = 'missing'; // 'missing' | 'all' | 'dept' | 'alldepts'

/* ══════════════════════════════════════════════════════════
   SHOW / HIDE ANNOUNCE BUTTON (admin only, after login)
══════════════════════════════════════════════════════════ */
function _showAnnounceBtn() {
  const btn = document.getElementById('btn-announce');
  if (btn) btn.style.display = (window.S.role === 'admin') ? '' : 'none';
}

/* Hook into the existing loadSheet so button appears after data loads */
const _origLoadSheet_ann = window.loadSheet;
window.loadSheet = async function(sheetName) {
  await _origLoadSheet_ann.call(this, sheetName);
  _showAnnounceBtn();
};

/* ══════════════════════════════════════════════════════════
   OPEN MODAL
══════════════════════════════════════════════════════════ */
function openAnnounceModal() {
  if (window.S.role !== 'admin') { toast('⛔ Admin only', 'err'); return; }
  if (!window.S.rows.length)     { toast('Load a department first', 'err'); return; }

  // Reset to clean state
  _announceScope = 'missing';
  document.querySelectorAll('.scope-pill').forEach(p =>
    p.classList.toggle('active', p.getAttribute('data-scope') === 'missing')
  );
  document.getElementById('ann-progress').classList.remove('show');
  document.getElementById('ann-result').classList.remove('show');
  document.getElementById('ann-btn-send').disabled = false;
  document.getElementById('ann-btn-label').textContent = 'Send Emails';
  document.getElementById('ann-btn-icon').textContent  = '📧';
  document.getElementById('ann-btn-cancel').textContent = 'Cancel';

  // Update preview stats
  _updateAnnouncePreview();

  // Open
  document.getElementById('announce-overlay').classList.add('open');
  document.body.style.overflow = 'hidden';
}

/* ══════════════════════════════════════════════════════════
   CLOSE MODAL
══════════════════════════════════════════════════════════ */
function closeAnnounceModal(e) {
  // If called from overlay click, only close if clicking the overlay itself
  if (e && e.target !== document.getElementById('announce-overlay')) return;
  document.getElementById('announce-overlay').classList.remove('open');
  document.body.style.overflow = '';
}

/* ══════════════════════════════════════════════════════════
   SCOPE SELECTOR
══════════════════════════════════════════════════════════ */
function setAnnounceScope(scope, el) {
  _announceScope = scope;
  document.querySelectorAll('.scope-pill').forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  _updateAnnouncePreview();
}

/* ══════════════════════════════════════════════════════════
   LIVE PREVIEW STATS
══════════════════════════════════════════════════════════ */
function _updateAnnouncePreview() {
  const rows  = window.S.rows;
  const w     = window.S.week || _getActiveWeek();
  const total = rows.length;
  const miss  = rows.filter(r => r.weeks[w] === 'n' || r.weeks[w] === '').length;

  let recipients;
  switch (_announceScope) {
    case 'missing':  recipients = miss;  break;
    case 'all':      recipients = total; break;
    case 'dept':     recipients = total; break;
    case 'alldepts': recipients = '?';   break; // unknown until all sheets load
    default:         recipients = miss;
  }

  document.getElementById('ann-stat-total').textContent   = recipients === '?' ? '?' : recipients;
  document.getElementById('ann-stat-missing').textContent = miss;
  document.getElementById('ann-stat-week').textContent    = 'W' + w;

  // Footer scope label
  const scopeLabels = {
    missing:  '❌ Missing professors — ' + window.S.sheet,
    all:      '👥 All professors — '     + window.S.sheet,
    dept:     '🏢 Current dept — '       + window.S.sheet,
    alldepts: '🏫 All departments (7 sheets)',
  };
  document.getElementById('ann-footer-scope').textContent = scopeLabels[_announceScope] || '—';
  document.getElementById('ann-subtitle').textContent =
    _announceScope === 'missing'
      ? `${miss} professors with missing Week ${w} uploads`
      : _announceScope === 'alldepts'
      ? 'All departments — all missing professors'
      : `${total} professors in ${window.S.sheet}`;
}

/* helper: detect active week from loaded data */
function _getActiveWeek() {
  const rows = window.S.rows;
  let latest = 1;
  rows.forEach(r => {
    for (let w = TOTAL_WEEKS; w >= 1; w--) {
      if (r.weeks[w] === 'c' || r.weeks[w] === 'n') { if (w > latest) latest = w; break; }
    }
  });
  return latest;
}

/* ══════════════════════════════════════════════════════════
   SEND
══════════════════════════════════════════════════════════ */
async function sendAnnounce() {
  if (!WEBAPP_URL || !WEBAPP_URL.trim()) {
    toast('⚠️ Set WEBAPP_URL in pcc-app.js first!', 'err');
    return;
  }

  const w         = window.S.week || _getActiveWeek();
  const rows      = window.S.rows;
  const missing   = rows.filter(r => r.weeks[w] === 'n' || r.weeks[w] === '').length;
  const subject   = document.getElementById('ann-subject').value.trim();
  const extraMsg  = document.getElementById('ann-body').value.trim();
  const ccHeads   = document.getElementById('ann-cc-heads').checked;
  const sendSum   = document.getElementById('ann-send-summary').checked;
  const isAllDepts = _announceScope === 'alldepts';
  const sheetParam = isAllDepts ? '' : encodeURIComponent(window.S.sheet);

  // Recipients count for confirm dialog
  const recipCount = _announceScope === 'missing'  ? missing
                   : _announceScope === 'alldepts' ? 'all departments'
                   : rows.length;

  const confirmed = confirm(
    `📧 Send email announcements?\n\n` +
    `• Scope: ${document.getElementById('ann-footer-scope').textContent}\n` +
    `• Week: ${w}\n` +
    `• Recipients: ${recipCount} professor${recipCount === 1 ? '' : 's'}\n` +
    (ccHeads  ? '• Dept heads will be CC\'d\n' : '') +
    (sendSum  ? '• Admin summary digest will be sent\n' : '') +
    `\nThis sends REAL emails. Proceed?`
  );
  if (!confirmed) return;

  // ── UI: sending state ──
  const btnSend   = document.getElementById('ann-btn-send');
  const btnCancel = document.getElementById('ann-btn-cancel');
  const progress  = document.getElementById('ann-progress');
  const result    = document.getElementById('ann-result');

  btnSend.disabled = true;
  btnSend.querySelector ? null : null;
  document.getElementById('ann-btn-label').textContent = 'Sending…';
  document.getElementById('ann-btn-icon').textContent  = '⏳';
  btnCancel.textContent = 'Close';
  result.classList.remove('show');
  progress.classList.add('show');

  // Animate progress bar
  const bar = document.getElementById('ann-prog-bar');
  const statusEl = document.getElementById('ann-prog-status');
  bar.style.width = '15%';
  statusEl.textContent = '⏳ Connecting to mail server…';
  _setSyncStatus('saving', 'Sending emails…');

  const steps = [
    { w: '35%', msg: '📋 Reading professor data…' },
    { w: '55%', msg: '✉️ Composing emails…' },
    { w: '75%', msg: '📤 Sending reminders…' },
    { w: '90%', msg: '📊 Generating summary…' },
  ];
  let si = 0;
  const stepTimer = setInterval(() => {
    if (si < steps.length) {
      bar.style.width = steps[si].w;
      statusEl.textContent = steps[si].msg;
      si++;
    }
  }, 1800);

  try {
    const url = WEBAPP_URL
      + '?action=sendReminders'
      + (sheetParam ? '&sheet=' + sheetParam : '')
      + '&scope=' + _announceScope
      + (extraMsg  ? '&extraMsg=' + encodeURIComponent(extraMsg) : '')
      + (subject   ? '&subject='  + encodeURIComponent(subject)  : '')
      + '&ccHeads='   + (ccHeads  ? '1' : '0')
      + '&sendSummary=' + (sendSum ? '1' : '0');

    const res  = await fetch(url);
    const data = await res.json();

    clearInterval(stepTimer);

    if (!data.ok) throw new Error(data.error || 'Server error');

    // ── Success ──
    bar.style.width = '100%';
    statusEl.textContent = '✅ Complete!';

    const sent   = data.emailsSent   || 0;
    const failed = data.emailsFailed || 0;
    const wk     = data.weekNum      || w;

    // Show result panel
    document.getElementById('ann-result-title').textContent =
      sent > 0 ? `✅ ${sent} email${sent !== 1 ? 's' : ''} sent successfully!` : '⚠️ No emails sent';
    document.getElementById('ann-result-sub').textContent =
      `Week ${wk} · ${window.S.sheet}${isAllDepts ? ' + all departments' : ''}` +
      (failed > 0 ? ` · ⚠️ ${failed} delivery failure${failed !== 1 ? 's' : ''}` : '');
    document.getElementById('rg-sent').textContent   = sent;
    document.getElementById('rg-failed').textContent = failed;
    document.getElementById('rg-week').textContent   = 'W' + wk;

    // Show failed addresses if any
    const failList = document.getElementById('rg-failed-list');
    if (failed > 0 && data.failedAddresses && data.failedAddresses.length) {
      failList.style.display = 'block';
      failList.innerHTML = `
        <div style="background:var(--red-bg);border:1.5px solid rgba(240,82,82,.3);border-radius:10px;padding:12px 14px;">
          <div style="font-size:12px;font-weight:800;color:var(--red);margin-bottom:6px;">⚠️ Failed Deliveries</div>
          ${data.failedAddresses.map(f =>
            `<div style="font-size:11px;color:var(--t2);font-family:'JetBrains Mono',monospace;margin-bottom:3px;">${esc(f.email)}</div>`
          ).join('')}
          <div style="font-size:11px;color:var(--t3);margin-top:6px;">Verify name formats for these professors.</div>
        </div>`;
    } else {
      failList.style.display = 'none';
    }

    result.classList.add('show');
    document.getElementById('ann-btn-label').textContent = 'Sent ✓';
    document.getElementById('ann-btn-icon').textContent  = '✅';

    _setSyncStatus('saved', `${sent} emails sent ✓`);
    setTimeout(() => _setSyncStatus('', ''), 5000);
    toast(`✅ ${sent} reminder${sent !== 1 ? 's' : ''} sent — Week ${wk}`, 'ok');

  } catch (err) {
    clearInterval(stepTimer);
    bar.style.width = '100%';
    bar.style.background = 'var(--red)';
    statusEl.textContent = '❌ Failed: ' + err.message;

    document.getElementById('ann-btn-send').disabled = false;
    document.getElementById('ann-btn-label').textContent = 'Retry';
    document.getElementById('ann-btn-icon').textContent  = '🔄';

    _setSyncStatus('error', 'Email failed');
    setTimeout(() => _setSyncStatus('', ''), 4000);
    toast('❌ Email error: ' + err.message, 'err');
  }
}
