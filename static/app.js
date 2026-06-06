/* ============================================================
   Parser Hub — app.js
   ============================================================ */

// ── State ──────────────────────────────────────────────────
const activeTasks = {};
const taskPollers = {};
let linksData        = [];
let sortCol          = 'profit';
let sortDir          = 'desc';
let currentLinksFile = null;
let scanData = { chrome_profiles:[], session_profiles:{}, sda_profiles:[], proxy_files:[] };
let externalFiles    = [];  // CSV files from external_data/

// ── Init ───────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', async () => {
  await loadScanData();
  await loadConfig();
  await loadExternalFiles();
  refreshFiles();
  setInterval(refreshFiles, 8000);
  setInterval(loadExternalFiles, 15000);
});

// ── Tab switching ───────────────────────────────────────────
const TAB_IDS = ['parsers', 'links', 'profiles', 'settings'];
function switchTab(tab) {
  TAB_IDS.forEach(id => {
    document.getElementById(`tab-${id}`).style.display = id === tab ? '' : 'none';
    const btn = document.getElementById(`tab-btn-${id}`);
    if (btn) {
      btn.style.borderColor = id === tab ? 'var(--blue)' : '';
      btn.style.color       = id === tab ? 'var(--blue)' : '';
    }
  });
  if (tab === 'profiles') refreshChromeProfiles();
}

// ── Toast ──────────────────────────────────────────────────
function showToast(msg, type = 'info', duration = 3500) {
  const icons = { success:'✅', error:'❌', info:'ℹ️' };
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  el.innerHTML = `<span>${icons[type]||''}</span> ${escHtml(msg)}`;
  document.getElementById('toast-container').appendChild(el);
  setTimeout(() => {
    el.style.animation = 'slide-out 0.25s ease forwards';
    setTimeout(() => el.remove(), 250);
  }, duration);
}

// ── Scan environment ───────────────────────────────────────
async function loadScanData() {
  try {
    const res = await fetch('/api/scan');
    scanData  = await res.json();
    applyScannedData();
  } catch(e) { console.warn('scan failed', e); }
}

function applyScannedData() {
  const { chrome_profiles, session_profiles, proxy_files, sda_profiles } = scanData;

  // Proxy selects
  const proxyOpts = '<option value="">— без прокси —</option>' +
    proxy_files.map(f => `<option value="${f}">${f}</option>`).join('');
  ['steam_market-proxies','steam_priceoverview-proxies','tradeit_site-proxies','cfg-proxies'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.innerHTML = proxyOpts;
  });

  // Chrome profile selects (parser cards)
  const chromeOpts = '<option value="">— без профиля —</option>' +
    chrome_profiles.map(p => `<option value="${p.path}">${p.name}</option>`).join('');
  ['swapgg_site-user_data_dir','swapgg_user-user_data_dir','skinswap_site-user_data_dir'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.innerHTML = chromeOpts;
  });

  // Session profiles
  _fillSessionSelect('tradeit_user',  session_profiles.tradeit  || [], 'main');
  _fillSessionSelect('skinswap_user', session_profiles.skinswap || [], 'default');

  // SDA profiles
  const sdaList = document.getElementById('sda-list');
  if (sdaList) {
    sdaList.innerHTML = sda_profiles.length
      ? sda_profiles.map(s => `<div class="file-chip">🔑 ${escHtml(s)}</div>`).join('')
      : '<span style="color:var(--text-dim);font-size:0.78rem;">Не найдено</span>';
  }
}

function _fillSessionSelect(parserId, names, defaultName) {
  const sel = document.getElementById(`${parserId}-user_profile`);
  if (!sel) return;
  sel.innerHTML = (names.length
    ? names.map(n => `<option value="${n}">${n}</option>`).join('')
    : `<option value="${defaultName}">${defaultName}</option>`)
    + '<option value="">✎ Новый...</option>';
  if (names.includes(defaultName)) sel.value = defaultName;
  else if (names.length) sel.value = names[0];
}

function toggleCustomProfile(parserId) {
  const sel    = document.getElementById(`${parserId}-user_profile`);
  const custom = document.getElementById(`${parserId}-user_profile-custom`);
  if (!sel || !custom) return;
  const show = custom.style.display === 'none';
  sel.style.display    = show ? 'none' : '';
  custom.style.display = show ? '' : 'none';
  if (show) custom.focus();
}

function getProfileValue(parserId) {
  const custom = document.getElementById(`${parserId}-user_profile-custom`);
  const sel    = document.getElementById(`${parserId}-user_profile`);
  if (custom && custom.style.display !== 'none') return custom.value.trim();
  return sel ? sel.value : '';
}

// ── Chrome Profiles Tab ─────────────────────────────────────
async function refreshChromeProfiles() {
  const container = document.getElementById('chrome-profiles-list');
  if (!container) return;
  try {
    const res      = await fetch('/api/chrome-profiles');
    const profiles = await res.json();
    if (!profiles.length) {
      container.innerHTML = '<span style="color:var(--text-dim);font-size:0.82rem;">Нет созданных профилей.</span>';
      return;
    }
    container.innerHTML = profiles.map(p => `
      <div class="chrome-profile-row">
        <div class="chrome-profile-info">
          <div class="chrome-profile-icon">🌐</div>
          <div>
            <div class="chrome-profile-name">${escHtml(p.name)}</div>
            <div class="chrome-profile-path">${escHtml(p.path)}</div>
          </div>
        </div>
        <button class="btn btn-ghost btn-sm" style="color:var(--red);border-color:var(--red);"
          onclick="deleteChromeProfile('${escHtml(p.name)}')">🗑 Удалить</button>
      </div>
    `).join('');

    // Also refresh chrome selects in parser cards
    await loadScanData();
  } catch(e) {
    container.innerHTML = `<span style="color:var(--red);">Ошибка: ${escHtml(e.message)}</span>`;
  }
}

async function createChromeProfile() {
  const nameEl = document.getElementById('new-profile-name');
  const name   = nameEl.value.trim();
  if (!name) { showToast('Введите название профиля', 'error'); return; }

  try {
    const res  = await fetch('/api/create-chrome-profile', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ name }),
    });
    const data = await res.json();
    if (data.error) { showToast(data.error, 'error'); return; }

    nameEl.value = '';
    showToast(`Профиль "${data.name}" создан`, 'success');
    refreshChromeProfiles();
  } catch(e) {
    showToast(`Ошибка: ${e.message}`, 'error');
  }
}

async function deleteChromeProfile(name) {
  if (!confirm(`Удалить профиль "${name}" и все его данные?`)) return;
  try {
    const res  = await fetch('/api/delete-chrome-profile', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ name }),
    });
    const data = await res.json();
    if (data.error) { showToast(data.error, 'error'); return; }
    showToast(`Профиль "${name}" удалён`, 'success');
    refreshChromeProfiles();
  } catch(e) {
    showToast(`Ошибка: ${e.message}`, 'error');
  }
}

// ── Config ─────────────────────────────────────────────────
async function loadConfig() {
  try {
    const cfg = await (await fetch('/api/config')).json();
    _setVal('cfg-proxies',  cfg.proxies_file);
    _setVal('cfg-names',    cfg.names_file);
    _setVal('cfg-appid',    cfg.default_appid);
    _setVal('steam_market-appid',     cfg.default_appid);
    _setVal('tradeit_site-game_id',   cfg.default_game_id);
    _setVal('tradeit_user-game_id',   cfg.default_game_id);
    if (cfg.workers) Object.entries(cfg.workers).forEach(([p,w]) => _setVal(`${p}-workers`, w));
    if (cfg.outputs) Object.entries(cfg.outputs).forEach(([p,n]) => _setVal(`${p}-out`, n));
    if (cfg.proxies_file) {
      ['steam_market-proxies','steam_priceoverview-proxies','tradeit_site-proxies'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = cfg.proxies_file;
      });
    }
  } catch(e) { console.warn('config load failed', e); }
}

async function saveConfig() {
  const cfg = {
    proxies_file:    document.getElementById('cfg-proxies')?.value || '',
    names_file:      document.getElementById('cfg-names')?.value   || 'names.txt',
    default_appid:   document.getElementById('cfg-appid')?.value   || '252490',
    default_game_id: document.getElementById('cfg-appid')?.value   || '252490',
  };
  try {
    const res = await fetch('/api/config', {
      method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(cfg),
    });
    const data = await res.json();
    if (data.ok) showToast('Настройки сохранены', 'success');
    else showToast('Ошибка сохранения', 'error');
  } catch(e) { showToast(`Ошибка: ${e.message}`, 'error'); }
}

function _setVal(id, val) {
  const el = document.getElementById(id);
  if (el && val != null) el.value = val;
}

// ── Log toggle ──────────────────────────────────────────────
function toggleLog(parserId) {
  document.getElementById(`log-${parserId}`).classList.toggle('visible');
}

// ── External CSV Files ──────────────────────────────────────
async function loadExternalFiles() {
  try {
    const res = await fetch('/api/list-external-files');
    externalFiles = await res.json();
    _refreshLinksSelects();
    _renderExternalFilesList();
  } catch(e) { console.warn('external files load failed', e); }
}

function _renderExternalFilesList() {
  const el = document.getElementById('external-files-list');
  if (!el) return;
  if (!externalFiles || !externalFiles.length) {
    el.innerHTML = '<span style="color:var(--text-dim);font-size:0.78rem;">Нет CSV файлов в external_data/</span>';
    return;
  }
  el.innerHTML = externalFiles.map(f => `
    <div class="file-chip" style="border-color:rgba(59,224,138,0.3);">
      📂 <span style="color:var(--green);">${escHtml(f.name)}</span>
      <span class="file-size">${formatBytes(f.size)}</span>
    </div>
  `).join('');
}

// ── Files ───────────────────────────────────────────────────
async function refreshFiles() {
  try {
    const files = await (await fetch('/api/list-files')).json();
    renderFilesList(files);
    updateFileSelects(files);
  } catch(e) {}
}

function formatBytes(b) {
  if (b < 1024)         return `${b} B`;
  if (b < 1024*1024)   return `${(b/1024).toFixed(1)} KB`;
  return `${(b/1024/1024).toFixed(1)} MB`;
}

function renderFilesList(files) {
  const el = document.getElementById('files-list');
  if (!el) return;
  if (!files.length) {
    el.innerHTML = '<span style="color:var(--text-dim);font-size:0.78rem;">Нет файлов. Запустите парсер.</span>';
    return;
  }
  el.innerHTML = files.map(f => `
    <div class="file-chip">
      📊 <a href="/api/download/${encodeURIComponent(f.name)}" download>${escHtml(f.name)}</a>
      <span class="file-size">${formatBytes(f.size)}</span>
    </div>
  `).join('');
}

function updateFileSelects(files) {
  _refreshLinksSelects(files);
}

function _refreshLinksSelects(xlsxFiles) {
  // xlsxFiles may be undefined on first call from loadExternalFiles()
  // In that case we just rebuild with whatever externalFiles has.
  // We always use the latest DOM values for xlsx list.
  ['source-file','dest-file'].forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    const prev = el.value;

    let html = '<option value="">— выберите файл —</option>';

    // Group 1: Excel files from data/
    const xFiles = xlsxFiles ?? _getCurrentXlsxOptions(el);
    if (xFiles && xFiles.length) {
      html += '<optgroup label="📊 Парсеры (xlsx)">' +
        xFiles.map(f => `<option value="${f.name}">${escHtml(f.name)}</option>`).join('') +
        '</optgroup>';
    }

    // Group 2: External CSV files
    if (externalFiles && externalFiles.length) {
      html += '<optgroup label="📂 Внешние данные (csv)">' +
        externalFiles.map(f => `<option value="${f.name}" style="color:var(--accent)">[CSV] ${escHtml(f.name)}</option>`).join('') +
        '</optgroup>';
    }

    el.innerHTML = html;
    if (prev && el.querySelector(`option[value="${CSS.escape(prev)}"]`)) el.value = prev;
  });
}

/** Read current xlsx options from a select (preserves them during external-only refresh). */
function _getCurrentXlsxOptions(el) {
  const result = [];
  el.querySelectorAll('option[value]').forEach(opt => {
    const v = opt.value;
    if (v && v.endsWith('.xlsx')) result.push({ name: v });
  });
  return result;
}

// ── Run parser ──────────────────────────────────────────────
async function runParser(parserId) {
  const args = {};

  // simple inputs
  ['workers','appid','game_id','max_items','names','sort','timeout'].forEach(field => {
    const el = document.getElementById(`${parserId}-${field}`);
    if (el && el.value.trim()) args[field] = el.value.trim();
  });

  // proxy select
  const proxyEl = document.getElementById(`${parserId}-proxies`);
  if (proxyEl && proxyEl.value) args['proxies'] = proxyEl.value;

  // chrome profile select
  const chromeEl = document.getElementById(`${parserId}-user_data_dir`);
  if (chromeEl && chromeEl.value) args['user_data_dir'] = chromeEl.value;

  // session profile
  const profileVal = getProfileValue(parserId);
  if (profileVal) args['user_profile'] = profileVal;

  // checkboxes
  ['verbose','accepted_only','append','headless'].forEach(field => {
    const el = document.getElementById(`${parserId}-${field}`);
    if (el) args[field] = el.checked;
  });

  // max_items (include 0)
  const maxEl = document.getElementById(`${parserId}-max_items`);
  if (maxEl) args['max_items'] = maxEl.value.trim() || '0';

  const outEl = document.getElementById(`${parserId}-out`);
  const out   = outEl?.value.trim() || `${parserId}.xlsx`;

  setStatus(parserId, 'running');
  clearLog(parserId);
  appendLog(parserId, `▶ Запуск парсера [${parserId}]...`);
  document.getElementById(`log-${parserId}`).classList.add('visible');

  try {
    const res  = await fetch('/api/run-parser', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ parser_id: parserId, out, args }),
    });
    const data = await res.json();
    if (data.error) {
      setStatus(parserId, 'error');
      appendLog(parserId, `❌ ${data.error}`);
      showToast(`Ошибка: ${data.error}`, 'error');
      return;
    }
    activeTasks[parserId] = data.task_id;
    showToast(`Парсер запущен (${parserId})`, 'info');
    startPolling(parserId, data.task_id);
  } catch(e) {
    setStatus(parserId, 'error');
    appendLog(parserId, `❌ Ошибка: ${e.message}`);
    showToast(`Ошибка: ${e.message}`, 'error');
  }
}

function startPolling(parserId, taskId) {
  if (taskPollers[parserId]) clearInterval(taskPollers[parserId]);
  let lastLen = 0;
  taskPollers[parserId] = setInterval(async () => {
    try {
      const task = await (await fetch(`/api/task-status/${taskId}`)).json();
      if (task.error) return;
      const log = task.log || [];
      if (log.length > lastLen) {
        log.slice(lastLen).forEach(l => appendLog(parserId, l));
        lastLen = log.length;
      }
      if (task.status === 'done') {
        setStatus(parserId, 'done');
        clearInterval(taskPollers[parserId]);
        delete activeTasks[parserId];
        appendLog(parserId, '✅ Готово!');
        showToast(`Парсер завершён (${parserId})`, 'success');
        refreshFiles();
        loadScanData();
      } else if (task.status === 'error') {
        setStatus(parserId, 'error');
        clearInterval(taskPollers[parserId]);
        delete activeTasks[parserId];
        showToast(`Ошибка парсера (${parserId})`, 'error');
        refreshFiles();
      }
    } catch(e) {}
  }, 1000);
}

// ── Cancel parser ────────────────────────────────────────────
async function cancelParser(parserId) {
  const taskId = activeTasks[parserId];
  if (!taskId) { showToast('Нет активной задачи для отмены', 'info'); return; }
  try {
    await fetch(`/api/cancel-task/${taskId}`, { method: 'POST' });
    if (taskPollers[parserId]) { clearInterval(taskPollers[parserId]); delete taskPollers[parserId]; }
    delete activeTasks[parserId];
    setStatus(parserId, 'cancelled');
    appendLog(parserId, '⏹ Парсер остановлен пользователем');
    showToast(`Парсер остановлен (${parserId})`, 'info');
  } catch(e) {
    showToast(`Ошибка отмены: ${e.message}`, 'error');
  }
}

const STATUS_LABELS = { idle:'Ожидание', running:'Запущен', done:'Готово', error:'Ошибка', cancelled:'Отменено' };
function setStatus(parserId, status) {
  const el = document.getElementById(`status-${parserId}`);
  if (el) {
    el.className  = `status-pill status-${status}`;
    el.textContent = STATUS_LABELS[status] || status;
  }
  // Show cancel button only while running
  const cancelBtn = document.getElementById(`cancel-btn-${parserId}`);
  if (cancelBtn) cancelBtn.style.display = status === 'running' ? '' : 'none';
}
function clearLog(parserId) {
  const el = document.getElementById(`log-${parserId}`);
  if (el) el.innerHTML = '';
}
function appendLog(parserId, line) {
  const el = document.getElementById(`log-${parserId}`);
  if (!el) return;
  const div = document.createElement('div');
  div.textContent = line;
  const lo = line.toLowerCase();
  if (lo.includes('error') || lo.includes('ошибка') || line.includes('❌'))
    div.style.color = 'var(--red)';
  else if (line.includes('✅') || lo.includes('готово') || lo.includes('saved') || lo.includes('сохранено'))
    div.style.color = 'var(--green)';
  el.appendChild(div);
  el.scrollTop = el.scrollHeight;
  while (el.children.length > 300) el.removeChild(el.firstChild);
}

// ── Build Links ─────────────────────────────────────────────
async function buildLinks() {
  const sourceFile  = document.getElementById('source-file').value;
  const destFile    = document.getElementById('dest-file').value;
  const sourceCoeff = parseFloat(document.getElementById('source-coeff').value) || 1.0;
  const destCoeff   = parseFloat(document.getElementById('dest-coeff').value)   || 1.0;
  const minProfit   = parseFloat(document.getElementById('min-profit').value)   || 0;
  const minRoi      = parseFloat(document.getElementById('min-roi').value)      || 0;
  const outFile     = document.getElementById('links-out').value.trim()         || 'links_result.xlsx';

  if (!sourceFile) { showToast('Выберите файл источника (площадка отдаёт)', 'error'); return; }
  if (!destFile)   { showToast('Выберите файл назначения (площадка принимает)', 'error'); return; }
  if (sourceFile === destFile) { showToast('Файлы совпадают', 'error'); return; }

  const btn = document.getElementById('build-btn');
  btn.disabled  = true;
  btn.innerHTML = '<span class="spinner"></span> Обработка...';
  document.getElementById('results-panel').style.display = 'none';
  document.getElementById('links-empty').style.display   = 'none';

  try {
    const res  = await fetch('/api/build-links', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ source_file:sourceFile, dest_file:destFile,
                             source_coeff:sourceCoeff, dest_coeff:destCoeff,
                             min_profit:minProfit, min_roi:minRoi, out_file:outFile }),
    });
    const data = await res.json();
    if (data.error) { showToast(`Ошибка: ${data.error}`, 'error'); return; }

    linksData        = data.items || [];
    currentLinksFile = data.out_file;

    if (!linksData.length) {
      document.getElementById('links-empty').style.display = 'block';
      showToast(`Нет совпадений (${data.source_total} источников × ${data.dest_total} назначений)`, 'info');
    } else {
      renderLinksTable(linksData);
      document.getElementById('results-panel').style.display = 'block';
      document.getElementById('download-links-btn').style.display = '';

      const top = linksData[0];
      document.getElementById('results-meta').innerHTML = `
        <span class="meta-chip">📦 Источник: <strong>${data.source_total}</strong> предм.</span>
        <span class="meta-chip">👤 Назначение: <strong>${data.dest_total}</strong> предм.</span>
        <span class="meta-chip">Совпало: <strong>${data.matched}</strong></span>
        <span class="meta-chip" style="border-color:var(--green);color:var(--green);">
          Топ прибыль: <strong>+${top?.profit?.toFixed(2) ?? '—'}</strong>
          (ROI: ${top?.roi?.toFixed(1) ?? '—'}%)
        </span>
      `;
      showToast(`Найдено ${data.matched} связок!`, 'success');
    }
    refreshFiles();
  } catch(e) {
    showToast(`Ошибка: ${e.message}`, 'error');
  } finally {
    btn.disabled  = false;
    btn.innerHTML = '⚡ Построить связки';
  }
}

function renderLinksTable(data) {
  const tbody = document.getElementById('links-tbody');
  tbody.innerHTML = '';
  [...data].sort((a,b) => {
    let av = a[sortCol], bv = b[sortCol];
    if (typeof av === 'string') av = av.toLowerCase();
    if (typeof bv === 'string') bv = bv.toLowerCase();
    if (av < bv) return sortDir === 'asc' ? -1 :  1;
    if (av > bv) return sortDir === 'asc' ?  1 : -1;
    return 0;
  }).forEach(item => {
    const pCls = item.profit > 0 ? 'profit-pos' : item.profit < 0 ? 'profit-neg' : 'profit-zero';
    const rCls = item.roi > 20 ? 'roi-high' : item.roi > 5 ? 'roi-med' : 'roi-low';
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="td-name" title="${escHtml(item.name)}">${escHtml(item.name)}</td>
      <td class="td-price" style="color:var(--blue)">${item.source_price.toFixed(2)}</td>
      <td class="td-price" style="color:var(--green)">${item.dest_price.toFixed(2)}</td>
      <td class="td-price" style="color:var(--text-muted)">${item.source_eff.toFixed(2)}</td>
      <td class="td-price" style="color:var(--text-muted)">${item.dest_eff.toFixed(2)}</td>
      <td class="td-profit ${pCls}">${item.profit > 0 ? '+' : ''}${item.profit.toFixed(2)}</td>
      <td class="td-roi ${rCls}">${item.roi > 0 ? '+' : ''}${item.roi.toFixed(1)}%</td>
    `;
    tbody.appendChild(tr);
  });
}

function sortTable(col) {
  sortDir = (sortCol === col && sortDir === 'desc') ? 'asc' : 'desc';
  sortCol = col;
  if (linksData.length) renderLinksTable(linksData);
}

function downloadLinks() {
  if (!currentLinksFile) return;
  const a = document.createElement('a');
  a.href = `/api/download/${encodeURIComponent(currentLinksFile)}`;
  a.download = currentLinksFile;
  a.click();
}

function escHtml(str) {
  return String(str)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
