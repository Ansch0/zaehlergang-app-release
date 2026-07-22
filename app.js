/* global XLSX */
(() => {
  'use strict';

  const STORAGE_KEY = 'zaehlergang.session.v002';
  const TEN_OCLOCK_COLUMNS = new Set(['BG', 'BH', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR']);
  const app = document.getElementById('app');

  const state = {
    session: loadSession(),
    screen: 'home',
    selectedLocationId: null,
    selectedMeterId: null,
    filter: 'all',
    input: '',
    message: '',
    messageType: 'info'
  };
  if (state.session) state.screen = 'locations';

  function normalizeNumber(value) {
    if (value === null || value === undefined || value === '') return null;
    if (typeof value === 'number') return Number.isFinite(value) ? value : null;
    const text = String(value).trim().replace(/\s/g, '');
    if (!text) return null;
    const normalized = text.includes(',') ? text.replace(/\./g, '').replace(',', '.') : text;
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : null;
  }

  function formatNumber(value) {
    if (value === null || value === undefined || value === '') return '—';
    return new Intl.NumberFormat('de-DE', { maximumFractionDigits: 3 }).format(value);
  }

  function dateString(date = new Date()) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  function displayDate(iso) {
    const [y, m, d] = String(iso).split('-');
    return y && m && d ? `${d}.${m}.${y}` : iso;
  }

  function toBase64(buffer) {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    for (let offset = 0; offset < bytes.length; offset += 32768) {
      binary += String.fromCharCode(...bytes.subarray(offset, offset + 32768));
    }
    return btoa(binary);
  }

  function fromBase64(base64) {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
    return bytes.buffer;
  }

  function loadSession() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch { return null; }
  }

  function persist() {
    if (!state.session) return;
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state.session));
    } catch {
      showMessage('Der lokale Browser-Speicher ist voll. Bitte sofort „Excel sichern“ verwenden.', 'error');
    }
  }

  function showMessage(text, type = 'info') {
    state.message = text;
    state.messageType = type;
    render();
  }

  function clearMessage() {
    state.message = '';
  }

  function parseWorkbook(file, buffer) {
    if (!window.XLSX) throw new Error('Die Excel-Bibliothek konnte nicht geladen werden. Bitte die App einmal mit Internetverbindung öffnen.');
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
    const dataSheet = workbook.Sheets.Daten_Ausdruck;
    const sortedSheet = workbook.Sheets.Zaehlerliste_Sortiert;
    if (!dataSheet || !sortedSheet) throw new Error('Benötigte Blätter „Daten_Ausdruck“ oder „Zaehlerliste_Sortiert“ fehlen.');

    const metersByNumber = new Map();
    for (let row = 5; row <= 524; row += 1) {
      const raw = dataSheet[`A${row}`]?.v;
      if (raw === undefined || raw === null || raw === '') continue;
      const number = String(raw).trim();
      metersByNumber.set(number, {
        id: `${number}-${row}`,
        number,
        timestamp: dataSheet[`B${row}`]?.v ? new Date(dataSheet[`B${row}`].v).toISOString() : null,
        currentValue: normalizeNumber(dataSheet[`C${row}`]?.v),
        previousValue: normalizeNumber(dataSheet[`D${row}`]?.v),
        info: '',
        locationId: null,
        location: 'Unzugeordnet',
        sourceRow: row,
        changed: false
      });
    }

    const locations = [];
    const range = XLSX.utils.decode_range(sortedSheet['!ref'] || 'A1:A1');
    for (let col = 1; col <= range.e.c; col += 1) {
      const column = XLSX.utils.encode_col(col);
      const title = sortedSheet[`${column}1`]?.v;
      if (!title) continue;
      const numbers = [];
      for (let row = 2; row <= 41; row += 1) {
        const value = sortedSheet[`${column}${row}`]?.v;
        if (value !== undefined && value !== null && value !== '') numbers.push(String(value).trim());
      }
      if (!numbers.length) continue;
      const name = String(title).trim();
      const location = { id: column, name, group: TEN_OCLOCK_COLUMNS.has(column) ? '10 Uhr' : 'Normal', meterNumbers: numbers };
      locations.push(location);
      numbers.forEach((number, index) => {
        const meter = metersByNumber.get(number);
        if (!meter) return;
        meter.locationId = column;
        meter.location = name;
        meter.info = String(sortedSheet[`${column}${45 + index}`]?.v ?? '').trim();
      });
    }

    const workbookDate = dataSheet.B3?.v;
    const date = workbookDate instanceof Date && !Number.isNaN(workbookDate.valueOf()) ? dateString(workbookDate) : dateString();
    const meters = [...metersByNumber.values()].filter(m => m.locationId);
    if (!locations.length || !meters.length) throw new Error('Es konnten keine Standorte oder Zähler erkannt werden.');

    return { fileName: file.name, date, meters, locations, workbookBase64: toBase64(buffer), importedAt: new Date().toISOString() };
  }

  function progress() {
    const meters = state.session?.meters || [];
    if (!meters.length) return 0;
    return Math.round(meters.filter(m => m.currentValue !== null).length / meters.length * 100);
  }

  function locationStats(locationId) {
    const meters = state.session.meters.filter(m => m.locationId === locationId);
    const done = meters.filter(m => m.currentValue !== null).length;
    return { total: meters.length, done, status: done === 0 ? 'open' : done === meters.length ? 'done' : 'partial' };
  }

  function selectedLocation() { return state.session?.locations.find(l => l.id === state.selectedLocationId) || null; }
  function selectedMeter() { return state.session?.meters.find(m => m.id === state.selectedMeterId) || null; }

  async function importWorkbook(file) {
    clearMessage();
    try {
      const buffer = await file.arrayBuffer();
      state.session = parseWorkbook(file, buffer);
      state.screen = 'locations';
      state.selectedLocationId = null;
      state.selectedMeterId = null;
      persist();
      render();
    } catch (error) {
      showMessage(error instanceof Error ? error.message : 'Die Datei konnte nicht gelesen werden.', 'error');
    }
  }

  function createNewSession() {
    if (!state.session) return;
    const now = dateString();
    state.session = {
      ...state.session,
      date: now,
      fileName: `Zählergang ${now}.xlsx`,
      meters: state.session.meters.map(m => ({
        ...m,
        previousValue: m.currentValue !== null ? m.currentValue : m.previousValue,
        currentValue: null,
        timestamp: null,
        changed: false
      }))
    };
    state.screen = 'locations';
    state.selectedLocationId = null;
    state.selectedMeterId = null;
    persist();
    showMessage('Neuer Zählergang wurde angelegt. Die Werte werden automatisch lokal gesichert.', 'success');
  }

  function saveMeterValue() {
    const meter = selectedMeter();
    if (!meter) return;
    const value = normalizeNumber(state.input);
    if (value === null) return showMessage('Bitte einen gültigen Zahlenwert eingeben.', 'error');

    if (meter.currentValue !== null && !confirm('Für diesen Zähler wurde bereits ein Wert eingetragen. Bestehenden Wert ersetzen?')) return;
    if (meter.previousValue !== null && value < meter.previousValue && !confirm('Dieser Wert ist kleiner als der Vormonatswert. Trotzdem übernehmen?')) return;
    if (meter.previousValue !== null && value > meter.previousValue * 1.2 && !confirm('Der Wert ist über 20 % größer als der Vormonatswert. Trotzdem übernehmen?')) return;

    const hadValue = meter.currentValue !== null;
    state.session.meters = state.session.meters.map(m => m.id === meter.id ? {
      ...m,
      currentValue: value,
      timestamp: new Date().toISOString(),
      changed: hadValue || m.changed
    } : m);
    state.input = '';
    persist();
    clearMessage();
    render();
  }

  function exportWorkbook() {
    if (!state.session) return;
    try {
      const workbook = XLSX.read(fromBase64(state.session.workbookBase64), { type: 'array', cellDates: true });
      const sheet = workbook.Sheets.Daten_Ausdruck;
      sheet.B3 = { t: 'd', v: new Date(`${state.session.date}T12:00:00`) };
      state.session.meters.forEach(m => {
        const row = m.sourceRow;
        if (m.timestamp) sheet[`B${row}`] = { t: 'd', v: new Date(m.timestamp) }; else delete sheet[`B${row}`];
        if (m.currentValue !== null) sheet[`C${row}`] = { t: 'n', v: m.currentValue }; else delete sheet[`C${row}`];
        if (m.previousValue !== null) sheet[`D${row}`] = { t: 'n', v: m.previousValue }; else delete sheet[`D${row}`];
      });
      const safeName = state.session.fileName.replace(/\.xlsm$/i, '.xlsx');
      XLSX.writeFile(workbook, safeName, { bookType: 'xlsx', compression: true, cellDates: true });
      showMessage(`Excel-Datei „${safeName}“ wurde erstellt.`, 'success');
    } catch (error) {
      showMessage(`Excel-Datei konnte nicht erstellt werden: ${error.message || error}`, 'error');
    }
  }

  function clearLocalSession() {
    if (!confirm('Den lokal gespeicherten Zählergang wirklich entfernen?')) return;
    localStorage.removeItem(STORAGE_KEY);
    state.session = null;
    state.screen = 'home';
    state.selectedLocationId = null;
    state.selectedMeterId = null;
    clearMessage();
    render();
  }

  function messageHtml() {
    return state.message ? `<div class="message ${state.messageType}" role="alert">${escapeHtml(state.message)}</div>` : '';
  }

  function escapeHtml(value) {
    return String(value).replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
  }

  function renderHome() {
    app.innerHTML = `<main class="page home">
      <div class="brandMark">ZG</div>
      <h1>Zählergang</h1>
      <p>Offlinefähige Zählererfassung mit Excel-Archivdateien.</p>
      ${messageHtml()}
      <label class="button primary buttonLike">Excel-Datei öffnen<input id="fileInput" hidden type="file" accept=".xlsx,.xlsm"></label>
      ${state.session ? `<button class="button secondary" data-action="continue">Letzten Stand fortsetzen</button>
        <button class="button ghost" data-action="new">Neuer Zählergang</button>
        <button class="button dangerText" data-action="clear">Lokalen Stand entfernen</button>` : ''}
      <small>Alle Daten bleiben auf diesem Gerät. Es wird nichts hochgeladen.</small>
    </main>`;
    document.getElementById('fileInput').addEventListener('change', e => { const f = e.target.files?.[0]; if (f) importWorkbook(f); });
    app.querySelector('[data-action="continue"]')?.addEventListener('click', () => { state.screen = 'locations'; clearMessage(); render(); });
    app.querySelector('[data-action="new"]')?.addEventListener('click', createNewSession);
    app.querySelector('[data-action="clear"]')?.addEventListener('click', clearLocalSession);
  }

  function renderLocations() {
    const pct = progress();
    const sections = ['Normal', '10 Uhr'].map(group => {
      const locations = state.session.locations.filter(l => l.group === group);
      if (!locations.length) return '';
      return `<section><h2>${group === 'Normal' ? 'Standorte' : '10-Uhr-Ablesungen'}</h2><div class="grid">${locations.map(l => {
        const s = locationStats(l.id);
        return `<button class="location ${s.status}" data-location="${escapeHtml(l.id)}"><strong>${escapeHtml(l.name)}</strong><span>${s.done} / ${s.total} Zähler</span></button>`;
      }).join('')}</div></section>`;
    }).join('');
    app.innerHTML = `<main class="page">
      <header class="topbar"><button class="button ghost compact" data-action="home">Start</button><div class="fileMeta"><strong>${escapeHtml(state.session.fileName)}</strong><span>${displayDate(state.session.date)} · ${pct} % abgeschlossen</span></div><button class="button secondary compact" data-action="export">Excel sichern</button></header>
      ${messageHtml()}<div class="progress"><span style="width:${pct}%"></span></div>${sections}
    </main>`;
    app.querySelector('[data-action="home"]').addEventListener('click', () => { state.screen = 'home'; clearMessage(); render(); });
    app.querySelector('[data-action="export"]').addEventListener('click', exportWorkbook);
    app.querySelectorAll('[data-location]').forEach(btn => btn.addEventListener('click', () => {
      state.selectedLocationId = btn.dataset.location; state.filter = 'all'; state.screen = 'meters'; clearMessage(); render();
    }));
  }

  function renderMeters() {
    const location = selectedLocation();
    let meters = state.session.meters.filter(m => m.locationId === location.id);
    if (state.filter === 'open') meters = meters.filter(m => m.currentValue === null);
    if (state.filter === 'done') meters = meters.filter(m => m.currentValue !== null);
    app.innerHTML = `<main class="page"><header class="topbar"><button class="button ghost compact" data-action="back">Zurück</button><h1>${escapeHtml(location.name)}</h1><span></span></header>
      <div class="filters"><button data-filter="all" class="${state.filter === 'all' ? 'active' : ''}">Alle</button><button data-filter="open" class="${state.filter === 'open' ? 'active' : ''}">Nur offene</button><button data-filter="done" class="${state.filter === 'done' ? 'active' : ''}">Erledigte</button></div>
      <div class="meterList">${meters.length ? meters.map(m => `<button class="meter ${m.currentValue !== null ? 'done' : 'open'}" data-meter="${escapeHtml(m.id)}"><span class="meterNumber">${escapeHtml(m.number)}</span><span>${m.currentValue !== null ? `✓ ${formatNumber(m.currentValue)}` : 'Offen'}${m.changed ? ' · geändert' : ''}</span></button>`).join('') : '<p>In diesem Filter sind keine Zähler vorhanden.</p>'}</div></main>`;
    app.querySelector('[data-action="back"]').addEventListener('click', () => { state.screen = 'locations'; render(); });
    app.querySelectorAll('[data-filter]').forEach(btn => btn.addEventListener('click', () => { state.filter = btn.dataset.filter; render(); }));
    app.querySelectorAll('[data-meter]').forEach(btn => btn.addEventListener('click', () => { state.selectedMeterId = btn.dataset.meter; state.input = ''; state.screen = 'entry'; clearMessage(); render(); }));
  }

  function renderEntry() {
    const meter = selectedMeter();
    app.innerHTML = `<main class="page entry"><header class="topbar"><button class="button ghost compact" data-action="back">Zurück</button><h1>${escapeHtml(meter.location)}</h1><span></span></header>${messageHtml()}
      <div class="entryGrid"><section class="card details"><h2>Zähler ${escapeHtml(meter.number)}</h2><dl><div><dt>Vormonat</dt><dd>${formatNumber(meter.previousValue)}</dd></div><div><dt>Aktualwert</dt><dd>${formatNumber(meter.currentValue)}</dd></div><div><dt>Info</dt><dd>${escapeHtml(meter.info || '—')}</dd></div></dl></section>
      <section class="card inputCard"><label for="valueInput">Neuer Zählerstand</label><input id="valueInput" class="valueInput" inputmode="decimal" autocomplete="off" placeholder="0,0" value="${escapeHtml(state.input)}">
      <div class="keypad">${['1','2','3','4','5','6','7','8','9',',','0','⌫'].map(k => `<button data-key="${k}">${k}</button>`).join('')}</div><button class="button primary" data-action="save">Eingabe übernehmen</button></section></div></main>`;
    const input = document.getElementById('valueInput');
    input.addEventListener('input', e => { state.input = e.target.value; });
    app.querySelector('[data-action="back"]').addEventListener('click', () => { state.screen = 'meters'; clearMessage(); render(); });
    app.querySelector('[data-action="save"]').addEventListener('click', saveMeterValue);
    app.querySelectorAll('[data-key]').forEach(btn => btn.addEventListener('click', () => {
      state.input = btn.dataset.key === '⌫' ? state.input.slice(0, -1) : state.input + btn.dataset.key;
      input.value = state.input; input.focus();
    }));
  }

  function render() {
    if (!state.session || state.screen === 'home') return renderHome();
    if (state.screen === 'locations') return renderLocations();
    if (state.screen === 'meters' && selectedLocation()) return renderMeters();
    if (state.screen === 'entry' && selectedMeter()) return renderEntry();
    state.screen = 'home'; renderHome();
  }

  if ('serviceWorker' in navigator) window.addEventListener('load', () => navigator.serviceWorker.register('./sw.js').catch(() => {}));
  render();
})();
