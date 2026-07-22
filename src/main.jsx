import React, { useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import './styles.css';

const STORAGE_KEY = 'zaehlergang-session-v001';
const TEN_OCLOCK_COLUMNS = new Set(['BG', 'BH', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR']);

function normalizeNumber(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  const text = String(value).trim().replace(/\s/g, '');
  const normalized = text.includes(',')
    ? text.replace(/\./g, '').replace(',', '.')
    : text;
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

function formatNumber(value) {
  if (value === null || value === undefined || value === '') return '—';
  return new Intl.NumberFormat('de-DE', { maximumFractionDigits: 3 }).format(value);
}

function toLocalDateString(date = new Date()) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function arrayBufferToBase64(buffer) {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize));
  }
  return btoa(binary);
}

function base64ToArrayBuffer(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

function parseWorkbook(file, buffer) {
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
  const dataSheet = workbook.Sheets.Daten_Ausdruck;
  const sortedSheet = workbook.Sheets.Zaehlerliste_Sortiert;
  if (!dataSheet || !sortedSheet) {
    throw new Error('Die Blätter „Daten_Ausdruck“ und „Zaehlerliste_Sortiert“ wurden nicht gefunden.');
  }

  const metersByNumber = new Map();
  for (let row = 5; row <= 524; row += 1) {
    const number = dataSheet[`A${row}`]?.v;
    if (number === undefined || number === null || number === '') continue;
    const numberText = String(number).trim();
    metersByNumber.set(numberText, {
      id: `${numberText}-${row}`,
      number: numberText,
      timestamp: dataSheet[`B${row}`]?.v ?? null,
      currentValue: normalizeNumber(dataSheet[`C${row}`]?.v),
      previousValue: normalizeNumber(dataSheet[`D${row}`]?.v),
      info: '',
      location: 'Unzugeordnet',
      sourceRow: row,
      changed: false
    });
  }

  const locations = [];
  const range = XLSX.utils.decode_range(sortedSheet['!ref'] || 'A1:A1');
  for (let col = 1; col <= range.e.c; col += 1) {
    const colName = XLSX.utils.encode_col(col);
    const title = sortedSheet[`${colName}1`]?.v;
    if (!title) continue;

    const meterNumbers = [];
    for (let row = 2; row <= 41; row += 1) {
      const value = sortedSheet[`${colName}${row}`]?.v;
      if (value !== undefined && value !== null && value !== '') meterNumbers.push(String(value).trim());
    }
    if (!meterNumbers.length) continue;

    const name = String(title).trim();
    locations.push({
      id: colName,
      name,
      group: TEN_OCLOCK_COLUMNS.has(colName) ? '10 Uhr' : 'Normal',
      meterNumbers
    });

    meterNumbers.forEach((meterNumber, index) => {
      const meter = metersByNumber.get(meterNumber);
      if (!meter) return;
      meter.location = name;
      meter.info = sortedSheet[`${colName}${45 + index}`]?.v ?? '';
    });
  }

  const workbookDate = dataSheet.B3?.v;
  const date = workbookDate instanceof Date && !Number.isNaN(workbookDate.valueOf())
    ? toLocalDateString(workbookDate)
    : toLocalDateString();

  return {
    fileName: file.name,
    date,
    meters: [...metersByNumber.values()],
    locations,
    workbookBase64: arrayBufferToBase64(buffer)
  };
}

function App() {
  const [session, setSession] = useState(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      return saved ? JSON.parse(saved) : null;
    } catch {
      return null;
    }
  });
  const [screen, setScreen] = useState(session ? 'locations' : 'home');
  const [selectedLocationId, setSelectedLocationId] = useState(null);
  const [selectedMeterId, setSelectedMeterId] = useState(null);
  const [filter, setFilter] = useState('all');
  const [input, setInput] = useState('');
  const [error, setError] = useState('');

  useEffect(() => {
    if (!session) return;
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(session));
    } catch {
      setError('Der Browser-Speicher ist voll. Bitte jetzt „Excel sichern“ verwenden.');
    }
  }, [session]);

  useEffect(() => {
    if ('serviceWorker' in navigator) navigator.serviceWorker.register('/sw.js').catch(() => {});
  }, []);

  const progress = useMemo(() => {
    if (!session?.meters?.length) return 0;
    return Math.round(session.meters.filter(m => m.currentValue !== null).length / session.meters.length * 100);
  }, [session]);

  const selectedLocation = session?.locations.find(location => location.id === selectedLocationId) ?? null;
  const selectedMeter = session?.meters.find(meter => meter.id === selectedMeterId) ?? null;

  async function importWorkbook(event) {
    const file = event.target.files?.[0];
    event.target.value = '';
    if (!file) return;
    setError('');
    try {
      const buffer = await file.arrayBuffer();
      const imported = parseWorkbook(file, buffer);
      setSession(imported);
      setSelectedLocationId(null);
      setSelectedMeterId(null);
      setScreen('locations');
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : 'Die Datei konnte nicht gelesen werden.');
    }
  }

  function createNewSession() {
    if (!session) return;
    const date = toLocalDateString();
    const meters = session.meters.map(meter => ({
      ...meter,
      previousValue: meter.currentValue ?? meter.previousValue,
      currentValue: null,
      timestamp: null,
      changed: false
    }));
    setSession({ ...session, date, fileName: `Zählergang ${date}.xlsx`, meters });
    setSelectedLocationId(null);
    setSelectedMeterId(null);
    setScreen('locations');
  }

  function saveValue() {
    if (!selectedMeter) return;
    const value = normalizeNumber(input);
    if (value === null) {
      setError('Bitte einen gültigen Zahlenwert eingeben.');
      return;
    }
    setError('');
    const previous = selectedMeter.previousValue;
    if (previous !== null && value < previous && !window.confirm('Dieser Wert ist kleiner als der vorherige. Trotzdem übernehmen?')) return;
    if (previous !== null && value > previous * 1.2 && !window.confirm('Der Wert ist über 20 % größer als der vorherige. Trotzdem übernehmen?')) return;
    if (selectedMeter.currentValue !== null && !window.confirm('Für diesen Zähler wurde bereits ein Wert eingetragen. Bestehenden Wert ersetzen?')) return;

    setSession(current => ({
      ...current,
      meters: current.meters.map(meter => meter.id === selectedMeter.id ? {
        ...meter,
        currentValue: value,
        timestamp: new Date().toISOString(),
        changed: meter.currentValue !== null
      } : meter)
    }));
    setInput('');
  }

  function exportWorkbook() {
    if (!session) return;
    const workbook = XLSX.read(base64ToArrayBuffer(session.workbookBase64), { type: 'array', cellDates: true });
    const sheet = workbook.Sheets.Daten_Ausdruck;
    sheet.B3 = { t: 'd', v: new Date(`${session.date}T12:00:00`) };

    session.meters.forEach(meter => {
      const row = meter.sourceRow;
      if (meter.timestamp) sheet[`B${row}`] = { t: 'd', v: new Date(meter.timestamp) };
      else delete sheet[`B${row}`];
      if (meter.currentValue !== null) sheet[`C${row}`] = { t: 'n', v: meter.currentValue };
      else delete sheet[`C${row}`];
      if (meter.previousValue !== null) sheet[`D${row}`] = { t: 'n', v: meter.previousValue };
      else delete sheet[`D${row}`];
    });

    const fileName = session.fileName.replace(/\.xlsm$/i, '.xlsx');
    XLSX.writeFile(workbook, fileName, { bookType: 'xlsx', compression: true });
  }

  function clearSavedSession() {
    if (!window.confirm('Den lokal gespeicherten Zwischenstand wirklich entfernen?')) return;
    localStorage.removeItem(STORAGE_KEY);
    setSession(null);
    setSelectedLocationId(null);
    setSelectedMeterId(null);
    setScreen('home');
  }

  if (screen === 'home') {
    return <main className="page home">
      <div className="brandMark">ZG</div>
      <h1>Zählergang</h1>
      <p>Lokale Zählererfassung für iPad und Windows.</p>
      {error && <div className="error" role="alert">{error}</div>}
      <label className="button primary buttonLike">
        Zählergang öffnen
        <input hidden type="file" accept=".xlsx,.xlsm" onChange={importWorkbook} />
      </label>
      <button className="button secondary" onClick={createNewSession} disabled={!session}>Neuer Zählergang</button>
      {session && <>
        <button className="button ghost" onClick={() => setScreen('locations')}>Letzten Stand fortsetzen</button>
        <button className="button dangerText" onClick={clearSavedSession}>Lokalen Stand entfernen</button>
      </>}
    </main>;
  }

  if (screen === 'locations') {
    return <main className="page">
      <header className="topbar">
        <button className="button ghost compact" onClick={() => setScreen('home')}>Start</button>
        <div className="fileMeta"><strong>{session.fileName}</strong><span>{progress} % abgeschlossen</span></div>
        <button className="button secondary compact" onClick={exportWorkbook}>Excel sichern</button>
      </header>
      {error && <div className="error" role="alert">{error}</div>}
      <div className="progress" aria-label={`${progress} Prozent abgeschlossen`}><span style={{ width: `${progress}%` }} /></div>
      {['Normal', '10 Uhr'].map(group => {
        const groupLocations = session.locations.filter(location => location.group === group);
        if (!groupLocations.length) return null;
        return <section key={group}>
          <h2>{group === 'Normal' ? 'Standorte' : '10-Uhr-Ablesungen'}</h2>
          <div className="grid">
            {groupLocations.map(location => {
              const meters = session.meters.filter(meter => meter.location === location.name);
              const done = meters.filter(meter => meter.currentValue !== null).length;
              const status = done === meters.length ? 'done' : done > 0 ? 'partial' : 'open';
              return <button key={location.id} className={`location ${status}`} onClick={() => {
                setSelectedLocationId(location.id);
                setFilter('all');
                setScreen('meters');
              }}>
                <strong>{location.name}</strong>
                <span>{done} / {meters.length} Zähler</span>
              </button>;
            })}
          </div>
        </section>;
      })}
    </main>;
  }

  if (screen === 'meters' && selectedLocation) {
    let meters = session.meters.filter(meter => meter.location === selectedLocation.name);
    if (filter === 'open') meters = meters.filter(meter => meter.currentValue === null);
    if (filter === 'done') meters = meters.filter(meter => meter.currentValue !== null);
    return <main className="page">
      <header className="topbar">
        <button className="button ghost compact" onClick={() => setScreen('locations')}>Zurück</button>
        <h1>{selectedLocation.name}</h1>
        <span />
      </header>
      <div className="filters" role="group" aria-label="Zähler filtern">
        <button onClick={() => setFilter('all')} className={filter === 'all' ? 'active' : ''}>Alle</button>
        <button onClick={() => setFilter('open')} className={filter === 'open' ? 'active' : ''}>Nur offene</button>
        <button onClick={() => setFilter('done')} className={filter === 'done' ? 'active' : ''}>Erledigte</button>
      </div>
      <div className="meterList">
        {meters.map(meter => <button key={meter.id} className={`meter ${meter.currentValue !== null ? 'done' : 'open'}`} onClick={() => {
          setSelectedMeterId(meter.id);
          setInput('');
          setScreen('entry');
        }}>
          <span className="meterNumber">{meter.number}</span>
          <span>{meter.currentValue !== null ? `✓ ${formatNumber(meter.currentValue)}` : 'Offen'}{meter.changed ? ' · geändert' : ''}</span>
        </button>)}
      </div>
    </main>;
  }

  if (screen === 'entry' && selectedMeter) {
    return <main className="page entry">
      <header className="topbar">
        <button className="button ghost compact" onClick={() => setScreen('meters')}>Zurück</button>
        <h1>{selectedMeter.location}</h1>
        <span />
      </header>
      {error && <div className="error" role="alert">{error}</div>}
      <div className="entryGrid">
        <section className="card details">
          <h2>Zähler {selectedMeter.number}</h2>
          <dl>
            <div><dt>Vormonat</dt><dd>{formatNumber(selectedMeter.previousValue)}</dd></div>
            <div><dt>Aktualwert</dt><dd>{formatNumber(selectedMeter.currentValue)}</dd></div>
            <div><dt>Info</dt><dd>{selectedMeter.info || '—'}</dd></div>
          </dl>
        </section>
        <section className="card inputCard">
          <label htmlFor="meterValue">Neuer Zählerstand</label>
          <input id="meterValue" className="valueInput" value={input} onChange={event => setInput(event.target.value)} inputMode="decimal" placeholder="0,0" autoFocus />
          <div className="keypad">
            {['1', '2', '3', '4', '5', '6', '7', '8', '9', ',', '0', '⌫'].map(key => <button key={key} onClick={() => key === '⌫' ? setInput(value => value.slice(0, -1)) : setInput(value => value + key)}>{key}</button>)}
          </div>
          <button className="button primary" onClick={saveValue}>Eingabe übernehmen</button>
        </section>
      </div>
    </main>;
  }

  return null;
}

createRoot(document.getElementById('root')).render(<App />);
