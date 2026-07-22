import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";

const XLSX = window.XLSX;
const STORAGE_KEY = "zaehlergang-session-v1";
const NORMAL_LOCATION_MAX_COLUMN = 72;
const TEN_O_CLOCK_COLUMNS = new Set([59, 60, 65, 66, 67, 68, 69, 70]);

function colLetter(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function numeric(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") return value;
  const parsed = Number(String(value).replace(/\./g, "").replace(",", "."));
  return Number.isFinite(parsed) ? parsed : null;
}

function formatValue(value) {
  const n = numeric(value);
  return n === null ? "–" : new Intl.NumberFormat("de-DE", { maximumFractionDigits: 4 }).format(n);
}

function cellValue(sheet, address) {
  return sheet?.[address]?.v ?? "";
}

function parseWorkbook(arrayBuffer, fileName) {
  const workbook = XLSX.read(arrayBuffer, { type: "array", cellStyles: true, cellDates: true });
  const dataSheet = workbook.Sheets["Daten_Ausdruck"];
  const sortedSheet = workbook.Sheets["Zaehlerliste_Sortiert"];
  if (!dataSheet || !sortedSheet) throw new Error("Die Arbeitsblätter Daten_Ausdruck und Zaehlerliste_Sortiert werden benötigt.");

  const locations = [];
  for (let col = 1; col <= 72; col += 1) {
    const letter = colLetter(col);
    const name = String(cellValue(sortedSheet, `${letter}1`) || "").trim();
    if (!name) continue;

    const meters = [];
    for (let row = 2; row <= 41; row += 1) {
      const meterRaw = cellValue(sortedSheet, `${letter}${row}`);
      if (meterRaw === "" || meterRaw === null || meterRaw === undefined) continue;
      const meterNumber = String(meterRaw).trim();
      const dataRow = Number(meterNumber) + 4;
      if (!Number.isFinite(dataRow) || dataRow < 5) continue;
      meters.push({
        id: `${letter}-${meterNumber}`,
        number: meterNumber,
        row: dataRow,
        info: String(cellValue(sortedSheet, `${letter}${row + 43}`) || "").trim(),
        timestamp: cellValue(dataSheet, `B${dataRow}`) || "",
        current: cellValue(dataSheet, `C${dataRow}`) || "",
        previous: cellValue(dataSheet, `D${dataRow}`) || "",
        changed: false
      });
    }
    if (!meters.length) continue;
    locations.push({ id: letter, name, column: col, group: TEN_O_CLOCK_COLUMNS.has(col) ? "ten" : "normal", meters });
  }

  return { workbook, fileName, locations, date: cellValue(dataSheet, "B3") || new Date() };
}

function sessionForStorage(session) {
  return session ? { fileName: session.fileName, date: session.date, locations: session.locations } : null;
}

function App() {
  const [screen, setScreen] = useState("start");
  const [session, setSession] = useState(null);
  const [selectedLocationId, setSelectedLocationId] = useState(null);
  const [selectedMeterId, setSelectedMeterId] = useState(null);
  const [filter, setFilter] = useState("all");
  const [input, setInput] = useState("");
  const [dialog, setDialog] = useState(null);
  const [toast, setToast] = useState("");
  const [isDirty, setIsDirty] = useState(false);
const [lastSavedAt, setLastSavedAt] = useState(null);
  const fileInput = useRef(null);

  useEffect(() => {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        setSession(parsed);
      } catch { localStorage.removeItem(STORAGE_KEY); }
    }
  }, []);

  useEffect(() => {
    if (session) localStorage.setItem(STORAGE_KEY, JSON.stringify(sessionForStorage(session)));
  }, [session]);

  useEffect(() => {
    if (!toast) return;
    const id = setTimeout(() => setToast(""), 2200);
    return () => clearTimeout(id);
  }, [toast]);

  const selectedLocation = session?.locations.find(x => x.id === selectedLocationId) ?? null;
  const selectedMeter = selectedLocation?.meters.find(x => x.id === selectedMeterId) ?? null;

  const totals = useMemo(() => {
    const meters = session?.locations.flatMap(x => x.meters) ?? [];
    const done = meters.filter(x => numeric(x.current) !== null).length;
    return { all: meters.length, done, percent: meters.length ? Math.round(done / meters.length * 100) : 0 };
  }, [session]);

  function showToast(message) { setToast(message); }

  async function openFile(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const buffer = await file.arrayBuffer();
      const parsed = parseWorkbook(buffer, file.name);
      setSession(parsed);
      setScreen("locations");
      showToast("Excel-Datei wurde geladen.");
    } catch (error) {
      setDialog({ title: "Datei konnte nicht geöffnet werden", message: error.message, confirmOnly: true });
    } finally {
      event.target.value = "";
    }
  }

  function openLocation(location) {
    setSelectedLocationId(location.id);
    setSelectedMeterId(null);
    setFilter("all");
    setScreen("meters");
  }

  function openMeter(meter) {
    setSelectedMeterId(meter.id);
    setInput("");
    setScreen("entry");
  }

  function appendKey(key) {
    if (key === "back") setInput(v => v.slice(0, -1));
    else if (key === "," && input.includes(",")) return;
    else setInput(v => v + key);
  }

  function commitValue() {
    if (!selectedLocation || !selectedMeter) return;
    const value = numeric(input);
    if (value === null) {
      setDialog({ title: "Ungültige Eingabe", message: "Bitte einen gültigen Zählerwert eingeben.", confirmOnly: true });
      return;
    }
    const previous = numeric(selectedMeter.previous);
    const currentExists = numeric(selectedMeter.current) !== null;
    const warnings = [];
    if (previous !== null && value < previous) warnings.push("Der neue Wert ist kleiner als der Vormonatswert.");
    if (previous !== null && value > previous * 1.2) warnings.push("Der neue Wert ist mehr als 20 % größer als der Vormonatswert.");
    if (selectedMeter.number === "400" && previous !== null && value === previous) warnings.push("Zähler 400 ist unverändert. Bitte Motor/Spannungsversorgung kontrollieren und Bergmann informieren.");
    const message = [currentExists ? "Für diesen Zähler existiert bereits ein Aktualwert und wird überschrieben." : "Der Wert wird übernommen.", ...warnings].join("\n\n");
    setDialog({
      title: warnings.length ? "Bitte Eingabe prüfen" : "Eingabe bestätigen",
      message,
      onConfirm: () => saveValue(value, currentExists)
    });
  }

  function saveValue(value, changed) {
  const timestamp = new Date().toISOString();

  setSession(old => ({
    ...old,
    locations: old.locations.map(location =>
      location.id !== selectedLocationId
        ? location
        : {
            ...location,
            meters: location.meters.map(meter =>
              meter.id !== selectedMeterId
                ? meter
                : {
                    ...meter,
                    current: value,
                    timestamp,
                    changed: changed || meter.changed
                  }
            )
          }
    )
  }));

  setIsDirty(true);
  setInput("");
  setDialog(null);
  showToast("Wert übernommen – noch nicht gespeichert.");
}

  function startNewSession() {
    if (!session) return;
    setDialog({
      title: "Neuen Zählergang beginnen",
      message: "Die aktuellen Werte werden als Vormonatswerte übernommen. Aktualwerte und Zeitstempel werden geleert.",
      onConfirm: () => {
        const now = new Date();
        setSession(old => ({
          ...old,
          fileName: makeNewFileName(old.fileName, now),
          date: now.toISOString(),
          locations: old.locations.map(location => ({
            ...location,
            meters: location.meters.map(meter => ({ ...meter, previous: meter.current !== "" ? meter.current : meter.previous, current: "", timestamp: "", changed: false }))
          }))
        }));
        setDialog(null);
        setScreen("locations");
        showToast("Neuer Zählergang wurde angelegt.");
      }
    });
  }

  function makeNewFileName(name, date) {
    const base = `Zählergang_${date.toISOString().slice(0,10)}`;
    const ext = name?.toLowerCase().endsWith(".xlsm") ? ".xlsm" : ".xlsx";
    return `${base}_001${ext}`;
  }

  function exportExcel() {
  if (!session?.workbook) {
    setDialog({
      title: "Speichern nicht möglich",
      message: "Bitte die letzte Excel-Datei erneut öffnen.",
      confirmOnly: true
    });
    return;
  }

  const wb = session.workbook;
  const dataSheet = wb.Sheets["Daten_Ausdruck"];

  session.locations
    .flatMap(location => location.meters)
    .forEach(meter => {
      XLSX.utils.sheet_add_aoa(
        dataSheet,
        [[
          meter.timestamp || "",
          meter.current === "" ? "" : numeric(meter.current),
          meter.previous === "" ? "" : numeric(meter.previous)
        ]],
        { origin: `B${meter.row}` }
      );
    });

  XLSX.utils.sheet_add_aoa(
    dataSheet,
    [[new Date(session.date)]],
    { origin: "B3" }
  );

  XLSX.writeFile(
    wb,
    session.fileName || "Zaehlergang.xlsx",
    { bookType: "xlsx" }
  );

  const savedAt = new Date();

  setIsDirty(false);
  setLastSavedAt(savedAt);

  showToast(
    `Excel gespeichert um ${savedAt.toLocaleTimeString("de-DE", {
      hour: "2-digit",
      minute: "2-digit"
    })}`
  );
}
  const normalLocations = session?.locations.filter(x => x.group === "normal") ?? [];
  const tenLocations = session?.locations.filter(x => x.group === "ten") ?? [];
  const filteredMeters = selectedLocation?.meters.filter(m => filter === "all" || (filter === "open" ? numeric(m.current) === null : numeric(m.current) !== null)) ?? [];

  return React.createElement("div", { className: "app" },
    React.createElement("header", { className: "header" },
      React.createElement("div", null, React.createElement("p", { className: "eyebrow" }, "Digitale Zählerablesung"), React.createElement("h1", null, "Zählergang")),
      React.createElement("div", { className: "status-pill" }, session?.fileName || "Keine Datei geöffnet")
    ),
    React.createElement("main", { className: "main" },
      screen === "start" && React.createElement(StartScreen, { session, onOpen: () => fileInput.current?.click(), onNew: startNewSession, onContinue: () => setScreen("locations"), onExport: exportExcel }),
      screen === "locations" && React.createElement(LocationsScreen, { session, totals, normalLocations, tenLocations, onBack: () => setScreen("start"), onOpen: openLocation, onExport: exportExcel }),
      screen === "meters" && React.createElement(MetersScreen, { location: selectedLocation, meters: filteredMeters, filter, onFilter: setFilter, onBack: () => setScreen("locations"), onOpen: openMeter }),
      screen === "entry" && React.createElement(EntryScreen, { location: selectedLocation, meter: selectedMeter, input, onInput: setInput, onKey: appendKey, onBack: () => setScreen("meters"), onSelect: openMeter, onSave: commitValue })
    ),
    React.createElement("input", { ref: fileInput, type: "file", accept: ".xlsx,.xlsm", hidden: true, onChange: openFile }),
    dialog && React.createElement(Dialog, { dialog, onClose: () => setDialog(null) }),
    toast && React.createElement("div", { className: "toast" }, toast)
  );
}

function StartScreen({ session, onOpen, onNew, onContinue, onExport }) {
  return React.createElement("section", { className: "card hero" },
    React.createElement("h2", null, "Zählergang starten"),
    React.createElement("p", null, "Öffne eine vorhandene Excel-Datei. Danach kannst du Standorte und Zähler bearbeiten, einen neuen Zählergang beginnen und die fertige Excel-Datei sichern."),
    React.createElement("div", { className: "actions" },
      React.createElement("button", { className: "btn btn-primary", onClick: onOpen }, "Excel-Datei öffnen"),
      React.createElement("button", { className: "btn btn-secondary", disabled: !session, onClick: onNew }, "Neuer Zählergang"),
      session && React.createElement("button", { className: "btn btn-secondary", onClick: onContinue }, "Zwischenstand fortsetzen"),
      session && React.createElement("button", { className: "btn btn-secondary", onClick: onExport }, "Excel sichern")
    )
  );
}

function LocationsScreen({ session, totals, normalLocations, tenLocations, onBack, onOpen, onExport }) {
  return React.createElement(React.Fragment, null,
    React.createElement("div", { className: "toolbar" }, React.createElement("button", { className: "back", onClick: onBack }, "Zurück"), React.createElement("div", null, React.createElement("p", { className: "eyebrow" }, "Aktueller Zählergang"), React.createElement("h2", null, session?.fileName))),
    React.createElement("div", { className: "card progress" }, React.createElement("div", { className: "progress-row" }, React.createElement("span", null, "Gesamtfortschritt"), React.createElement("strong", null, `${totals.percent} %`)), React.createElement("div", { className: "track" }, React.createElement("div", { className: "bar", style: { width: `${totals.percent}%` } })), React.createElement("span", null, `${totals.done} von ${totals.all} Zählern bearbeitet`), React.createElement("div", { className: "footer-actions" }, React.createElement("button", { className: "btn btn-secondary", onClick: onExport }, "Excel sichern"))),
    React.createElement("h3", { className: "section-title" }, "Standorte"), React.createElement("div", { className: "grid" }, normalLocations.map(l => React.createElement(LocationCard, { key: l.id, location: l, onClick: () => onOpen(l) }))),
    tenLocations.length > 0 && React.createElement(React.Fragment, null, React.createElement("h3", { className: "section-title" }, "10-Uhr-Ablesungen"), React.createElement("div", { className: "grid" }, tenLocations.map(l => React.createElement(LocationCard, { key: l.id, location: l, onClick: () => onOpen(l) }))))
  );
}

function LocationCard({ location, onClick }) {
  const done = location.meters.filter(m => numeric(m.current) !== null).length;
  const percent = Math.round(done / location.meters.length * 100);
  const cls = done === location.meters.length ? "done" : done > 0 ? "partial" : "";
  return React.createElement("button", { className: `location ${cls}`, onClick }, React.createElement("strong", null, location.name), React.createElement("small", null, `${done} / ${location.meters.length} Zähler · ${percent} %`));
}

function MetersScreen({ location, meters, filter, onFilter, onBack, onOpen }) {
  return React.createElement(React.Fragment, null,
    React.createElement("div", { className: "toolbar" }, React.createElement("button", { className: "back", onClick: onBack }, "Zurück"), React.createElement("div", null, React.createElement("p", { className: "eyebrow" }, "Standort"), React.createElement("h2", null, location?.name || "Zähler"))),
    React.createElement("div", { className: "filters" }, [["all","Alle"],["open","Nur offene"],["completed","Erledigte"]].map(([id,label]) => React.createElement("button", { key: id, className: `filter ${filter === id ? "active" : ""}`, onClick: () => onFilter(id) }, label))),
    React.createElement("div", { className: "meter-list" }, meters.length ? meters.map(m => { const done = numeric(m.current) !== null; return React.createElement("button", { key: m.id, className: `meter-row ${done ? "done" : ""} ${m.changed ? "changed" : ""}`, onClick: () => onOpen(m) }, React.createElement("div", { className: "meter-meta" }, React.createElement("span", { className: `dot ${m.changed ? "changed" : done ? "done" : ""}` }), React.createElement("strong", null, m.number)), React.createElement("span", null, m.changed ? "geändert" : done ? "erledigt" : "offen")); }) : React.createElement("div", { className: "empty" }, "Keine Zähler in diesem Filter."))
  );
}

function EntryScreen({ location, meter, input, onInput, onKey, onBack, onSelect, onSave }) {
  if (!location || !meter) return null;
  const p = numeric(meter.previous); const n = numeric(input); let notice = "";
  if (n !== null && p !== null && n < p) notice = "Der neue Wert ist kleiner als der Vormonatswert.";
  else if (n !== null && p !== null && n > p * 1.2) notice = "Der neue Wert ist mehr als 20 % größer als der Vormonatswert.";
  return React.createElement(React.Fragment, null,
    React.createElement("div", { className: "toolbar" }, React.createElement("button", { className: "back", onClick: onBack }, "Zurück"), React.createElement("div", null, React.createElement("p", { className: "eyebrow" }, location.name), React.createElement("h2", null, `Zähler ${meter.number}`))),
    React.createElement("div", { className: "entry-layout" },
      React.createElement("aside", { className: "card sidebar" }, React.createElement("h3", null, "Zählernummern"), location.meters.map(m => React.createElement("button", { key: m.id, className: `compact-meter ${m.id === meter.id ? "active" : ""}`, onClick: () => onSelect(m) }, React.createElement("span", null, m.number), React.createElement("span", null, numeric(m.current) !== null ? "✓" : "○")))),
      React.createElement("section", { className: "card entry" }, React.createElement("h3", null, "Zählerwert"), React.createElement("div", { className: "value-grid" }, React.createElement("div", { className: "value" }, React.createElement("span", null, "Vormonat"), React.createElement("strong", null, formatValue(meter.previous))), React.createElement("div", { className: "value" }, React.createElement("span", null, "Aktualwert"), React.createElement("strong", null, formatValue(meter.current)))), React.createElement("div", { className: "info" }, React.createElement("strong", null, "Info"), React.createElement("div", null, meter.info || "Keine Information hinterlegt.")), React.createElement("input", { className: "reading", inputMode: "decimal", value: input, placeholder: "0,0", onChange: e => onInput(e.target.value.replace(/[^0-9,]/g, "")) }), notice && React.createElement("div", { className: "notice" }, notice), React.createElement("div", { className: "keypad" }, ["7","8","9","4","5","6","1","2","3",",","0","back"].map(k => React.createElement("button", { key: k, onClick: () => onKey(k) }, k === "back" ? "⌫" : k))), React.createElement("button", { className: "btn btn-primary", style: { width: "100%" }, onClick: onSave }, "Eingabe übernehmen"))
    )
  );
}

function Dialog({ dialog, onClose }) {
  return React.createElement("div", { className: "dialog-backdrop" }, React.createElement("div", { className: "dialog" }, React.createElement("h3", null, dialog.title), React.createElement("p", { style: { whiteSpace: "pre-line" } }, dialog.message), React.createElement("div", { className: "dialog-actions" }, !dialog.confirmOnly && React.createElement("button", { className: "btn btn-secondary", onClick: onClose }, "Abbrechen"), React.createElement("button", { className: "btn btn-primary", onClick: dialog.onConfirm || onClose }, dialog.confirmOnly ? "OK" : "Bestätigen"))));
}

createRoot(document.getElementById("root")).render(React.createElement(App));

if ("serviceWorker" in navigator) window.addEventListener("load", () => navigator.serviceWorker.register("service-worker.js").catch(console.warn));
