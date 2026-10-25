const statCols = [
  ["contact_text", "contact_tone"],
  ["whiff_text", "whiff_tone"],
  ["chase_text", "chase_tone"],
  ["execution_text", "execution_tone"],
  ["risp_text", "risp_tone"],
  ["move_text", "move_tone"],
  ["qab_text", "qab_tone"],
];

const MONTHS = [
  ["april", "April 2026", "2026-04-01", "2026-04-30", "april.json"],
  ["may", "May 2026", "2026-05-01", "2026-05-31", "may.json"],
  ["june", "June 2026", "2026-06-01", "2026-06-30", "june.json"],
  ["july", "July 2026", "2026-07-01", "2026-07-31", "july.json"],
  ["august", "August 2026", "2026-08-01", "2026-08-31", "august.json"],
  ["september", "September 2026", "2026-09-01", "2026-09-30", "september.json"],
  ["october", "October 2026", "2026-10-01", "2026-10-31", "october.json"],
];

const SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sept", "Oct", "Nov", "Dec"];

function cell(text, tone) {
  const td = document.createElement("td");
  td.className = `stat ${tone || "blank"}`;
  td.textContent = text || "";
  return td;
}

function pct(rate) {
  if (rate === null || rate === undefined) return "—";
  return `${(rate * 100).toFixed(1)}%`;
}

const EXEC_INPUTS = ["risp", "move", "qab"];
const EXEC_BAND = 0.015;

function finiteRate(value) {
  return typeof value === "number" && Number.isFinite(value);
}

function mean(values) {
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function hitterExecution(row) {
  const rates = EXEC_INPUTS.map((key) => row[key]).filter(finiteRate);
  if (!rates.length) return null;
  return mean(rates);
}

function leagueExecution(league) {
  const rates = EXEC_INPUTS.map((key) => (league || {})[key]).filter(finiteRate);
  if (rates.length !== EXEC_INPUTS.length) return null;
  return mean(rates);
}

function executionDisplay(row, league) {
  const value = hitterExecution(row);
  if (value === null) return { text: "", tone: "blank" };
  const baseline = leagueExecution(league);
  if (baseline === null) return { text: pct(value), tone: "blank" };
  const gap = value - baseline;
  if (gap > EXEC_BAND + 1e-12) return { text: pct(value), tone: "green" };
  if (gap < -(EXEC_BAND + 1e-12)) return { text: pct(value), tone: "red" };
  return { text: pct(value), tone: "white" };
}

function num(value, digits = 1) {
  if (value === null || value === undefined) return "—";
  return Number(value).toFixed(digits);
}

function fmt(iso) {
  if (!iso) return "";
  const [, month, day] = iso.split("-").map(Number);
  return `${SHORT[month - 1]} ${day}`;
}

function span(start, end) {
  if (!start || !end) return "";
  return `${fmt(start)}–${fmt(end)}`;
}

function monthLine(view) {
  const range = span(view.start, view.end);
  return range ? `${view.label}, ${range}` : view.label;
}

function last10Line(view) {
  const hitters = chartHitters(view);
  const starts = hitters.map((h) => h.date_start).filter(Boolean).sort();
  const ends = hitters.map((h) => h.date_end).filter(Boolean).sort();
  const range = span(starts[0], ends[ends.length - 1]);
  return range ? `Last 10 games, ${range}` : "Last 10 games";
}

function hitterDates(row) {
  if (!row.date_start || !row.date_end) return "No PA";
  return span(row.date_start, row.date_end);
}

function renderLeague(view) {
  const lg = view.league || {};
  const box = document.getElementById("league");
  box.innerHTML = `<strong>${view.comparison}</strong> · barrel ${pct(lg.barrel)} · damage ${pct(lg.damage)} · in-zone whiff ${pct(lg.whiff)} · chase ${pct(lg.chase)} · RISP ${pct(lg.risp)} · move-over ${pct(lg.move)} · QAB ${pct(lg.qab)} · avg EV ${num(lg.ev)} mph`;
}

function renderGroups(archetypes) {
  const box = document.getElementById("groups");
  const buckets = { Damage: [], Barrel: [], Unclassified: [] };
  for (const a of archetypes) {
    const key = a.archetype || "Unclassified";
    (buckets[key] || buckets.Unclassified).push(a);
  }
  box.innerHTML = "";
  for (const [label, rows] of Object.entries(buckets)) {
    const div = document.createElement("div");
    const title = document.createElement("h2");
    title.textContent = `${label} (${rows.length})`;
    const list = document.createElement("ul");
    for (const a of rows) {
      const li = document.createElement("li");
      const ev = a.ev === null || a.ev === undefined ? "EV n/a" : `${a.ev.toFixed(1)} mph EV`;
      li.textContent = `${a.name} · ${ev}`;
      list.appendChild(li);
    }
    div.append(title, list);
    box.appendChild(div);
  }
}

function hasPa(row) {
  return !!row && !row.no_pa;
}

function appendRow(body, row, note, marks, league) {
  const tr = document.createElement("tr");
  const name = document.createElement("td");
  name.className = "name";
  const mark = marks && row.archetype_mark ? "<sup>1</sup>" : "";
  const arch = row.archetype ? `${row.archetype}${mark}` : "Unclassified";
  const metric = row.contact_metric === "Damage" ? "damage rate" : row.contact_metric === "Barrel" ? "barrel rate" : "";
  const sub = [row.pos, arch, metric, note || ""].filter(Boolean).join(" · ");
  name.innerHTML = `<strong>${row.name}</strong><span>${sub}</span>`;
  tr.appendChild(name);
  const exec = executionDisplay(row, league);
  for (const [textKey, toneKey] of statCols) {
    if (textKey === "execution_text") tr.appendChild(cell(exec.text, exec.tone));
    else tr.appendChild(cell(row[textKey], row[toneKey]));
  }
  body.appendChild(tr);
}

let omitIds = new Set();
let ilNoteText = "";

function onChart(row) {
  return !omitIds.has(row.id);
}

function chartHitters(view) {
  return (view.hitters || []).filter(onChart);
}

function monthHasPa(view) {
  return chartHitters(view).some((row) => !row.no_pa);
}

function renderIlNote() {
  const text = ilNoteText;
  for (const id of ["il-note", "print-il-note"]) {
    const el = document.getElementById(id);
    el.textContent = text;
    el.hidden = !text;
  }
}

function renderPairs(body, monthView, lastView, marks) {
  body.innerHTML = "";
  const lastById = new Map(chartHitters(lastView).map((h) => [h.id, h]));
  const seen = new Set();
  const monthName = monthView.label.replace(/ 2026$/, "");
  const monthLeague = monthView.league || {};
  const lastLeague = lastView.league || {};
  for (const row of chartHitters(monthView)) {
    seen.add(row.id);
    if (hasPa(row)) appendRow(body, row, monthName, marks, monthLeague);
    const last = lastById.get(row.id);
    if (hasPa(last)) appendRow(body, last, `Last 10 · ${hitterDates(last)}`, marks, lastLeague);
  }
  for (const row of chartHitters(lastView)) {
    if (!seen.has(row.id) && hasPa(row)) appendRow(body, row, `Last 10 · ${hitterDates(row)}`, marks, lastLeague);
  }
}

function renderPrint(monthView, lastView) {
  document.getElementById("print-month").textContent = monthLine(monthView);
  document.getElementById("print-last10").textContent = last10Line(lastView);
  renderPairs(document.querySelector("#print-table tbody"), monthView, lastView, false);
  renderIlNote();
}

async function loadView(file) {
  const res = await fetch(`data/${file}`);
  if (!res.ok) throw new Error(`Could not load ${file}`);
  return res.json();
}

async function main() {
  const status = document.getElementById("status");
  const manifestRes = await fetch("data/manifest.json");
  if (!manifestRes.ok) throw new Error("Could not load data/manifest.json");
  const manifest = await manifestRes.json();
  document.getElementById("subtitle").textContent = manifest.subtitle || "";
  document.getElementById("roster").textContent = manifest.roster_note || "";
  document.getElementById("footnote1").textContent = manifest.footnote1 || "";
  const footnote2 = manifest.footnote2 || "";
  document.getElementById("footnote2").textContent = /^footnote\s*2/i.test(footnote2)
    ? footnote2
    : `Footnote 2. ${footnote2}`;
  document.getElementById("color").textContent = manifest.color_note || "";
  omitIds = new Set(manifest.il_omit_ids || []);
  ilNoteText = manifest.il_note || "";
  renderIlNote();
  renderGroups(manifest.archetypes || []);

  let initialFile = MONTHS[MONTHS.length - 1][4];
  for (let i = MONTHS.length - 1; i >= 0; i -= 1) {
    const file = MONTHS[i][4];
    try {
      const probe = await loadView(file);
      if (monthHasPa(probe)) {
        initialFile = file;
        break;
      }
    } catch (err) {
      /* try the previous month */
    }
  }

  const select = document.getElementById("month");
  for (const [id, label, , , file] of MONTHS) {
    const opt = document.createElement("option");
    opt.value = file;
    opt.textContent = label;
    opt.dataset.id = id;
    opt.selected = file === initialFile;
    select.appendChild(opt);
  }

  let monthView = null;
  const lastView = await loadView("last10.json");
  document.getElementById("last10").textContent = last10Line(lastView);

  async function showMonth() {
    status.textContent = "Loading…";
    const meta = MONTHS.find((row) => row[4] === select.value);
    let loadError = "";
    try {
      monthView = await loadView(select.value);
    } catch (err) {
      loadError = err.message;
      monthView = {
        label: meta ? meta[1] : "Month",
        start: meta ? meta[2] : "",
        end: meta ? meta[3] : "",
        comparison: "",
        league: {},
        hitters: [],
      };
    }
    if (meta) {
      monthView.label = meta[1];
      monthView.start = monthView.start || meta[2];
      monthView.end = monthView.end || meta[3];
    }
    renderLeague(monthView);
    renderPairs(document.querySelector("#sheet tbody"), monthView, lastView, true);
    document.getElementById("month-line").textContent = monthLine(monthView);
    renderPrint(monthView, lastView);
    const withPa = chartHitters(monthView).filter((row) => !row.no_pa).length;
    if (loadError) {
      status.textContent = `${loadError}. Showing last 10 games.`;
    } else if (withPa) {
      status.textContent = `${monthView.label}: ${withPa} hitters with month plate appearances, month row above last 10`;
    } else {
      status.textContent = `${monthView.label}: no plate appearances, showing last 10 games`;
    }
  }

  document.getElementById("export-pdf").addEventListener("click", () => {
    if (monthView) renderPrint(monthView, lastView);
    const label = select.options[select.selectedIndex]?.textContent || "Hitter sheet";
    const previous = document.title;
    document.title = `Twins Advance — ${label}`;
    const restore = () => {
      document.title = previous;
      window.removeEventListener("afterprint", restore);
    };
    window.addEventListener("afterprint", restore);
    window.print();
  });

  select.addEventListener("change", () => {
    showMonth().catch((err) => {
      status.textContent = err.message;
    });
  });
  await showMonth();
}

main().catch((err) => {
  document.getElementById("status").textContent = err.message;
});
