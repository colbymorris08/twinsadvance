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
  const starts = view.hitters.map((h) => h.date_start).filter(Boolean).sort();
  const ends = view.hitters.map((h) => h.date_end).filter(Boolean).sort();
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

function appendRow(body, row, note, marks) {
  const tr = document.createElement("tr");
  if (row.no_pa) tr.className = "noppa";
  const name = document.createElement("td");
  name.className = "name";
  const mark = marks && row.archetype_mark ? "<sup>1</sup>" : "";
  const arch = row.archetype ? `${row.archetype}${mark}` : "Unclassified";
  const metric = row.contact_metric === "Damage" ? "damage rate" : row.contact_metric === "Barrel" ? "barrel rate" : "";
  let detail = note || "";
  if (row.no_pa && !/No PA/.test(detail)) detail = detail ? `${detail} · No PA` : "No PA";
  const sub = [row.pos, arch, metric, detail].filter(Boolean).join(" · ");
  name.innerHTML = `<strong>${row.name}</strong><span>${sub}</span>`;
  tr.appendChild(name);
  for (const [textKey, toneKey] of statCols) {
    tr.appendChild(cell(row[textKey], row[toneKey]));
  }
  body.appendChild(tr);
}

function renderTable(view, selector, noteFor) {
  const body = document.querySelector(selector);
  body.innerHTML = "";
  for (const row of view.hitters) {
    appendRow(body, row, noteFor ? noteFor(row) : "", true);
  }
}

function renderPrint(monthView, lastView) {
  document.getElementById("print-month").textContent = monthLine(monthView);
  document.getElementById("print-last10").textContent = last10Line(lastView);
  const body = document.querySelector("#print-table tbody");
  body.innerHTML = "";
  const lastById = new Map(lastView.hitters.map((h) => [h.id, h]));
  const seen = new Set();
  const monthName = monthView.label.replace(/ 2026$/, "");
  for (const row of monthView.hitters) {
    seen.add(row.id);
    appendRow(body, row, monthName, false);
    const last = lastById.get(row.id);
    if (last) appendRow(body, last, `Last 10 · ${hitterDates(last)}`, false);
  }
  for (const row of lastView.hitters) {
    if (!seen.has(row.id)) appendRow(body, row, `Last 10 · ${hitterDates(row)}`, false);
  }
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
  renderGroups(manifest.archetypes || []);

  const select = document.getElementById("month");
  for (const [id, label, , , file] of MONTHS) {
    const opt = document.createElement("option");
    opt.value = file;
    opt.textContent = label;
    opt.dataset.id = id;
    opt.selected = id === "october";
    select.appendChild(opt);
  }

  let monthView = null;
  const lastView = await loadView("last10.json");
  document.getElementById("last10-line").textContent = last10Line(lastView);
  renderTable(lastView, "#last10-sheet tbody", (row) => hitterDates(row));

  async function showMonth() {
    status.textContent = "Loading…";
    monthView = await loadView(select.value);
    const meta = MONTHS.find((row) => row[4] === select.value);
    if (meta) {
      monthView.label = meta[1];
      monthView.start = monthView.start || meta[2];
      monthView.end = monthView.end || meta[3];
    }
    renderLeague(monthView);
    renderTable(monthView, "#sheet tbody");
    document.getElementById("month-line").textContent = monthLine(monthView);
    renderPrint(monthView, lastView);
    status.textContent = `${monthView.label}: ${monthView.hitters.length} hitters · Last 10 games: ${lastView.hitters.length} hitters`;
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
