const statCols = [
  ["contact_text", "contact_tone"],
  ["whiff_text", "whiff_tone"],
  ["chase_text", "chase_tone"],
  ["execution_text", "execution_tone"],
  ["risp_text", "risp_tone"],
  ["move_text", "move_tone"],
  ["qab_text", "qab_tone"],
];

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

function renderTable(view) {
  const body = document.querySelector("#sheet tbody");
  body.innerHTML = "";
  for (const row of view.hitters) {
    const tr = document.createElement("tr");
    if (row.no_pa) tr.className = "noppa";
    const name = document.createElement("td");
    name.className = "name";
    const mark = row.archetype_mark ? "<sup>1</sup>" : "";
    const arch = row.archetype ? `${row.archetype}${mark}` : "Unclassified";
    const metric = row.contact_metric === "Damage" ? "damage rate" : row.contact_metric === "Barrel" ? "barrel rate" : "";
    name.innerHTML = `<strong>${row.name}</strong><span>${row.pos} · ${arch}${metric ? ` · ${metric}` : ""}</span>`;
    tr.appendChild(name);
    for (const [textKey, toneKey] of statCols) {
      tr.appendChild(cell(row[textKey], row[toneKey]));
    }
    const pa = document.createElement("td");
    pa.textContent = row.no_pa ? "No PA" : String(row.pa);
    tr.appendChild(pa);
    const games = document.createElement("td");
    games.textContent = row.no_pa ? "" : String(row.games);
    tr.appendChild(games);
    body.appendChild(tr);
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
  for (const view of manifest.views) {
    const opt = document.createElement("option");
    opt.value = view.file;
    opt.textContent = view.label;
    opt.selected = view.id === manifest.default;
    select.appendChild(opt);
  }
  async function show() {
    status.textContent = "Loading…";
    const view = await loadView(select.value);
    renderLeague(view);
    renderTable(view);
    document.getElementById("print-head").textContent = `Twins Advance · ${view.label}`;
    status.textContent = `${view.label}: ${view.hitters.length} hitters`;
  }
  document.getElementById("export-pdf").addEventListener("click", () => {
    const label = select.options[select.selectedIndex]?.textContent || "Hitter sheet";
    document.getElementById("print-head").textContent = `Twins Advance · ${label}`;
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
    show().catch((err) => {
      status.textContent = err.message;
    });
  });
  await show();
}

main().catch((err) => {
  document.getElementById("status").textContent = err.message;
});
