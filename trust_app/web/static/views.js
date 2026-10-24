// Rendering: each function turns the summary from /api/data into HTML for one section of the page.
export const DIMENSION_COLOUR = {
  "DQ Score": "var(--d-DQ)", Glossary: "var(--d-Glossary)", Ownership: "var(--d-Ownership)", Metadata: "var(--d-Metadata)",
};
const DQ_HEADERS = ["#2E75B6", "#00A3A1", "#8E7CC3", "#C00000", "#7F7F7F"];
const DIMENSIONS = ["DQ Score", "Glossary", "Ownership", "Metadata"];
const RING_LENGTH = 2 * Math.PI * 70;

export const esc = (text) => String(text ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

export function bandClass(value, data) {
  const [, good, ok] = data.bands.map((b) => b.min);
  return value >= good ? "good" : value >= ok ? "warn" : "bad";
}
const bandColour = { good: "#3aa655", warn: "#e0a800", bad: "#e5484d" };

export function renderHero(el, data) {
  const level = bandClass(data.score, data);
  const offset = RING_LENGTH * (1 - Math.min(data.score, 100) / 100);
  const score = `<div class="tile score"><h3>TRUST SCORE: ${esc(data.table)}</h3>
    <div class="ring"><svg width="170" height="170" viewBox="0 0 170 170">
      <circle class="track" cx="85" cy="85" r="70" fill="none" stroke-width="14"/>
      <circle class="bar" cx="85" cy="85" r="70" fill="none" stroke-width="14" stroke-linecap="round"
        stroke="${bandColour[level]}" stroke-dasharray="${RING_LENGTH}" stroke-dashoffset="${offset}"/></svg>
      <div class="value"><b>${data.score.toFixed(1)}</b><small>out of 100</small></div></div>
    <span class="pill ${level}">${esc(data.band)}</span></div>`;
  const cards = data.cards.map((c) => {
    const tone = bandClass(c.pct, data);
    return `<div class="tile"><h3 style="background:${DIMENSION_COLOUR[c.name] || "#555"}">${esc(c.name.toUpperCase())}</h3>
      <div class="num">${Math.round(c.pct)}%</div>
      <div class="meter"><i style="width:${c.pct}%;background:${bandColour[tone]}"></i></div>
      <div class="sub">weight ${Math.round(c.weight * 100)}%<br>${c.yes} of ${c.yes + c.no} CDEs</div></div>`;
  });
  el.innerHTML = score + cards.join("");
}

export function renderBands(el, data) {
  el.innerHTML = data.bands.map((b, i) => {
    const upper = i ? data.bands[i - 1].min - 1 : 100;
    const range = i < data.bands.length - 1 ? `${b.min}-${upper}` : `below ${data.bands[i - 1].min}`;
    return `<div style="background:#${b.colour}">${range}&nbsp; ${esc(b.label)}</div>`;
  }).join("");
}

export function renderGrid(el, data) {
  const heads = data.dims.map((d, i) =>
    `<th style="background:${d === "Timeliness" ? "#ED7D31" : DQ_HEADERS[i % DQ_HEADERS.length]}">${esc(d)}</th>`).join("");
  const body = data.grid.map((row) => {
    const cells = row.values.map((v) => v == null ? "<td>-</td>" : `<td class="${bandClass(v, data)}">${Math.round(v)}%</td>`);
    return `<tr class="${row.table === "All Domain" ? "all" : ""}"><td>${esc(row.table)}</td>${cells.join("")}</tr>`;
  }).join("");
  el.innerHTML = `<thead><tr><th>Domain</th>${heads}</tr></thead><tbody>${body}</tbody>`;
}

export function renderChartData(el, data) {
  const body = data.cards.map((c) => `<tr><td><span class="chipdim" style="background:${DIMENSION_COLOUR[c.name]}">${esc(c.name)}</span></td>
    <td class="good">${c.yes}</td><td class="bad">${c.no}</td><td class="${bandClass(c.pct, data)}">${Math.round(c.pct)}%</td>
    <td>${Math.round(c.weight * 100)}%</td></tr>`).join("");
  el.innerHTML = `<thead><tr><th>Dimension</th><th>Trusted</th><th>Not Trusted</th><th>Coverage %</th><th>Weight</th></tr></thead><tbody>${body}</tbody>`;
}

export function renderRules(el, data) {
  el.innerHTML = data.rules.length
    ? data.rules.map((r) => `<div class="rule"><div class="name">${esc(r.label)}</div>
        <div class="track"><div class="fill" style="width:${r.score}%"></div><b>${r.score}</b></div></div>`).join("")
    : '<div class="empty">No scored DQ rules for this selection.</div>';
}

export function fillDimensionFilter(select) {
  select.innerHTML = '<option value="">All dimensions</option>' + DIMENSIONS.map((d) => `<option>${d}</option>`).join("");
}

export function renderCoverage(el, countEl, data, { query, dimension, yn }) {
  const rows = data.rows.filter((r) =>
    (!query || Object.values(r).join(" ").toLowerCase().includes(query)) &&
    (!yn || (dimension ? r[dimension] === yn : DIMENSIONS.some((d) => r[d] === yn))));
  countEl.textContent = `(${rows.length} of ${data.rows.length})`;
  const rules = (r) => r["DQ Rule Name"]
    ? r["DQ Rule Name"].split("; ").map((n, i) => `${n} (${r["DQ Dimension"].split("; ")[i]}: ${r["DQ Rule Score"].split("; ")[i]})`).join("; ")
    : "-";
  const head = DIMENSIONS.map((d) => `<th style="background:${DIMENSION_COLOUR[d]}">${d}</th>`).join("");
  const body = rows.map((r) => `<tr><td>${esc(r["Table Name"])}</td><td class="l">${esc(r["CDE ColumnName"])}</td>
    <td class="l">${esc(r["Column Name"])}</td>${DIMENSIONS.map((d) => `<td><span class="yn ${r[d]}">${r[d]}</span></td>`).join("")}
    <td class="l">${esc(rules(r))}</td><td class="l">${esc(r["Glossary Name"] || "-")}</td><td class="l">${esc(r["Owner Name"] || "-")}</td></tr>`).join("");
  el.innerHTML = `<thead><tr><th>Table Name</th><th>CDE ColumnName</th><th>Column Name</th>${head}<th>DQ rules</th><th>Glossary term</th><th>Owner</th></tr></thead><tbody>${body}</tbody>`;
}

export function renderRawFiles(el, files) {
  el.innerHTML = files.length
    ? files.map((f, i) => `<button class="file" data-i="${i}"><span class="glyph">{ }</span>
        <span><b>${esc(f.name)}</b><small>${esc(f.modified)} &middot; ${esc(f.size)}</small></span></button>`).join("")
    : '<div class="empty">No raw JSON files yet. Use "Browse another file..." or save one from a Purview load.</div>';
}
