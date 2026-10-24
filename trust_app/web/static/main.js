// Wiring: buttons -> API calls -> views. Everything else lives in api.js and views.js.
import { api } from "./api.js";
import * as view from "./views.js";

const $ = (selector) => document.querySelector(selector);
const state = { data: null, timer: null, table: "All Domain" };
const SOURCE_BUTTONS = ["#fromPurview", "#fromRaw"];

// ---- theme -------------------------------------------------------------
$("#theme").onclick = () => {
  const root = document.documentElement;
  root.dataset.theme = root.dataset.theme === "dark" ? "light" : "dark";
};

// ---- status polling ----------------------------------------------------
function showMessage(status, text) {
  const box = $("#msg");
  box.className = "msg " + status;
  box.textContent = text;
}

async function syncStatus() {
  const s = await api.status();
  showMessage(s.status, s.message);
  $("#formula").textContent = "Trust Score = " + s.formula;
  $("#stamp").textContent = s.loaded_at ? `Last refresh: ${s.loaded_at.replace("T", " ")} - ${s.source}` : "Not refreshed yet";
  const busy = s.status === "loading", ready = s.status === "ready";
  SOURCE_BUTTONS.forEach((id) => { $(id).disabled = busy; });
  ["#xlsx", "#saveRaw"].forEach((id) => { $(id).disabled = !ready; });
  if (busy && !state.timer) state.timer = setInterval(syncStatus, 700);
  if (!busy && state.timer) {
    clearInterval(state.timer);
    state.timer = null;
    if (ready) await loadDashboard("All Domain");
  }
  if (ready && !state.data) await loadDashboard("All Domain");  // page reopened while data is loaded
}

async function start(request) {
  const result = await request();
  if (result.error) showMessage("error", result.error);
  else await syncStatus();
}

// ---- the two sources -----------------------------------------------------
$("#fromPurview").onclick = () => start(api.fromPurview);

$("#fromRaw").onclick = async () => {
  const files = await api.rawFiles();
  view.renderRawFiles($("#rawFiles"), files);
  $("#rawFiles").querySelectorAll(".file").forEach((button) => {
    button.onclick = () => { $("#rawDialog").close(); start(() => api.openRaw(files[button.dataset.i].name)); };
  });
  $("#rawDialog").showModal();
};
$("#browse").onclick = () => $("#file").click();
$("#file").onchange = (event) => {
  const [file] = event.target.files;
  event.target.value = "";
  if (!file) return;
  $("#rawDialog").close();
  start(() => api.uploadRaw(file));
};

// ---- toolbar ---------------------------------------------------------------
$("#xlsx").onclick = () => { location.href = api.reportUrl; };
$("#saveRaw").onclick = async () => {
  const result = await api.saveRaw();
  showMessage(result.error ? "error" : "ready", result.error || `Saved ${result.name} in the raw_json folder.`);
};
$("#table").onchange = () => loadDashboard($("#table").value);

// ---- dashboard -----------------------------------------------------------------
async function loadDashboard(table) {
  const data = await api.data(table);
  if (data.error) return;
  state.data = data;
  $("#sources").classList.add("compact");
  $("#dash").hidden = false;
  $("#table").innerHTML = ["All Domain", ...data.tables]
    .map((t) => `<option${t === data.table ? " selected" : ""}>${view.esc(t)}</option>`).join("");
  view.renderHero($("#hero"), data);
  view.renderBands($("#bands"), data);
  view.renderGrid($("#grid"), data);
  view.renderChartData($("#chartdata"), data);
  view.renderRules($("#rules"), data);
  if (!$("#fdim").options.length) view.fillDimensionFilter($("#fdim"));
  drawCoverage();
}

function drawCoverage() {
  if (!state.data) return;
  view.renderCoverage($("#cde"), $("#count"), state.data,
    { query: $("#q").value.toLowerCase(), dimension: $("#fdim").value, yn: $("#fyn").value });
}
["#q", "#fdim", "#fyn"].forEach((id) => { $(id).oninput = drawCoverage; });

syncStatus();
