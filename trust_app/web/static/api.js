// Thin wrapper over the server's JSON API.
async function call(url, options) {
  const response = await fetch(url, options);
  return response.json();
}

const post = (url, body) => call(url, {
  method: "POST",
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify(body ?? {}),
});

export const api = {
  status: () => call("/api/status"),
  data: (table) => call("/api/data?table=" + encodeURIComponent(table)),
  fromPurview: () => post("/api/purview"),
  rawFiles: () => call("/api/raw-files"),
  openRaw: (name) => post("/api/raw-open", { name }),
  uploadRaw: (file) => {
    const body = new FormData();
    body.append("file", file);
    return call("/api/raw-upload", { method: "POST", body });
  },
  saveRaw: () => post("/api/save-raw"),
  reportUrl: "/api/report",
};
