// Runs the FUEL parser (Python, via Pyodide/WebAssembly) inside this browser tab.
// Files are passed in memory from the page; nothing is sent to any server.
const VERSION = "0.26.4";
const CDN = `https://cdn.jsdelivr.net/pyodide/v${VERSION}/full/`;
let py = null;
let bridge = null;

function loadPyodideScript() {
  // Local development/testing may self-host Pyodide under /pyodide/; production uses the CDN.
  const local = ["localhost", "127.0.0.1"].includes(self.location.hostname);
  if (local) {
    try { importScripts("/pyodide/pyodide.js"); return "/pyodide/"; } catch (e) { /* fall back to CDN */ }
  }
  importScripts(CDN + "pyodide.js");
  return CDN;
}

function syncfs(populate) {
  return new Promise((resolve, reject) => py.FS.syncfs(populate, (err) => (err ? reject(err) : resolve())));
}

async function init() {
  const base = loadPyodideScript();
  py = await loadPyodide({ indexURL: base });
  // openpyxl is not in the Pyodide distribution: its pure-Python wheels are served from this site.
  const wheel = (name) => new URL("wheels/" + name, self.location).href;
  await py.loadPackage(["pyyaml", "sqlite3", wheel("et_xmlfile-2.0.0-py3-none-any.whl"), wheel("openpyxl-3.1.5-py2.py3-none-any.whl")]);
  const bundle = await (await fetch("fuel_bundle.json", { cache: "no-cache" })).json();
  for (const [path, text] of Object.entries(bundle.files)) {
    const full = "/app/" + path;
    py.FS.mkdirTree(full.slice(0, full.lastIndexOf("/")));
    py.FS.writeFile(full, text);
  }
  py.FS.mkdirTree("/ledger");
  py.FS.mount(py.FS.filesystems.IDBFS, {}, "/ledger"); // import history persists in this browser only
  await syncfs(true);
  py.FS.mkdirTree("/work/in");
  py.runPython("import sys; sys.path.insert(0, '/app')");
  bridge = py.pyimport("bridge");
  return { version: bundle.version };
}

function writeInputs(files) {
  const paths = {};
  for (const [role, f] of Object.entries(files || {})) {
    const dir = `/work/in/${role}`;
    py.FS.mkdirTree(dir);
    for (const name of py.FS.readdir(dir)) if (name !== "." && name !== "..") py.FS.unlink(`${dir}/${name}`);
    if (!f) continue;
    const safe = f.name.replace(/[^\w.\- ()]/g, "_");
    const path = `${dir}/${safe}`; // one folder per role keeps the original file name for output names
    py.FS.writeFile(path, new Uint8Array(f.buffer));
    paths[role] = path;
  }
  return paths;
}

const ops = {
  async init() { return await init(); },
  async classify({ file }) {
    const path = "/work/classify__" + file.name.replace(/[^\w.\- ()]/g, "_");
    py.FS.writeFile(path, new Uint8Array(file.buffer));
    try { return bridge.classify(path); } finally { py.FS.unlink(path); }
  },
  async run({ params, files }) {
    const paths = writeInputs(files);
    const res = JSON.parse(bridge.run(JSON.stringify({ ...params, source: paths.source, template: paths.template, sellers: paths.sellers || null })));
    const outputs = {};
    const transfer = [];
    if (res.status === "ok") {
      for (const [kind, path] of Object.entries(res.outputs)) {
        const bytes = py.FS.readFile(path);
        outputs[kind] = { name: path.split("/").pop(), bytes };
        transfer.push(bytes.buffer);
      }
      if (!params || params.mode === "export") await syncfs(false);
    }
    return { result: res, outputs, transfer };
  },
  async batches() { return JSON.parse(bridge.batches()); },
  async voidBatch({ batchId }) { const r = JSON.parse(bridge.void_batch(batchId)); await syncfs(false); return r; },
};

self.onmessage = async (ev) => {
  const { id, op, args } = ev.data;
  try {
    const data = await ops[op](args || {});
    const transfer = (data && data.transfer) || [];
    if (data && data.transfer) delete data.transfer;
    self.postMessage({ id, ok: true, data }, transfer);
  } catch (err) {
    self.postMessage({ id, ok: false, error: String((err && err.message) || err) });
  }
};
