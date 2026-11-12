/*
 * Report engine for the web page. Runs in the background (a web worker) so the page stays responsive.
 * It runs the very same pmsg_report_app.py as the desktop app, inside the browser (Pyodide),
 * so the Excel report is identical. The CSV files are read on this computer and never sent anywhere.
 */
const PYODIDE = 'https://cdn.jsdelivr.net/pyodide/v0.26.4/full/';
let py = null, booting = null;

/* Starts the engine once; later calls wait for the same start. */
function boot(version) {
  if (!booting) booting = start(version).catch((e) => { booting = null; throw e; });
  return booting;
}
async function start(version) {
  importScripts(PYODIDE + 'pyodide.js');
  const engine = await loadPyodide({ indexURL: PYODIDE });
  const res = await fetch('pmsg_report_app.py?v=' + encodeURIComponent(version || ''), { cache: 'no-cache' });
  if (!res.ok) throw new Error('ENGINE_FILE');
  engine.FS.mkdirTree('/work');
  engine.FS.writeFile('/work/pmsg_report_app.py', await res.text());
  engine.runPython("import sys\nif '/work' not in sys.path: sys.path.insert(0, '/work')\nimport pmsg_report_app");
  py = engine;
  return py;
}

/* Puts one chosen file at /in/<slot>/<its name> without copying it into memory (read on demand). */
function place(slot, file) {
  if (!file) return '';
  const dir = '/in/' + slot;
  try { py.FS.unmount(dir); } catch (e) { /* not mounted yet */ }
  py.FS.mkdirTree(dir);
  try {
    py.FS.mount(py.FS.filesystems.WORKERFS, { files: [file] }, dir);
  } catch (e) {                                  // older browsers: copy the file in instead
    return null;
  }
  return dir + '/' + file.name;
}

self.onmessage = async (ev) => {
  const msg = ev.data;
  if (msg.type === 'warm') {
    try { await boot(msg.version); self.postMessage({ type: 'ready' }); } catch (e) { /* tried again on Create */ }
    return;
  }
  if (msg.type !== 'run') return;
  try {
    self.postMessage({ type: 'stage', stage: 'engine' });
    await boot(msg.version);
    self.postMessage({ type: 'stage', stage: 'read' });
    const paths = {};
    for (const slot of ['sbpdcl', 'nbpdcl', 'targets']) {
      const file = msg.files[slot];
      let p = place(slot, file);
      if (p === null) {                         // fallback copy
        const dir = '/copy/' + slot; py.FS.mkdirTree(dir);
        p = dir + '/' + file.name;
        py.FS.writeFile(p, new Uint8Array(await file.arrayBuffer()));
      }
      paths[slot] = p;
    }
    try { py.FS.mkdirTree('/work/out'); } catch (e) { /* exists */ }
    for (const f of py.FS.readdir('/work/out')) if (f !== '.' && f !== '..') py.FS.unlink('/work/out/' + f);
    py.globals.set('js_log', (text) => self.postMessage({ type: 'log', text: String(text) }));
    py.globals.set('js_progress', (phase, frac, text) => self.postMessage({ type: 'progress', phase: String(phase), frac: Number(frac), text: String(text) }));
    py.globals.set('in_paths', py.toPy(paths));
    const answer = JSON.parse(py.runPython(`
import json, pmsg_report_app as app
def _run():
    p = dict(in_paths)
    try:
        path, facts = app.make_report(p.get('sbpdcl') or '', p.get('nbpdcl') or '', '/work/out', p.get('targets') or None,
                                      progress=lambda ph, fr, tx: js_progress(ph, fr, tx), log=lambda tx: js_log(tx))
    except app.UserError as exc:
        return {"ok": False, "error": str(exc)}
    as_of = facts.get("as_of")
    return {"ok": True, "path": path, "name": path.rsplit("/", 1)[-1], "applications": int(facts.get("applications") or 0),
            "as_of": as_of.strftime("%d %b %Y") if as_of else "", "sheets": list(facts.get("sheets") or []),
            "warnings": list(facts.get("warnings") or []), "ula": int(sum((facts.get("ula_left_out") or {}).values()))}
json.dumps(_run())
`));
    if (!answer.ok) { self.postMessage({ type: 'error', text: answer.error }); return; }
    const bytes = py.FS.readFile(answer.path);
    self.postMessage({ type: 'done', facts: answer, bytes }, [bytes.buffer]);
  } catch (e) {
    const text = String(e && e.message || e);
    let friendly = 'Something went wrong while making the report. ' + text.split('\n').filter(Boolean).slice(-1)[0];
    if (/ENGINE_FILE/.test(text)) friendly = 'The report engine file (pmsg_report_app.py) is missing from this web site.';
    else if (/importScripts|NetworkError|Failed to fetch|load/i.test(text) && !py) friendly = 'Could not get the report engine. Check the internet connection and try again.';
    else if (/memory|Memory/.test(text)) friendly = 'This browser ran out of memory. Close other tabs, or use a desktop browser such as Chrome or Edge.';
    self.postMessage({ type: 'error', text: friendly });
  }
};
