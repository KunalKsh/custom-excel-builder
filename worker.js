// Runs Python (Pyodide) off the main thread so the page stays responsive on big files.
importScripts('https://cdn.jsdelivr.net/pyodide/v0.26.4/full/pyodide.js');

let py;
const ready = (async () => {
  postMessage({ progress: 'Starting the Excel engine (first load takes about 20 seconds)...' });
  py = await loadPyodide();
  await py.loadPackage('micropip');
  await py.pyimport('micropip').install(['pyxlsb', 'xlsxwriter', 'openpyxl', 'xlrd']);
  py.runPython(await (await fetch('engine.py', { cache: 'no-cache' })).text());
  postMessage({ progress: '', ready: true });
})();
ready.catch(e => postMessage({ fatal: 'Engine failed to start: ' + e }));

const call = (name, ...args) => py.globals.get(name)(...args);

onmessage = async ev => {
  const { id, type, payload } = ev.data;
  try {
    await ready;
    let result;
    if (type === 'add') {
      const dir = '/mnt/' + payload.fid;
      py.FS.mkdirTree(dir);
      py.FS.mount(py.FS.filesystems.WORKERFS, { files: [payload.file] }, dir); // read in place, no copy
      result = JSON.parse(call('add_file', payload.fid, payload.name, dir + '/' + payload.file.name, payload.hrow));
    } else if (type === 'headers') {
      result = JSON.parse(call('headers', JSON.stringify(payload.ids), payload.hrow));
    } else if (type === 'remove') {
      call('remove_file', payload.id);
      try { py.FS.unmount('/mnt/' + payload.id); } catch (e) { /* already gone */ }
    } else if (type === 'export') {
      const r = JSON.parse(call('export', JSON.stringify(payload), (m, p) => postMessage({ progress: m, pct: p == null ? null : p })));
      const data = py.FS.readFile(r.path);
      py.FS.unlink(r.path);
      postMessage({ id, result: { buffer: data.buffer, name: r.name, truncated: r.truncated } }, [data.buffer]);
      return;
    }
    postMessage({ id, result });
  } catch (e) {
    postMessage({ id, error: String(e.message || e).split('\n').filter(Boolean).pop() });
  }
};
