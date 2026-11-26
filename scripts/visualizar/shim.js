/* Adaptador: reproduz a API "dash" dos painéis do Claude com dados embutidos no arquivo. */
(function () {
  const raw = window.__DATASETS || {};
  const datasets = {};
  Object.keys(raw).forEach((id) => { datasets[id] = d3.csvParse(raw[id].trim()); });
  const calcs = {}, cache = {}, draws = {}, onDatas = [];
  function get(id) {
    if (id in datasets) return datasets[id];
    if (id in cache) return cache[id];
    const c = calcs[id];
    if (!c) return null;
    cache[id] = c.fn(...c.inputs.map(get));
    return cache[id];
  }
  window.dash = {
    colors: ['var(--c0)', 'var(--c1)', 'var(--c2)', 'var(--c3)', 'var(--c4)', 'var(--c5)'],
    data(id) {
      const d = get(id);
      if (d === null || d === undefined) return { status: 'missing', data: [], columns: [] };
      return { status: 'ok', data: d, columns: d.columns || [], refreshing: false, truncated: false };
    },
    calc(id, spec) { calcs[id] = spec; },
    loader() {},
    onData(fn) { onDatas.push(fn); },
    draw(name, fn) {
      if (fn) { draws[name] = fn; return; }
      run(name);
    },
    params() { return {}; }, setParams() {}, resetParams() {}, setLink() {}, refresh() {},
  };
  function run(name) {
    const fn = draws[name];
    if (!fn) return;
    document.querySelectorAll('[data-draw="' + name + '"]').forEach((el) => {
      try {
        if (typeof fn === 'function') fn(el);
        else { const made = fn.mount ? fn.mount(el) : undefined; if (fn.update) fn.update(el, made); }
      } catch (e) { console.error('draw ' + name, e); }
    });
  }
  function runAll() { onDatas.forEach((f) => { try { f(); } catch (e) { console.error(e); } }); Object.keys(draws).forEach(run); }
  window.__dashStart = function () {
    runAll();
    let t = null, w = window.innerWidth;
    window.addEventListener('resize', () => {
      if (window.innerWidth === w) return; w = window.innerWidth;
      clearTimeout(t); t = setTimeout(runAll, 150);
    });
  };
})();
