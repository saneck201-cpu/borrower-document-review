/* merge.js — слияние записей из разных файлов в единый набор данных по периодам. Чистая логика, без DOM. */
(function (root) {
  'use strict';
  const AZ = root.AZ = root.AZ || {};
  const M = AZ.model;
  const tol = v => Math.max(1, Math.abs(v) * 0.001);
  const NEG_CODES = ['2120', '2210', '2220', '2330', '2350', '2410']; // расходы в форме 2 хранятся со знаком «минус»

  /**
   * records: [{code,key,value,src,role,unitsKnown,label}]
   * opts: {defaultKey} — куда отнести записи «последний период (дата не распознана)»
   * Возвращает {periods, data, src, conflicts, notes}
   */
  function merge(records, opts) {
    opts = opts || {};
    const rank = r => (r.role === 'cur' ? 2 : 1);
    const best = new Map(); const conflicts = []; const notes = []; const unitsUnknown = new Set();
    const realKeys = new Set(records.filter(r => /^\d{4}-[1-4]$/.test(r.key)).map(r => r.key));
    const latestKey = opts.defaultKey || (realKeys.size ? [...realKeys].sort(M.keyCmp).pop() : null);
    for (const r0 of records) {
      let r = r0; if (r.key === 'latest') { if (!latestKey) { notes.push('Не определён период для «' + r.code + '» (' + r.src + ')'); continue; } r = Object.assign({}, r, { key: latestKey, role: 'cmp' }); }
      if (!/^\d{4}-[1-4]$/.test(r.key)) continue;
      if (r.unitsKnown === false) unitsUnknown.add(String(r.src).split(' · ')[0]);
      const k = r.key + '|' + r.code; const cur = best.get(k);
      if (!cur) { best.set(k, r); continue; }
      if (Math.abs(cur.value - r.value) > tol(cur.value)) conflicts.push({ key: r.key, code: r.code, a: cur.value, srcA: cur.src, b: r.value, srcB: r.src, kept: rank(r) > rank(cur) ? 'b' : 'a' });
      if (rank(r) > rank(cur)) best.set(k, r);
    }
    const data = {}; const src = {};
    for (const [k, r] of best) { const [key, code] = k.split('|'); (data[key] = data[key] || {})[code] = r.value; (src[key] = src[key] || {})[code] = r.src; }
    // пустые колонки шаблонов: период без единого ненулевого значения отбрасываем
    for (const key of Object.keys(data)) { if (!Object.values(data[key]).some(v => v !== 0 && Number.isFinite(v))) { delete data[key]; delete src[key]; } }
    // знаки расходов формы 2
    for (const key of Object.keys(data)) for (const c of NEG_CODES) if (data[key][c] != null && data[key][c] > 0) data[key][c] = -data[key][c];
    let periods = Object.keys(data).sort(M.keyCmp);
    if (periods.length > 41) { notes.push('Периодов больше 41: взяты последние 41'); periods = periods.slice(-41); }
    for (const k of Object.keys(data)) if (!periods.includes(k)) { delete data[k]; delete src[k]; }
    if (unitsUnknown.size) notes.push('Единицы измерения не определены (значения взяты как есть, ожидаются тыс. руб.): ' + [...unitsUnknown].join('; '));
    return { periods, data, src, conflicts, notes };
  }
  AZ.merge = { merge, NEG_CODES, tol };
})(typeof window !== 'undefined' ? window : globalThis);
