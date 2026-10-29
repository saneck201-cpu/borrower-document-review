/* core.js — слияние записей из разных документов в единую отчетность и автоподстановка реквизитов (без DOM, тестируется в Node). */
(function (root) {
  'use strict';
  const AZ = root.AZ = root.AZ || {}; const M = AZ.model;
  const tol = v => Math.max(1, 0.001 * Math.abs(v));

  /** records → {vals:{code:{key:{value,src}}}, periods:[…], notes:[…], conflicts:n} */
  function mergeRecords(recs, opts) {
    opts = opts || {}; const notes = []; const best = new Map(); let conflicts = 0; let unitsUnknown = 0;
    const list = recs.map(r => (r.key === 'latest' && opts.fallbackKey ? Object.assign({}, r, { key: opts.fallbackKey, role: 'cur' }) : r));
    for (const r of list) {
      if (r.key === 'latest') continue; if (r.unitsKnown === false) unitsUnknown++;
      const k = r.code + '|' + r.key; const old = best.get(k);
      if (!old) { best.set(k, r); continue; }
      if (Math.abs(old.value - r.value) > tol(old.value)) {
        conflicts++;
        // «текущий» столбец таблицы надежнее сравнительного; при равенстве остается первое найденное (первичные листы идут раньше)
        if (old.role !== 'cur' && r.role === 'cur') { notes.push({ where: 'Слияние', period: M.periodFromKey(r.key).label, level: 'warn', text: `Код ${r.code}: ${old.value} (${old.src}) заменено на ${r.value} (${r.src}) — «текущий» столбец надежнее сравнительного.` }); best.set(k, r); }
        else notes.push({ where: 'Слияние', period: M.periodFromKey(r.key).label, level: 'warn', text: `Код ${r.code}: расхождение ${old.value} (${old.src}) и ${r.value} (${r.src}) — принято первое.` });
      }
    }
    if (list.some(r => r.key === 'latest') && !opts.fallbackKey) notes.push({ where: 'Импорт', level: 'warn', text: 'Найдена отчетность без распознанной даты — выберите отчетный период в карточке заемщика, чтобы она попала в расчет.' });
    if (unitsUnknown) notes.push({ where: 'Импорт', level: 'warn', text: `Единицы измерения не определены у ${unitsUnknown} значений — приняты тыс. руб. Проверьте масштаб (рубли/тыс./млн).` });
    const vals = {}; const nz = new Set();
    for (const r of best.values()) { (vals[r.code] = vals[r.code] || {})[r.key] = { value: r.value, src: r.src }; if (r.value) nz.add(r.key); }
    let keys = [...nz].sort(M.keyCmp);
    if (keys.length > 41) { notes.push({ where: 'Импорт', level: 'info', text: `Периодов ${keys.length}; в книгу взяты последние 41.` }); keys = keys.slice(-41); }
    const keep = new Set(keys); for (const c in vals) for (const k of Object.keys(vals[c])) if (!keep.has(k)) delete vals[c][k];
    return { vals, periods: keys.map(M.periodFromKey), notes, conflicts };
  }

  /** Реквизиты из текстов документов: берем самое частое значение, приоритет у регистрационных */
  function mergeEntities(items) {
    const out = {}; const pri = x => (/Регистрац/.test(x.category || '') ? 0 : 1);
    const sorted = items.slice().sort((a, b) => pri(a) - pri(b));
    for (const it of sorted) { const e = it.entities || {}; for (const k of ['inn', 'ogrn', 'regDate', 'okved', 'name']) if (!out[k] && e[k]) out[k] = e[k]; }
    return out;
  }

  /** Автоподстановка по реквизитам: отрасль, длительность работы (фактор 6) */
  function autoFill(ent, now) {
    const out = { industry: null, factors: {}, notes: [] };
    if (ent.okved) { out.industry = M.industryByOkved(ent.okved); out.notes.push({ where: 'Автоподстановка', level: 'info', text: `Отрасль «${out.industry}» определена по ОКВЭД ${ent.okved}.` }); }
    if (ent.regDate) {
      const d = new Date(ent.regDate.y, ent.regDate.mo - 1, ent.regDate.d); now = now || new Date();
      const months = (now.getFullYear() - d.getFullYear()) * 12 + now.getMonth() - d.getMonth();
      out.factors[5] = { score: months > 12 ? 3 : months >= 6 ? 2 : 1, auto: true, note: `Авто: дата регистрации ${ent.regDate.text}, работает ~${months} мес.` };
    }
    return out;
  }

  AZ.core = { mergeRecords, mergeEntities, autoFill };
})(typeof window !== 'undefined' ? window : globalThis);
