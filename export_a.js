/* export_a.js — ядро сборки книги и первичные листы: Актив, Пассив, Приб_Уб, Удельный вес в ВБ, Изм.стр. Ф1/Ф2.
 * Правило раскладки: период i всегда в колонке 3+i (D, E, …) на всех листах, кроме «Нефин показат». */
(function (root) {
  'use strict';
  const AZ = root.AZ = root.AZ || {}; const M = AZ.model; const XW = root.XW;
  const E = AZ.exp = AZ.exp || {};

  /* Составы итогов (включая подпункты 1210 и 1520) */
  const PARTS = Object.assign({}, M.TOTAL_PARTS, {
    '1210': ['1211', '1212', '1213', '1214', '1215', '1216', '1217'],
    '1520': ['1521', '1522', '1523', '1524', '1525', '1526', '1527', '1528'],
  });
  // 1210 и 1520 входят в итог раздела как целое, подпункты — «в т.ч.»
  const SUBS = new Set([...PARTS['1210'], ...PARTS['1520']]);
  // Форма 2: коды расходов хранятся положительными, знаки задаются формулой
  const PL_FORMULA = {
    '2100': [['2110', 1], ['2120', -1]],
    '2200': [['2100', 1], ['2210', -1], ['2220', -1]],
    '2300': [['2200', 1], ['2310', 1], ['2320', 1], ['2330', -1], ['2340', 1], ['2350', -1]],
    '2400': [['2300', 1], ['2410', -1], ['2460', 1]],
  };
  const PL_ORDER = ['2110', '2120', '2100', '2210', '2220', '2200', '2310', '2320', '2330', '2340', '2350', '2300', '2410', '2421', '2430', '2450', '2460', '2400'];
  const TOTALS = new Set(['1100', '1200', '1600', '1300', '1400', '1500', '1700', '2100', '2200', '2300', '2400']);
  const SEC = { '1110': 'I. ВНЕОБОРОТНЫЕ АКТИВЫ', '1210': 'II. ОБОРОТНЫЕ АКТИВЫ', '1310': 'III. КАПИТАЛ И РЕЗЕРВЫ', '1410': 'IV. ДОЛГОСРОЧНЫЕ ОБЯЗАТЕЛЬСТВА', '1510': 'V. КРАТКОСРОЧНЫЕ ОБЯЗАТЕЛЬСТВА', '2110': 'ФОРМА 2 — ОТЧЕТ О ФИНАНСОВЫХ РЕЗУЛЬТАТАХ (нарастающим итогом с начала года)', '5445': 'СПРАВОЧНО: РАСШИФРОВКИ ДЛЯ КОЭФФИЦИЕНТА ТЕКУЩЕЙ ЛИКВИДНОСТИ (Кл)' };

  const G = 'FF176B4E', GL = 'FFE3F0E9', GREY = 'FFF2F2F2', BLUE = 'FF0000FF', YEL = 'FFFFF2CC', ORG = 'FFFCE4D6', RED = 'FFFFC7CE';
  E.ST = {
    title: { b: 1, sz: 14, color: G }, sub: { i: 1, color: 'FF595959' },
    head: { b: 1, color: 'FFFFFFFF', fill: G, h: 'center', v: 'center', wrap: 1, bd: 1 },
    headL: { b: 1, color: 'FFFFFFFF', fill: G, v: 'center', wrap: 1, bd: 1 },
    sec: { b: 1, fill: GL, color: G },
    code: { color: 'FF7F7F7F', h: 'center' }, name: { wrap: 1 }, nameB: { b: 1, wrap: 1 },
    inp: { color: BLUE, nf: '#,##0;[Red]-#,##0;0' }, calc: { nf: '#,##0;[Red]-#,##0;0' }, tot: { b: 1, nf: '#,##0;[Red]-#,##0;0', fill: GREY, bd: 1 },
    totIn: { b: 1, color: BLUE, nf: '#,##0;[Red]-#,##0;0', fill: GREY, bd: 1 }, totName: { b: 1, fill: GREY, bd: 1 },
    pct: { nf: '0.0%;[Red]-0.0%' }, pctB: { b: 1, nf: '0.0%;[Red]-0.0%', fill: GREY }, num2: { nf: '0.00;[Red]-0.00' },
    miss: { fill: YEL, nf: '#,##0' }, note: { sz: 9, color: 'FF7F7F7F', wrap: 1 }, G, GL, GREY, BLUE, YEL, ORG, RED,
  };
  const ST = E.ST;

  /** Контекст сборки */
  E.makeCtx = function (data) {
    const periods = data.periods; const np = periods.length; const wb = new XW.Workbook({ title: 'Анализ заемщика' });
    const ctx = { wb, data, periods, np, rows: {}, flags: [], srcs: [], sheetOf: {}, pc: i => 3 + i, L: i => XW.colName(3 + i) };
    ctx.idxOfKey = {}; periods.forEach((p, i) => { ctx.idxOfKey[p.key] = i; });
    ctx.startIdx = i => { const k = M.yearStartKey(periods[i].key); return k in ctx.idxOfKey ? ctx.idxOfKey[k] : -1; };
    ctx.prevYearIdx = i => { const k = M.prevYearKey(periods[i].key); return k in ctx.idxOfKey ? ctx.idxOfKey[k] : -1; };
    ctx.prevQIdx = i => { const k = M.prevQuarterKey(periods[i].key); return k in ctx.idxOfKey ? ctx.idxOfKey[k] : -1; };
    ctx.val = (code, i) => { const m = data.vals[code]; const x = m && m[periods[i].key]; return x ? x.value : null; };
    ctx.rec = (code, i) => { const m = data.vals[code]; return m && m[periods[i].key] || null; };
    ctx.sheetName = code => (code[0] === '2' ? 'Приб_Уб' : code[0] === '5' || code === '1700' || (code[0] === '1' && ['3', '4', '5', '7'].includes(code[1])) ? 'Пассив' : 'Актив');
    /** Ссылка на ячейку первичного листа */
    ctx.ref = (code, i) => { const sn = ctx.sheetName(code); const r = ctx.rows[sn] && ctx.rows[sn][code]; if (r === undefined) return null; return XW.quoteSheet(sn) + '!' + XW.colName(3 + i) + (r + 1); };
    ctx.addFlag = (sheet, code, i, msg, level) => ctx.flags.push({ sheet, code, i, msg, level: level || 'warn' });
    return ctx;
  };

  function header(ctx, sh, title, subtitle, firstColLabel) {
    sh.set(0, 0, title, ST.title); sh.set(1, 0, subtitle || '', ST.sub);
    sh.set(3, 0, 'Код', ST.head); sh.set(3, 1, firstColLabel || 'Показатель', ST.headL); sh.set(3, 2, 'Источник / прим.', ST.head);
    ctx.periods.forEach((p, i) => sh.set(3, 3 + i, p.label, ST.head));
    sh.height(3, 30); sh.width(0, 7); sh.width(1, 52); sh.width(2, 14); for (let i = 0; i < ctx.np; i++) sh.width(3 + i, 12.5);
    sh.freeze(4, 3); sh.printTitles = XW.quoteSheet(sh.name) + '!$4:$4';
  }
  E.header = header;
  E.subtitle = ctx => (ctx.data.borrower.name || 'Заемщик не указан') + (ctx.data.borrower.inn ? ' · ИНН ' + ctx.data.borrower.inn : '') + ' · тыс. руб. · ' + (ctx.np ? ctx.periods[0].label + ' — ' + ctx.periods[ctx.np - 1].label : 'нет периодов');

  /** Из чего складывается строка: [[код, знак], …] или null для вводимых строк */
  function comps(code) {
    if (PL_FORMULA[code]) return PL_FORMULA[code];
    const pl = { '1200': ['1210', '1220', '1230', '1240', '1250', '1260'], '1500': ['1510', '1520', '1530', '1540', '1550'], '1600': ['1100', '1200'], '1700': ['1300', '1400', '1500'] };
    const list = pl[code] || (PARTS[code] && PARTS[code].filter(k => !SUBS.has(k) || code === '1210' || code === '1520'));
    return list ? list.map(k => [k, 1]) : null;
  }
  E.comps = comps;
  /** Решение по строке: формула (итог сходится) или константа из отчетности (+ замечание) */
  function resolve(ctx, code, i) {
    const memo = ctx.__memo || (ctx.__memo = {}); const mk = code + '|' + i; if (memo[mk]) return memo[mk];
    const cs = comps(code); const rec = ctx.val(code, i); let out;
    if (!cs) out = { mode: 'in', value: rec, has: rec !== null };
    else {
      let calc = 0, any = false; for (const [k, sg] of cs) { const x = resolve(ctx, k, i); calc += sg * (x.value || 0); if (x.has) any = true; }
      if (rec === null) out = { mode: 'formula', value: calc, has: any };
      else if (!any) out = { mode: 'const', value: rec, has: true, why: 'нет данных по составляющим — итог взят из отчетности' };
      else if (Math.abs(calc - rec) <= Math.max(1, 0.001 * Math.abs(rec))) out = { mode: 'formula', value: calc, has: true };
      else out = { mode: 'const', value: rec, has: true, why: 'итог по отчетности ' + rec + ' ≠ сумма статей ' + Math.round(calc), calc };
    }
    return (memo[mk] = out);
  }
  E.resolve = resolve;

  /** Ячейка первичной таблицы */
  function putPrimary(ctx, sh, r, code, i) {
    const c = 3 + i; const L = XW.colName(c); const tot = TOTALS.has(code); const rowOf = ctx.rows[sh.name];
    const res = resolve(ctx, code, i); const rec = ctx.rec(code, i); const stIn = tot ? ST.totIn : ST.inp, stCalc = tot ? ST.tot : ST.calc;
    if (res.mode === 'formula') {
      const cs = comps(code); const f = cs.map(([k, sg], n) => (sg < 0 ? '-' : n ? '+' : '') + L + (rowOf[k] + 1)).join('');
      sh.set(r, c, { f }, stCalc); return;
    }
    if (res.mode === 'in' && rec === null) { sh.set(r, c, null, stCalc); return; }
    const stl = res.why ? Object.assign({}, stIn, { fill: ST.ORG }) : stIn;
    sh.set(r, c, rec.value, stl);
    ctx.srcs.push({ sheet: sh.name, r, c, code, key: ctx.periods[i].key, value: rec.value, src: rec.src });
    if (res.why) ctx.addFlag(sh.name, code, i, res.why, res.calc !== undefined ? 'warn' : 'info');
  }

  function buildSheetRows(ctx, sh, codes, sideFilter) {
    const rows = {}; ctx.rows[sh.name] = rows; let r = 4;
    for (const code of codes) {
      if (SEC[code]) { r++; sh.set(r, 0, null, ST.sec); sh.set(r, 1, SEC[code], ST.sec); for (let c = 2; c < 3 + ctx.np; c++) sh.set(r, c, null, ST.sec); r++; }
      rows[code] = r; r++;
    }
    return rows;
  }

  E.buildPrimary = function (ctx) {
    const wb = ctx.wb; const sub = E.subtitle(ctx);
    const cn = M.CODE_NAME;
    const A = M.CODES.filter(c => c[2] === 'A').map(c => c[0]);
    const Lc = M.CODES.filter(c => c[2] === 'L').map(c => c[0]);
    const P = PL_ORDER; const Nn = ['5445', '5540', '5320'];
    const sA = wb.addSheet('Актив', { tab: ST.G }), sP = wb.addSheet('Пассив', { tab: ST.G }), sR = wb.addSheet('Приб_Уб', { tab: ST.G });
    // ряды: «Пассив» содержит обязательства, форму 2 вынесем на свой лист; справочные коды — внизу пассива
    buildSheetRows(ctx, sA, A); buildSheetRows(ctx, sP, Lc.concat(Nn)); buildSheetRows(ctx, sR, P);
    const draw = (sh, codes, title) => {
      header(ctx, sh, title, sub);
      for (const code of codes) {
        const r = ctx.rows[sh.name][code]; const tot = TOTALS.has(code);
        if (SEC[code]) { const rr = r - 1; sh.set(rr, 0, null, ST.sec); sh.set(rr, 1, SEC[code], ST.sec); for (let c = 2; c < 3 + ctx.np; c++) sh.set(rr, c, null, ST.sec); }
        sh.set(r, 0, code, tot ? ST.totName : ST.code); sh.set(r, 1, (SUBS.has(code) ? '   в т.ч. ' : '') + cn[code], tot ? ST.totName : ST.name); sh.set(r, 2, null, tot ? ST.totName : ST.name);
        for (let i = 0; i < ctx.np; i++) putPrimary(ctx, sh, r, code, i);
      }
    };
    draw(sA, A, 'Форма 1 — Бухгалтерский баланс: АКТИВ'); draw(sP, Lc.concat(Nn), 'Форма 1 — Бухгалтерский баланс: ПАССИВ');
    draw(sR, P, 'Форма 2 — Отчет о финансовых результатах');
    sR.set(2, 0, 'Знаки: расходы (2120, 2210, 2220, 2330, 2350, 2410) вносятся положительными числами, итоги считаются формулами.', ST.note);
    sA.set(2, 0, 'Синий шрифт — данные из документов. Оранжевая заливка — значение взято с оговоркой (см. «Проверки»). Итоги — формулы; расхождение с отчетным итогом выводится на листе «Проверки».', ST.note);
    // квартальный блок Приб_Уб (разностные значения из нарастающих)
    const q0 = ctx.rows['Приб_Уб']['2400'] + 4; ctx.qBlock = { r0: q0, rows: {} };
    sR.set(q0 - 1, 0, null, ST.sec); sR.set(q0 - 1, 1, 'ПО КВАРТАЛАМ (разность нарастающих итогов; I кв. = нарастающий итог)', ST.sec); for (let c = 2; c < 3 + ctx.np; c++) sR.set(q0 - 1, c, null, ST.sec);
    P.forEach((code, n) => {
      const r = q0 + n; ctx.qBlock.rows[code] = r; const tot = TOTALS.has(code);
      sR.set(r, 0, code, tot ? ST.totName : ST.code); sR.set(r, 1, cn[code], tot ? ST.totName : ST.name); sR.set(r, 2, null, tot ? ST.totName : ST.name);
      for (let i = 0; i < ctx.np; i++) {
        const cur = XW.colName(3 + i) + (ctx.rows['Приб_Уб'][code] + 1); const p = ctx.periods[i];
        if (p.q === 1) sR.set(r, 3 + i, { f: cur }, tot ? ST.tot : ST.calc);
        else { const j = ctx.prevQIdx(i); if (j >= 0) sR.set(r, 3 + i, { f: cur + '-' + XW.colName(3 + j) + (ctx.rows['Приб_Уб'][code] + 1) }, tot ? ST.tot : ST.calc); else sR.set(r, 3 + i, null, tot ? ST.tot : ST.calc); }
      }
    });
    return { sA, sP, sR };
  };

  /** Удельный вес в валюте баланса и изменения г/г */
  E.buildShares = function (ctx) {
    const wb = ctx.wb; const cn = M.CODE_NAME; const sub = E.subtitle(ctx);
    const sw = wb.addSheet('Удельный вес в ВБ', { tab: 'FF5B9BD5' });
    header(ctx, sw, 'Удельный вес статей в валюте баланса', sub); let r = 4;
    const list = ['A', 'L'].map(s => M.CODES.filter(c => c[2] === s).map(c => c[0]));
    ctx.rows['Вес'] = {};
    list.forEach((codes, k) => {
      const base = k === 0 ? '1600' : '1700';
      sw.set(r, 0, null, ST.sec); sw.set(r, 1, k === 0 ? 'АКТИВ (к 1600)' : 'ПАССИВ (к 1700)', ST.sec); for (let c = 2; c < 3 + ctx.np; c++) sw.set(r, c, null, ST.sec); r++;
      for (const code of codes) {
        const tot = TOTALS.has(code); ctx.rows['Вес'][code] = r;
        sw.set(r, 0, code, tot ? ST.totName : ST.code); sw.set(r, 1, (SUBS.has(code) ? '   в т.ч. ' : '') + cn[code], tot ? ST.totName : ST.name); sw.set(r, 2, null, tot ? ST.totName : ST.name);
        for (let i = 0; i < ctx.np; i++) sw.set(r, 3 + i, { f: 'IFERROR(' + ctx.ref(code, i) + '/' + ctx.ref(base, i) + ',"")' }, tot ? ST.pctB : ST.pct);
        r++;
      }
      r++;
    });
    const mk = (name, codes, title, pairs) => {
      const sh = wb.addSheet(name, { tab: 'FF5B9BD5' }); header(ctx, sh, title, sub + ' · сравнение с тем же периодом прошлого года'); let r2 = 4;
      for (const blk of [{ t: 'ИЗМЕНЕНИЕ, тыс. руб.', pct: false }, { t: 'ИЗМЕНЕНИЕ, %', pct: true }]) {
        sh.set(r2, 0, null, ST.sec); sh.set(r2, 1, blk.t, ST.sec); for (let c = 2; c < 3 + ctx.np; c++) sh.set(r2, c, null, ST.sec); r2++;
        for (const code of codes) {
          const tot = TOTALS.has(code);
          sh.set(r2, 0, code, tot ? ST.totName : ST.code); sh.set(r2, 1, (SUBS.has(code) ? '   в т.ч. ' : '') + cn[code], tot ? ST.totName : ST.name); sh.set(r2, 2, null, tot ? ST.totName : ST.name);
          for (let i = 0; i < ctx.np; i++) {
            const j = ctx.prevYearIdx(i); const a = ctx.ref(code, i), b = j >= 0 ? ctx.ref(code, j) : null;
            if (!b) { sh.set(r2, 3 + i, null, tot ? ST.tot : ST.calc); continue; }
            if (blk.pct) sh.set(r2, 3 + i, { f: 'IFERROR((' + a + '-' + b + ')/ABS(' + b + '),"")' }, tot ? ST.pctB : ST.pct);
            else sh.set(r2, 3 + i, { f: a + '-' + b }, tot ? ST.tot : ST.calc);
          }
          r2++;
        }
        r2++;
      }
    };
    mk('Изм.стр. Ф1', M.CODES.filter(c => c[2] === 'A' || c[2] === 'L').map(c => c[0]), 'Изменение статей баланса (Ф1), г/г');
    mk('Изм.стр. Ф2', PL_ORDER, 'Изменение статей отчета о финрезультатах (Ф2), г/г');
  };
  E.PL_ORDER = PL_ORDER; E.PARTS = PARTS; E.SUBS = SUBS; E.TOTALS = TOTALS; E.PL_FORMULA = PL_FORMULA;
})(typeof window !== 'undefined' ? window : globalThis);
