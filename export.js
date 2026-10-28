/* export.js — сборка книги Excel по шаблону «фин анализ»: все расчеты — формулами, периоды выровнены по колонкам D, E, …
 * Часть 1: контекст, балансовые листы, ОПУ, структура, динамика, чистые активы, коэффициенты, справочник. */
(function (root) {
  'use strict';
  const AZ = root.AZ = root.AZ || {};
  const M = AZ.model, XW = root.XW;
  const FC = 3, HR = 4, R0 = 5;               // первая колонка периодов (D), строка заголовка периодов, первая строка данных
  const SH = { SUM: 'Сводка', A: 'Актив', L: 'Пассив', P: 'Приб_Уб', W: 'Удельный вес в ВБ', D1: 'Изм.стр. Ф1', D2: 'Изм.стр. Ф2', NA: 'Чистые активы', K: 'Коэффициенты', S: 'Фин показ', N: 'Нефин показат', F: 'Фин состояние', G: 'графики', R: 'СПРАВОЧНИК', C: 'Проверки', SRC: 'Источники', DOC: 'Документы' };
  const NF = { num: '#,##0;[Red]-#,##0;0', pct: '0.0%;[Red]-0.0%;0.0%', k2: '0.00;[Red]-0.00;0.00', k1: '0.0;[Red]-0.0;0.0' };
  const C = { head: 'FF173B31', headTx: 'FFFFFFFF', sect: 'FFDDEBE3', input: 'FF0000FF', yellow: 'FFFFF2B3', orange: 'FFFFD9A8', grey: 'FFF1F3F2', red: 'FFFAD4D4', green: 'FFD6EFDC', link: 'FF0563C1' };
  const ST = {
    title: { b: 1, sz: 14, color: 'FF173B31' }, sub: { i: 1, color: 'FF5B6B64' },
    hdr: { b: 1, color: C.headTx, fill: C.head, h: 'center', wrap: 1, bd: { all: 'thin' } },
    hdrL: { b: 1, color: C.headTx, fill: C.head, wrap: 1, bd: { all: 'thin' } },
    lab: { bd: { all: 'thin' }, wrap: 1 }, labB: { b: 1, bd: { all: 'thin' }, fill: C.sect, wrap: 1 }, code: { h: 'center', color: 'FF5B6B64', bd: { all: 'thin' } },
    inp: (nf, fill) => ({ color: C.input, nf: nf || NF.num, bd: { all: 'thin' }, fill }), calc: (nf, b, fill) => ({ nf: nf || NF.num, bd: { all: 'thin' }, b, fill }),
    note: { i: 1, color: 'FF5B6B64', wrap: 1 }
  };
  const tolOf = v => Math.max(1, Math.abs(v) * 0.001);
  const SUMS = {
    A: { 1100: ['1110', '1120', '1130', '1140', '1150', '1160', '1170', '1180', '1190'], 1200: ['1210', '1220', '1230', '1240', '1250', '1260'], 1600: ['1100', '1200'] },
    L: { 1300: ['1310', '1320', '1340', '1350', '1360', '1370'], 1400: ['1410', '1420', '1430', '1450'], 1500: ['1510', '1520', '1530', '1540', '1550'], 1700: ['1300', '1400', '1500'] },
    P: { 2100: ['2110', '2120'], 2200: ['2100', '2210', '2220'], 2300: ['2200', '2310', '2320', '2330', '2340', '2350'], 2400: [['2300', '2410'], ['2300', '2410', '2460'], ['2300', '2410', '2430', '2450', '2460']] }
  };
  const KIND_SHEET = { A: SH.A, L: SH.L, P: SH.P };
  const REQUIRED = { A: ['1100', '1200', '1600'], L: ['1300', '1400', '1500', '1700'], P: ['2110', '2200', '2300'] };
  const INDUSTRY_TBL = [['Торговля', 5.8, 4.4, 5.8], ['Строительство', 9.3, 3.4, 9.3], ['Промышленность', 13.4, 7.2, 13.4], ['Прочее', 10.9, 5.0, 12.7]];
  const qIdx = key => { const m = /^(\d{4})-([1-4])$/.exec(key); return +m[1] * 4 + +m[2]; };

  function newCtx(S) {
    const X = { S, periods: S.periods.slice(), wb: new XW.Workbook({ title: 'Анализ заемщика' }), sheets: {}, rows: { A: {}, L: {}, P: {}, Q: {} }, checks: [], FC, HR, R0, SH, ST, NF, C, K: {}, flags: [] };
    X.n = X.periods.length;
    X.col = i => FC + i; X.L = i => XW.colName(FC + i);
    X.idx = key => X.periods.indexOf(key);
    X.ref = (sheet, r, i, absR, absC) => XW.quoteSheet(sheet) + '!' + XW.addr(r, FC + i, absR, absC);
    X.cell = (sh, r, i) => XW.addr(r, FC + i);
    X.lab = key => M.periodFromKey(key).label;
    X.bal = (kind, code, i) => X.ref(KIND_SHEET[kind], X.rows[kind][code], i);
    X.v = (key, code) => (S.data[key] && S.data[key][code] != null) ? S.data[key][code] : null;
    X.prevYear = i => X.idx(M.prevYearKey(X.periods[i]));
    X.prevQ = i => X.idx(M.prevQuarterKey(X.periods[i]));
    X.startIdx = i => X.idx(M.yearStartKey(X.periods[i]));
    X.sheet = (name, opts) => (X.sheets[name] = X.wb.addSheet(name, opts));
    return X;
  }

  /* ── заголовок листа: название, заемщик, подписи периодов ── */
  function sheetHead(X, sh, title, opts) {
    opts = opts || {};
    sh.set(0, 0, title, ST.title); sh.set(1, 0, 'Заемщик: ' + (X.S.borrower.name || '— не определен —') + (X.S.borrower.inn ? ' · ИНН ' + X.S.borrower.inn : ''), ST.sub);
    sh.set(HR, 0, opts.c0 || 'Показатель', ST.hdrL); sh.set(HR, 1, opts.c1 || 'Код', ST.hdr); sh.set(HR, 2, opts.c2 || '', ST.hdr);
    X.periods.forEach((k, i) => { sh.set(HR, FC + i, X.lab(k), ST.hdr); sh.width(FC + i, opts.pw || 13); });
    sh.width(0, opts.w0 || 46); sh.width(1, 8); sh.width(2, opts.w2 || 10); sh.height(HR, 30); sh.freeze(HR + 1, FC); sh.landscape = true;
    sh.link(0, 0, XW.quoteSheet(SH.SUM) + '!A1', 'К сводке'); sh.cf;
  }

  /* ── Актив / Пассив / Приб_Уб ── */
  function inputSheet(X, kind, title) {
    const sh = X.sheets[KIND_SHEET[kind]]; sheetHead(X, sh, title, { c2: 'Ед.' });
    const list = M.CODES.filter(c => c[2] === kind); const rows = X.rows[kind]; let r = R0;
    for (const [code, name] of list) rows[code] = r++;
    X.lastRow = X.lastRow || {}; X.lastRow[kind] = r - 1;
    if (kind === 'A') { r++; sh.set(r, 0, 'Расшифровки к отчетности (для коэффициента текущей ликвидности)', ST.labB); sh.merge(r, 0, r, 2); r++; for (const [code, name] of M.CODES.filter(c => c[2] === 'N')) { rows[code] = r++; } }
    const sums = SUMS[kind] || {};
    const eff = {};                                              // итоговое значение ячейки: eff[i][code]
    X.periods.forEach((key, i) => { eff[i] = {}; });
    const resolve = (i, code) => {                                // что окажется в ячейке: {f|v, mode}
      const key = X.periods[i]; if (eff[i][code] !== undefined) return eff[i][code];
      const rep = X.v(key, code); let res;
      const alts = sums[code] ? (Array.isArray(sums[code][0]) ? sums[code] : [sums[code]]) : null;
      if (alts) {
        const calc = parts => { const vals = parts.map(p => resolve(i, p)).filter(x => x.val != null); return { parts, vals, sum: vals.reduce((a, x) => a + x.val, 0) }; };
        const cands = alts.map(calc); const first = cands.find(c => c.vals.length);
        if (!first) res = { val: rep, mode: rep == null ? 'empty' : 'const' };
        else if (rep == null) res = { val: first.sum, mode: 'formula', parts: first.parts };
        else { const hit = cands.find(c => c.vals.length && Math.abs(c.sum - rep) <= tolOf(rep)); if (hit) res = { val: hit.sum, mode: 'formula', parts: hit.parts }; else { res = { val: rep, mode: 'mismatch', sum: first.sum }; X.checks.push({ sheet: SH[kind], code, key, rep, sum: first.sum }); } }
      } else res = { val: rep, mode: rep == null ? 'empty' : 'const' };
      return (eff[i][code] = res);
    };
    X.eff = X.eff || {}; X.eff[kind] = eff; X.resolve = X.resolve || {}; X.resolve[kind] = resolve;
    for (const [code, name] of M.CODES.filter(c => c[2] === kind || (kind === 'A' && c[2] === 'N'))) {
      const row = rows[code]; const isTot = !!sums[code] || code === '1600' || code === '1700'; const isSub = /^(121[1-7]|152[1-8]|2421)$/.test(code);
      sh.set(row, 0, (isSub ? '   ' : '') + name, isTot ? ST.labB : ST.lab); sh.set(row, 1, code, ST.code); sh.set(row, 2, 'тыс. руб.', ST.code);
      X.periods.forEach((key, i) => {
        const res = resolve(i, code); const c = FC + i; const nf = NF.num;
        if (res.mode === 'formula') { const parts = res.parts.filter(p => resolve(i, p).val != null); sh.set(row, c, { f: parts.map(p => XW.addr(rows[p], c)).join('+') }, ST.calc(nf, isTot, isTot ? C.sect : undefined)); }
        else if (res.mode === 'mismatch') sh.set(row, c, res.val, ST.inp(nf, C.orange));
        else if (res.mode === 'const') sh.set(row, c, res.val, ST.inp(nf, isTot ? C.sect : undefined));
        else if (code === '5445' || code === '5540' || code === '5320') { sh.set(row, c, 0, ST.inp(nf, C.orange)); X.flags.push({ sheet: SH.A, code, key, text: 'Расшифровка ' + code + ' (' + M.CODE_NAME[code] + ') не найдена в документах: принято 0' }); }
        else if ((REQUIRED[kind] || []).includes(code)) { sh.set(row, c, null, { fill: C.yellow, bd: { all: 'thin' } }); X.flags.push({ sheet: KIND_SHEET[kind], code, key, text: 'Нет данных по итоговой строке ' + code + ' — заполните вручную (желтая ячейка)' }); }
        else sh.set(row, c, null, { bd: { all: 'thin' } });
      });
    }
    // выделение строк с суммой, не сходящейся с итогом отчетности, — пояснение в конце листа
    const endR = Math.max(...Object.values(rows)) + 2;
    sh.set(endR, 0, 'Синий шрифт — данные из документов; черный — формулы; оранжевая заливка — расхождение с итогом отчетности или допущение; желтая — нужен ручной ввод.', ST.note); sh.merge(endR, 0, endR, 8); sh.height(endR, 30);
    if (kind === 'P') sh.set(endR + 1, 0, 'Форма 2 заполнена нарастающим итогом с начала года (3, 6, 9, 12 мес.). Расходы — со знаком «минус».', ST.note), sh.merge(endR + 1, 0, endR + 1, 8);
  }

  function quarterBlock(X) {
    const sh = X.sheets[SH.P]; const rows = X.rows.P; const codes = ['2110', '2120', '2100', '2210', '2220', '2200', '2300', '2400'];
    let r = X.lastRow.P + 5; sh.set(r, 0, 'Отчет по кварталам (прирост за квартал = нарастающий итог минус предыдущий нарастающий итог)', ST.labB); sh.merge(r, 0, r, 2);
    X.periods.forEach((k, i) => sh.set(r, FC + i, X.lab(k), ST.hdr)); r++;
    for (const code of codes) {
      X.rows.Q[code] = r; sh.set(r, 0, M.CODE_NAME[code], ST.lab); sh.set(r, 1, code, ST.code); sh.set(r, 2, 'тыс. руб.', ST.code);
      X.periods.forEach((key, i) => {
        const c = FC + i; const q = +key.split('-')[1];
        if (q === 1) sh.set(r, c, { f: XW.addr(rows[code], c) }, ST.calc());
        else { const j = X.prevQ(i); if (j >= 0) sh.set(r, c, { f: XW.addr(rows[code], c) + '-' + XW.addr(rows[code], FC + j) }, ST.calc()); else sh.set(r, c, null, { fill: C.grey, bd: { all: 'thin' } }); }
      }); r++;
    }
    sh.set(r + 1, 0, 'Серая ячейка — нет отчетности за предыдущий квартал, квартальное значение не определить.', ST.note); sh.merge(r + 1, 0, r + 1, 8);
  }

  /* ── Удельный вес в валюте баланса ── */
  function structureSheet(X) {
    const sh = X.sheets[SH.W]; sheetHead(X, sh, 'Структура баланса: удельный вес статей в валюте баланса', { c2: '' }); let r = R0;
    for (const kind of ['A', 'L']) {
      const base = XW.addr; const baseCode = kind === 'A' ? '1600' : '1700';
      sh.set(r, 0, kind === 'A' ? 'АКТИВ' : 'ПАССИВ', ST.labB); sh.merge(r, 0, r, 2); r++;
      for (const [code, name] of M.CODES.filter(c => c[2] === kind)) {
        sh.set(r, 0, name, ST.lab); sh.set(r, 1, code, ST.code); sh.set(r, 2, '', ST.lab);
        X.periods.forEach((k, i) => { const c = FC + i; sh.set(r, c, { f: 'IF(' + X.bal(kind, baseCode, i) + '=0,"",' + X.bal(kind, code, i) + '/' + X.bal(kind, baseCode, i) + ')' }, ST.calc(NF.pct)); });
        r++;
      }
      r++;
    }
    sh.cf(XW.addr(R0, FC) + ':' + XW.addr(r, FC + Math.max(0, X.n - 1)), { type: 'dataBar', color: 'FFB7D8C6' });
  }

  /* ── Динамика г/г ── */
  function dynSheet(X, name, title, blocks) {
    const sh = X.sheets[name]; sheetHead(X, sh, title, { c2: '' }); let r = R0;
    sh.set(r, 0, 'Изменение к аналогичному периоду прошлого года: тыс. руб. (серая ячейка — нет сопоставимого периода)', ST.labB); sh.merge(r, 0, r, 2); r++;
    const start = r;
    for (const mode of ['abs', 'pct']) {
      if (mode === 'pct') { r++; sh.set(r, 0, 'Изменение к аналогичному периоду прошлого года: %', ST.labB); sh.merge(r, 0, r, 2); r++; }
      for (const [kind, codes] of blocks) for (const code of codes) {
        sh.set(r, 0, M.CODE_NAME[code], ST.lab); sh.set(r, 1, code, ST.code); sh.set(r, 2, mode === 'abs' ? 'тыс. руб.' : '%', ST.code);
        X.periods.forEach((k, i) => {
          const j = X.prevYear(i); const c = FC + i;
          if (j < 0) { sh.set(r, c, null, { fill: C.grey, bd: { all: 'thin' } }); return; }
          const cur = X.bal(kind, code, i), prev = X.bal(kind, code, j);
          sh.set(r, c, { f: mode === 'abs' ? cur + '-' + prev : 'IF(' + prev + '=0,"",ROUND(' + cur + '/' + prev + '-1,3))' }, ST.calc(mode === 'abs' ? NF.num : NF.pct));
        }); r++;
      }
    }
    sh.cf(XW.addr(start, FC) + ':' + XW.addr(r, FC + Math.max(0, X.n - 1)), { type: 'expression', formula: 'AND(ISNUMBER(' + XW.addr(start, FC) + '),' + XW.addr(start, 2, false, true) + '="%",ABS(' + XW.addr(start, FC) + ')>0.4)', dxf: { fill: 'FFFFF2B3' } });
  }

  /* ── Чистые активы (приложение 4.4.3) ── */
  function netAssets(X) {
    const sh = X.sheets[SH.NA]; sheetHead(X, sh, 'Чистые активы (п. 5.1.1 методики)', { c2: '' }); const R = {}; let r = R0;
    const line = (key, label, fn, nf, opt) => { R[key] = r; sh.set(r, 0, label, opt && opt.b ? ST.labB : ST.lab); sh.set(r, 1, (opt && opt.code) || '', ST.code); sh.set(r, 2, '', ST.lab); X.periods.forEach((k, i) => { const f = fn(i); sh.set(r, FC + i, f == null ? null : { f }, ST.calc(nf || NF.num, opt && opt.b, opt && opt.b ? C.sect : undefined)); }); r++; };
    line('ta', 'Активы (баланс)', i => X.bal('A', '1600', i), NF.num, { code: '1600' });
    line('lt', 'Долгосрочные обязательства', i => X.bal('L', '1400', i), NF.num, { code: '1400' });
    line('st', 'Краткосрочные обязательства', i => X.bal('L', '1500', i), NF.num, { code: '1500' });
    line('dbp', 'Доходы будущих периодов (исключаются из обязательств)', i => X.bal('L', '1530', i), NF.num, { code: '1530' });
    line('na', 'ЧИСТЫЕ АКТИВЫ = активы − (долгосрочные + краткосрочные − ДБП)', i => XW.addr(R.ta, FC + i) + '-(' + XW.addr(R.lt, FC + i) + '+' + XW.addr(R.st, FC + i) + '-' + XW.addr(R.dbp, FC + i) + ')', NF.num, { b: 1 });
    line('uk', 'Уставный капитал', i => X.bal('L', '1310', i), NF.num, { code: '1310' });
    line('dif', 'Чистые активы − уставный капитал (< 0 → формальный признак ликвидации, п. 5.3.12)', i => XW.addr(R.na, FC + i) + '-' + XW.addr(R.uk, FC + i));
    line('mx', 'Максимум ЧА за последние 12 месяцев (4 квартала)', i => { const q = qIdx(X.periods[i]); let j0 = i; while (j0 > 0 && qIdx(X.periods[j0 - 1]) >= q - 3 && qIdx(X.periods[j0 - 1]) < qIdx(X.periods[j0])) j0--; return 'MAX(' + XW.addr(R.na, FC + j0) + ':' + XW.addr(R.na, FC + i) + ')'; });
    line('drop', 'Снижение ЧА от максимума, % (≥ 25% → негативное явление 7.2.11)', i => 'IF(' + XW.addr(R.mx, FC + i) + '<=0,"",ROUND((' + XW.addr(R.mx, FC + i) + '-' + XW.addr(R.na, FC + i) + ')/' + XW.addr(R.mx, FC + i) + ',3))', NF.pct);
    line('sc', 'Балл по методике: ЧА > 0 → 10, иначе 0', i => 'IF(ISNUMBER(' + XW.addr(R.na, FC + i) + '),IF(' + XW.addr(R.na, FC + i) + '>0,10,0),"")', '0', { b: 1 });
    X.K.na = R; sh.set(r + 1, 0, 'Формула ЧА соответствует шаблону (1600 − (1400 + 1500 − 1530)): строка 1527 уже входит в 1520 и повторно не вычитается.', ST.note); sh.merge(r + 1, 0, r + 1, 8);
  }

  /* ── СПРАВОЧНИК: пороги ── */
  function reference(X) {
    const sh = X.sheets[SH.R]; sh.set(0, 0, 'Справочник критериальных значений (Методика ЮЛ от 14.08.2025 с Изменениями № 1 от 01.06.2026)', ST.title);
    sh.set(2, 0, 'Рентабельность по отраслям', ST.labB); ['Отрасль', 'Рпр: 0 баллов при <', 'Рпр: 5 баллов до (включ.)', 'Ра: 5 баллов до (включ.)', 'Рпр: 10 баллов при >'].forEach((t, c) => sh.set(3, c, t, ST.hdr));
    sh.width(0, 38); [1, 2, 3, 4].forEach(c => sh.width(c, 20)); sh.height(3, 42);
    INDUSTRY_TBL.forEach((row, k) => { const r = 4 + k; sh.set(r, 0, row[0], ST.lab); sh.set(r, 1, 1, ST.inp(NF.k1)); sh.set(r, 2, row[1], ST.inp(NF.k1)); sh.set(r, 3, row[2], ST.inp(NF.k1)); sh.set(r, 4, row[3], ST.inp(NF.k1)); });
    sh.set(8, 0, 'Для «Прочего» методика определяет Рпр: 5 баллов — 1…10,9%, 10 баллов — более 12,7%. Интервал 10,9…12,7% не определен: в расчете принимается 5 баллов (консервативно).', ST.note); sh.merge(8, 0, 8, 4); sh.height(8, 32);
    sh.set(10, 0, 'Прочие показатели (одинаковы для всех отраслей)', ST.labB);
    ['Показатель', 'Нижняя граница', 'Верхняя граница', 'Баллы: ниже / между / выше'].forEach((t, c) => sh.set(11, c, t, ST.hdr));
    const T = [['Кфн (автономия) = 1300/1600', 0.1, 0.2, '0 / 3 / 5'], ['Ко (обеспеченность СОС) = (1300−1100)/1200', 0, 0.2, '0 / 3 / 5'], ['Кл (текущая ликвидность)', 1, 1.3, '0 / 5 / 10'], ['Кпл (степень платежеспособности, мес.): чем меньше, тем лучше', 4, 8, 'выше 8 → 0; 4…8 → 5; ниже 4 → 10'], ['Коб (годовой)', 0.5, 2, '0 / 5 / 10'], ['Кобдз (годовой)', 0.8, 2, '0 / 5 / 10'], ['Рск, %', 1, 4, '0 / 5 / 10'], ['Бизнес-риск: сумма 12 факторов', 15, 20, 'ниже 15 → 0; 15…20 → 5; выше 20 → 10'], ['Класс кредитоспособности: границы баллов (1 / 2 / 3 класс)', 30, 50, '80 и выше → 1; 50…79 → 2; 30…49 → 3; 0…29 → 4']];
    X.K.thr = {}; const keys = ['ka', 'ko', 'kl', 'kpl', 'kob', 'kobdz', 'rsk', 'risk', 'cls'];
    T.forEach((row, k) => { const r = 12 + k; sh.set(r, 0, row[0], ST.lab); sh.set(r, 1, row[1], ST.inp(NF.k2)); sh.set(r, 2, row[2], ST.inp(NF.k2)); sh.set(r, 3, row[3], ST.lab); X.K.thr[keys[k]] = { lo: XW.quoteSheet(SH.R) + '!$B$' + (r + 1), hi: XW.quoteSheet(SH.R) + '!$C$' + (r + 1) }; });
    sh.width(3, 38);
    sh.set(23, 0, 'Классы по сумме баллов: 1 класс — 80 и более; 2 класс — 50…79; 3 класс — 30…49; 4 класс — 0…29 (п. 5.4.2). Бизнес-риск входит в сумму (макс. 10).', ST.note); sh.merge(23, 0, 23, 4); sh.height(23, 30);
    X.K.ind = XW.quoteSheet(SH.R) + '!$A$5:$E$8';
  }

  /* ── Коэффициенты (приложение 4.4.4) ── */
  function coefficients(X) {
    const sh = X.sheets[SH.K]; sheetHead(X, sh, 'Финансовые коэффициенты (п. 5.1 методики)', { c1: '', c2: 'Норма', w2: 22, w0: 56 });
    sh.set(2, 0, 'Отрасль (выберите из списка)', ST.labB); sh.set(2, 1, '', ST.lab); sh.set(2, 2, X.S.industry || 'Прочее', { b: 1, color: C.input, fill: C.yellow, bd: { all: 'thin' } });
    sh.validate('C3', { list: INDUSTRY_TBL.map(x => x[0]), title: 'Отрасль', prompt: 'Отрасль с наибольшей долей в доходах (п. 5.1.11 методики). От нее зависят пороги Рпр и Ра.' });
    sh.set(2, 3, X.S.industryNote || '', ST.note); X.K.indCell = XW.quoteSheet(SH.K) + '!$C$3';
    const R = {}; let r = R0;
    const put = (key, label, norm, fn, nf, o) => { o = o || {}; R[key] = r; sh.set(r, 0, label, o.b ? ST.labB : ST.lab); sh.set(r, 1, '', ST.lab); sh.set(r, 2, norm || '', ST.code); X.periods.forEach((k, i) => { const res = fn(i); if (res == null) sh.set(r, FC + i, null, { fill: C.grey, bd: { all: 'thin' } }); else sh.set(r, FC + i, { f: res.f }, ST.calc(nf, o.b, res.fill)); }); r++; };
    const a = (kind, code, i) => X.bal(kind, code, i); const p = (code, i) => X.ref(SH.P, X.rows.P[code], i);
    const T = (i) => XW.addr(R.T, FC + i);
    put('T', 'Т — число месяцев в отчетном периоде', '', i => ({ f: String(+X.periods[i].split('-')[1] * 3) }), '0');
    put('ca', 'Оборотные активы (1200) на конец периода', '', i => ({ f: a('A', '1200', i) }));
    put('ca0', 'Оборотные активы (1200) на начало периода', '', i => { const j = X.startIdx(i); return { f: a('A', '1200', j >= 0 ? j : i), fill: j >= 0 ? undefined : C.orange }; });
    put('ar', 'Дебиторская задолженность (1230) на конец периода', '', i => ({ f: a('A', '1230', i) }));
    put('ar0', 'Дебиторская задолженность (1230) на начало периода', '', i => { const j = X.startIdx(i); return { f: a('A', '1230', j >= 0 ? j : i), fill: j >= 0 ? undefined : C.orange }; });
    put('ta', 'Валюта баланса (1600) на конец периода', '', i => ({ f: a('A', '1600', i) }));
    put('ta0', 'Валюта баланса (1600) на начало периода', '', i => { const j = X.startIdx(i); return { f: a('A', '1600', j >= 0 ? j : i), fill: j >= 0 ? undefined : C.orange }; });
    put('base', 'База «начала периода»', '', i => { const j = X.startIdx(i); return { f: j >= 0 ? '"конец ' + (+X.periods[i].split('-')[0] - 1) + ' г."' : '"⚠ нет баланса на 31.12: взят конец периода"', fill: j >= 0 ? undefined : C.orange }; }, '@');
    r++;
    const g = (i) => XW.addr; const c = (key, i) => XW.addr(R[key], FC + i);
    put('ka', 'Кфн — автономия: 1300 / 1600', '> 0,2', i => ({ f: 'IF(' + a('A', '1600', i) + '=0,"",ROUND(' + a('L', '1300', i) + '/' + a('A', '1600', i) + ',2))' }), NF.k2, { b: 1 });
    put('ko', 'Ко — обеспеченность собств. оборотными средствами: (1300 − 1100) / 1200', '> 0,2', i => ({ f: 'IF(' + a('A', '1200', i) + '=0,"",ROUND((' + a('L', '1300', i) + '-' + a('A', '1100', i) + ')/' + a('A', '1200', i) + ',2))' }), NF.k2, { b: 1 });
    put('kl', 'Кл — текущая ликвидность: (1200 − 5445 − 5540 − 5320) / (1510 + 1520 + 1550)', '> 1,3', i => ({ f: 'IF((' + a('L', '1510', i) + '+' + a('L', '1520', i) + '+' + a('L', '1550', i) + ')=0,"",ROUND((' + a('A', '1200', i) + '-' + a('A', '5445', i) + '-' + a('A', '5540', i) + '-' + a('A', '5320', i) + ')/(' + a('L', '1510', i) + '+' + a('L', '1520', i) + '+' + a('L', '1550', i) + '),2))' }), NF.k2, { b: 1 });
    put('kpl', 'Кпл — степень платежеспособности, мес.: (1510+1520+1550+1410+1450) / (2110 : Т)', '< 4', i => ({ f: 'IF(' + p('2110', i) + '<=0,"",ROUND((' + a('L', '1510', i) + '+' + a('L', '1520', i) + '+' + a('L', '1550', i) + '+' + a('L', '1410', i) + '+' + a('L', '1450', i) + ')/(' + p('2110', i) + '/' + T(i) + '),2))' }), NF.k2, { b: 1 });
    put('kob', 'Коб — оборачиваемость оборотных средств (годовая): 2110 / среднее 1200 × 12/Т', '> 2', i => ({ f: 'IF((' + c('ca0', i) + '+' + c('ca', i) + ')<=0,"",ROUND(' + p('2110', i) + '/((' + c('ca0', i) + '+' + c('ca', i) + ')/2)*12/' + T(i) + ',2))' }), NF.k2, { b: 1 });
    put('kobdz', 'Кобдз — оборачиваемость дебиторской задолженности (годовая): 2110 / среднее 1230 × 12/Т', '> 2', i => ({ f: 'IF((' + c('ar0', i) + '+' + c('ar', i) + ')<=0,"",ROUND(' + p('2110', i) + '/((' + c('ar0', i) + '+' + c('ar', i) + ')/2)*12/' + T(i) + ',2))' }), NF.k2, { b: 1 });
    put('rpr', 'Рпр — рентабельность продаж, %: 2200 / 2110 × 100', 'по отрасли', i => ({ f: 'IF(' + p('2110', i) + '=0,"",ROUND(' + p('2200', i) + '/' + p('2110', i) + '*100,1))' }), NF.k1, { b: 1 });
    put('rsk', 'Рск — рентабельность собственного капитала, %: 2300 / 1300 × 100', '> 4', i => ({ f: 'IF(' + a('L', '1300', i) + '=0,"",ROUND(' + p('2300', i) + '/' + a('L', '1300', i) + '*100,1))' }), NF.k1, { b: 1 });
    put('ra', 'Ра — рентабельность активов, %: 2300 / среднее 1600 × 100', 'по отрасли', i => ({ f: 'IF((' + c('ta0', i) + '+' + c('ta', i) + ')<=0,"",ROUND(' + p('2300', i) + '/((' + c('ta0', i) + '+' + c('ta', i) + ')/2)*100,1))' }), NF.k1, { b: 1 });
    r++; put('d1', 'Д1 — длительность оборота оборотных средств, дней: 360 / Коб', '', i => ({ f: 'IF(ISNUMBER(' + c('kob', i) + '),IF(' + c('kob', i) + '=0,"",ROUND(360/' + c('kob', i) + ',0)),"")' }), '0');
    put('d2', 'Д2 — срок погашения дебиторской задолженности, дней: 360 / Кобдз', '', i => ({ f: 'IF(ISNUMBER(' + c('kobdz', i) + '),IF(' + c('kobdz', i) + '=0,"",ROUND(360/' + c('kobdz', i) + ',0)),"")' }), '0');
    X.K.k = R; sh.set(r + 1, 0, 'Оранжевая заливка: отчетности на конец прошлого года нет, в качестве «начала периода» взят остаток на конец периода (среднее = конец). Значения округлены по п. 9 методики (2 знака; проценты — 1 знак). Рск и Рпр — без приведения к году.', ST.note); sh.merge(r + 1, 0, r + 1, 8); sh.height(r + 1, 34);
  }

  AZ.exp = AZ.exp || {};
  Object.assign(AZ.exp, { SH, ST, NF, C, FC, HR, R0, newCtx, sheetHead, inputSheet, quarterBlock, structureSheet, dynSheet, netAssets, reference, coefficients, qIdx, INDUSTRY_TBL, tolOf });
})(typeof window !== 'undefined' ? window : globalThis);
