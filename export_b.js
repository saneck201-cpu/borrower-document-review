/* export_b.js — аналитические листы: СПРАВОЧНИК, Чистые активы, Коэффициенты, Фин показ, Нефин показат, Фин состояние. */
(function (root) {
  'use strict';
  const AZ = root.AZ; const M = AZ.model; const XW = root.XW; const E = AZ.exp; const ST = E.ST;
  const q = XW.quoteSheet;

  /** Универсальный построитель строк по периодам: rows = [{key,label,code?,f:(i)=>formula|null,st,sec,note}] */
  function rowsSheet(ctx, name, title, rows, opts) {
    opts = opts || {};
    const sh = ctx.wb.addSheet(name, { tab: opts.tab || 'FF70AD47' }); E.header(ctx, sh, title, opts.sub || E.subtitle(ctx), opts.first);
    ctx.rows[name] = {}; let r = 4;
    for (const row of rows) {
      if (row.sec) { sh.set(r, 0, null, ST.sec); sh.set(r, 1, row.sec, ST.sec); for (let c = 2; c < 3 + ctx.np; c++) sh.set(r, c, null, ST.sec); r++; continue; }
      ctx.rows[name][row.key] = r; row.r = r; r++;
    }
    for (const row of rows) {
      if (row.sec) continue; const st = row.st || ST.calc; const nm = row.bold ? ST.totName : ST.name;
      sh.set(row.r, 0, row.code || '', row.bold ? ST.totName : ST.code); sh.set(row.r, 1, row.label, nm); sh.set(row.r, 2, row.unit || '', row.bold ? ST.totName : ST.code);
      for (let i = 0; i < ctx.np; i++) { const f = row.f(i); if (f === null || f === undefined) sh.set(row.r, 3 + i, null, st); else if (typeof f === 'number') sh.set(row.r, 3 + i, f, st); else sh.set(row.r, 3 + i, { f }, st); }
      if (row.height) sh.height(row.r, row.height);
    }
    return sh;
  }
  E.rowsSheet = rowsSheet;
  const R = (ctx, name, key, i) => q(name) + '!' + XW.colName(3 + i) + (ctx.rows[name][key] + 1);   // ссылка на строку листа по ключу
  E.R = R;

  E.buildReference = function (ctx) {
    const sh = ctx.wb.addSheet('СПРАВОЧНИК', { tab: 'FF7F7F7F' });
    sh.set(0, 0, 'Справочник: пороговые значения и шкалы (Методика ЮЛ 2025 + изменения № 1 от 01.06.2026)', ST.title);
    sh.width(0, 22); sh.width(1, 22); sh.width(2, 22); sh.width(3, 22); sh.width(4, 70);
    ['Отрасль', 'Рпр: верхний порог, % (по Изм. № 1)', 'Ра: верхний порог, %', 'Рпр: порог, применяемый в расчете, %', 'Примечание'].forEach((t, c) => sh.set(3, c, t, c === 4 ? ST.headL : ST.head)); sh.height(3, 44);
    M.INDUSTRIES.forEach((x, k) => {
      const eff = x[0] === 'Прочее' ? 12.7 : x[1];
      sh.set(4 + k, 0, x[0], ST.nameB); sh.set(4 + k, 1, x[1], ST.inp); sh.set(4 + k, 2, x[2], ST.inp); sh.set(4 + k, 3, eff, ST.inp);
      sh.set(4 + k, 4, x[0] === 'Прочее' ? 'Для «Прочего» методика не определяет балл Рпр на интервале 10,9–12,7 %: консервативно принято 5 баллов до 12,7 %.' : '', ST.note);
    });
    sh.set(10, 0, 'Шкалы баллов (максимум 100)', ST.sec); for (let c = 1; c < 5; c++) sh.set(10, c, null, ST.sec);
    const scale = [
      ['Кфн = 1300/1600', '<0,1 → 0; 0,1–0,2 → 3; >0,2 → 5'], ['Ко = (1300−1100)/1200', '<0 → 0; 0–0,2 → 3; >0,2 → 5'],
      ['Кл = (1200−5445−5540−5320)/(1510+1520+1550)', '<1 → 0; 1–1,3 → 5; >1,3 → 10'], ['Кпл = (1510+1520+1550+1410+1450)/(2110/мес.)', '>8 → 0; 4–8 → 5; <4 → 10 (в месяцах)'],
      ['Коб = выручка/средние оборотные ×12/T', '<0,5 → 0; 0,5–2 → 5; >2 → 10'], ['Кобдз = выручка/средняя ДЗ ×12/T', '<0,8 → 0; 0,8–2 → 5; >2 → 10'],
      ['Рпр = 2200/2110×100', '<1 → 0; 1–порог отрасли → 5; > порога → 10'], ['Ра = 2300/средние активы×100', '<1 → 0; 1–порог отрасли → 5; > порога → 10'],
      ['Рск = 2300/1300×100', '<1 → 0; 1–4 → 5; >4 → 10'], ['Чистые активы', '>0 → 10; иначе 0'], ['Бизнес-риск (сумма 12 факторов, макс. 31)', '<15 → 0; 15–20 → 5; >20 → 10'],
      ['Класс по сумме баллов', '80–100 → 1; 50–79 → 2; 30–49 → 3; 0–29 → 4'],
    ];
    scale.forEach((s, k) => { sh.set(11 + k, 0, s[0], ST.name); sh.merge(11 + k, 0, 11 + k, 1); sh.set(11 + k, 2, s[1], ST.name); sh.merge(11 + k, 2, 11 + k, 4); });
    sh.set(25, 0, 'Матрица финансового положения (п. 7.4): число негативных явлений', ST.sec); for (let c = 1; c < 5; c++) sh.set(25, c, null, ST.sec);
    [['Класс', 'Хорошее', 'Среднее', 'Плохое'], ['1', '0', 'любое', '«Х» (не применимо)'], ['2', '≤ 1', '≤ 3', '> 3 или банкрот'], ['3', '«Х»', '≤ 2', '> 2 или банкрот'], ['4', '«Х»', '≤ 1', '> 1 или банкрот']].forEach((a, k) => a.forEach((t, c) => sh.set(26 + k, c, t, k === 0 ? ST.head : ST.name)));
    sh.set(32, 0, 'Единицы: все суммы в тыс. руб.; Т — число месяцев нарастающего периода (3, 6, 9, 12); баланс «на начало» = конец предыдущего года.', ST.note);
    ctx.refRange = "'СПРАВОЧНИК'!$A$5:$D$8";
  };

  E.buildNetAssets = function (ctx) {
    const P = i => ctx.periods[i]; const qi = i => P(i).year * 4 + P(i).q; const ref = E.ref || ((c, i) => ctx.ref(c, i));
    const rows = [
      { key: 'a', label: 'Баланс (актив)', code: '1600', f: i => ctx.ref('1600', i) },
      { key: 'd', label: 'Долгосрочные обязательства', code: '1400', f: i => ctx.ref('1400', i) },
      { key: 'k', label: 'Краткосрочные обязательства', code: '1500', f: i => ctx.ref('1500', i) },
      { key: 'dbp', label: 'Доходы будущих периодов (учтены в 1500; не вычитаются)', code: '1530', f: i => ctx.ref('1530', i) },
      { key: 'ob', label: 'Обязательства, принимаемые к расчету', f: i => `${R(ctx, 'Чистые активы', 'd', i)}+${R(ctx, 'Чистые активы', 'k', i)}-${R(ctx, 'Чистые активы', 'dbp', i)}` },
      { key: 'na', label: 'ЧИСТЫЕ АКТИВЫ', bold: 1, st: ST.tot, f: i => `${R(ctx, 'Чистые активы', 'a', i)}-${R(ctx, 'Чистые активы', 'ob', i)}` },
      { key: 'uk', label: 'Уставный капитал', code: '1310', f: i => ctx.ref('1310', i) },
      { key: 'dif', label: 'Чистые активы − уставный капитал', f: i => `${R(ctx, 'Чистые активы', 'na', i)}-${R(ctx, 'Чистые активы', 'uk', i)}` },
      { key: 'max', label: 'Максимум ЧА за последние 12 месяцев', f: i => { const l = []; for (let j = 0; j <= i; j++) if (qi(i) - qi(j) <= 4) l.push(R(ctx, 'Чистые активы', 'na', j)); return 'MAX(' + l.join(',') + ')'; } },
      { key: 'drop', label: 'Снижение ЧА от максимума, %', st: ST.pct, f: i => `IFERROR(IF(${R(ctx, 'Чистые активы', 'max', i)}>0,(${R(ctx, 'Чистые активы', 'max', i)}-${R(ctx, 'Чистые активы', 'na', i)})/${R(ctx, 'Чистые активы', 'max', i)},0),"")` },
      { key: 'pt', label: 'Балл по чистым активам (макс. 10)', bold: 1, st: ST.tot, f: i => `IF(${R(ctx, 'Чистые активы', 'na', i)}>0,10,0)` },
    ];
    // R() обращается к ctx.rows до построения — задаем строки заранее
    ctx.rows['Чистые активы'] = {}; let r = 4; rows.forEach(x => { x.r = r; ctx.rows['Чистые активы'][x.key] = r++; });
    const sh = rowsSheetPre(ctx, 'Чистые активы', 'Расчет чистых активов (прил. 4.4.3)', rows);
    sh.set(2, 0, 'ЧА = 1600 − (1400 + 1500 − 1530). Статья 1527 уже входит в 1520 — повторно не вычитается.', ST.note);
  };
  /** Как rowsSheet, но ключи строк уже заданы (ссылки внутри формул строк) */
  function rowsSheetPre(ctx, name, title, rows, opts) {
    opts = opts || {}; const sh = ctx.wb.addSheet(name, { tab: opts.tab || 'FF70AD47' }); E.header(ctx, sh, title, opts.sub || E.subtitle(ctx), opts.first);
    for (const row of rows) {
      if (row.sec) { sh.set(row.r, 0, null, ST.sec); sh.set(row.r, 1, row.sec, ST.sec); for (let c = 2; c < 3 + ctx.np; c++) sh.set(row.r, c, null, ST.sec); continue; }
      const st = row.st || ST.calc; sh.set(row.r, 0, row.code || '', row.bold ? ST.totName : ST.code); sh.set(row.r, 1, row.label, row.bold ? ST.totName : ST.name); sh.set(row.r, 2, row.unit || '', row.bold ? ST.totName : ST.code);
      for (let i = 0; i < ctx.np; i++) { const f = row.f(i); if (f === null || f === undefined) sh.set(row.r, 3 + i, null, st); else sh.set(row.r, 3 + i, { f }, st); }
    }
    return sh;
  }
  E.rowsSheetPre = rowsSheetPre;
  function layout(ctx, name, rows) { ctx.rows[name] = {}; let r = 4; rows.forEach(x => { x.r = r; if (x.key) ctx.rows[name][x.key] = r; r++; }); }
  E.layout = layout;

  E.buildCoefficients = function (ctx) {
    const N = 'Коэффициенты'; const rf = (c, i) => ctx.ref(c, i); const me = (k, i) => R(ctx, N, k, i);
    const beg = (code, i) => { const j = ctx.startIdx(i); return j >= 0 ? rf(code, j) : rf(code, i); };
    const begNote = i => (ctx.startIdx(i) >= 0 ? '' : 'нет баланса на начало года — взят баланс на конец периода');
    const rows = [
      { sec: 'ИСХОДНЫЕ ВЕЛИЧИНЫ' },
      { key: 'T', label: 'Месяцев в периоде (T)', f: i => String(ctx.periods[i].months), unit: 'мес.' },
      { key: 'rev', label: 'Выручка (2110), нарастающим итогом', code: '2110', f: i => rf('2110', i) },
      { key: 'a1', label: 'Активы на начало периода (1600)', code: '1600', f: i => beg('1600', i) },
      { key: 'a2', label: 'Активы на конец периода (1600)', code: '1600', f: i => rf('1600', i) },
      { key: 'ca1', label: 'Оборотные активы на начало (1200)', code: '1200', f: i => beg('1200', i) },
      { key: 'ca2', label: 'Оборотные активы на конец (1200)', code: '1200', f: i => rf('1200', i) },
      { key: 'ar1', label: 'Дебиторская задолженность на начало (1230)', code: '1230', f: i => beg('1230', i) },
      { key: 'ar2', label: 'Дебиторская задолженность на конец (1230)', code: '1230', f: i => rf('1230', i) },
      { sec: 'КОЭФФИЦИЕНТЫ (прил. 4.4.4)' },
      { key: 'kfn', label: 'Кфн — финансовая независимость = 1300 / 1600', st: ST.num2, f: i => `IFERROR(${rf('1300', i)}/${rf('1600', i)},"")` },
      { key: 'ko', label: 'Ко — обеспеченность собственными оборотными = (1300 − 1100) / 1200', st: ST.num2, f: i => `IFERROR((${rf('1300', i)}-${rf('1100', i)})/${rf('1200', i)},"")` },
      { key: 'kl', label: 'Кл — текущая ликвидность = (1200 − 5445 − 5540 − 5320) / (1510 + 1520 + 1550)', st: ST.num2, f: i => `IFERROR((${rf('1200', i)}-${rf('5445', i)}-${rf('5540', i)}-${rf('5320', i)})/(${rf('1510', i)}+${rf('1520', i)}+${rf('1550', i)}),"")` },
      { key: 'kpl', label: 'Кпл — степень платежеспособности, мес. = (1510+1520+1550+1410+1450) / (2110 / T)', st: ST.num2, f: i => `IFERROR((${rf('1510', i)}+${rf('1520', i)}+${rf('1550', i)}+${rf('1410', i)}+${rf('1450', i)})/(${me('rev', i)}/${me('T', i)}),"")` },
      { key: 'kob', label: 'Коб — оборачиваемость оборотных средств (годовая) = 2110 / среднее(1200) × 12 / T', st: ST.num2, f: i => `IFERROR(${me('rev', i)}/((${me('ca1', i)}+${me('ca2', i)})/2)*12/${me('T', i)},"")` },
      { key: 'kobdz', label: 'Кобдз — оборачиваемость ДЗ (годовая) = 2110 / среднее(1230) × 12 / T', st: ST.num2, f: i => `IFERROR(${me('rev', i)}/((${me('ar1', i)}+${me('ar2', i)})/2)*12/${me('T', i)},"")` },
      { key: 'rpr', label: 'Рпр — рентабельность продаж, % = 2200 / 2110 × 100', st: ST.num2, f: i => `IFERROR(${rf('2200', i)}/${me('rev', i)}*100,"")` },
      { key: 'ra', label: 'Ра — рентабельность активов, % = 2300 / среднее(1600) × 100', st: ST.num2, f: i => `IFERROR(${rf('2300', i)}/((${me('a1', i)}+${me('a2', i)})/2)*100,"")` },
      { key: 'rsk', label: 'Рск — рентабельность СК, % = 2300 / 1300 × 100', st: ST.num2, f: i => `IFERROR(${rf('2300', i)}/${rf('1300', i)}*100,"")` },
      { key: 'na', label: 'Чистые активы, тыс. руб.', st: ST.calc, f: i => R(ctx, 'Чистые активы', 'na', i) },
      { sec: 'ПОРОГИ ОТРАСЛИ (выбор в ячейке C2)' },
      { key: 'hRpr', label: 'Рпр: порог 10 баллов, %', st: ST.num2, f: i => `VLOOKUP($C$2,${ctx.refRange},4,FALSE)` },
      { key: 'hRa', label: 'Ра: порог 10 баллов, %', st: ST.num2, f: i => `VLOOKUP($C$2,${ctx.refRange},3,FALSE)` },
    ];
    layout(ctx, N, rows);
    // сдвиг ссылок: строка 2 (индекс 1) занята выбором отрасли
    const sh = rowsSheetPre(ctx, N, 'Коэффициенты финансового состояния', rows);
    sh.set(1, 0, 'Отрасль:', ST.nameB); sh.set(1, 1, 'Выберите отрасль (определяется по ОКВЭД, можно изменить) →', ST.sub);
    sh.set(1, 2, ctx.data.borrower.industry || 'Прочее', { b: 1, fill: ST.YEL, bd: 1, h: 'center' });
    sh.validate('C2', { list: M.INDUSTRIES.map(x => x[0]), title: 'Отрасль', prompt: 'Определяет пороги Рпр и Ра' });
    // заметки о запасных значениях
    ctx.periods.forEach((p, i) => { const n = begNote(i); if (n) { sh.style(ctx.rows[N]['a1'], 3 + i, { fill: ST.ORG, nf: '#,##0' }); ctx.addFlag(N, '1600', i, n, 'info'); } });
    ctx.cNames = N;
  };

  E.buildScores = function (ctx) {
    const N = 'Фин показ'; const C = 'Коэффициенты'; const c = (k, i) => R(ctx, C, k, i); const me = (k, i) => R(ctx, N, k, i);
    const sc = (k, f) => i => `IF(ISNUMBER(${c(k, i)}),${f(c(k, i), i)},"")`;
    const defs = [
      ['kfn', 'Кфн — финансовая независимость', 5, x => `IF(${x}<0.1,0,IF(${x}<=0.2,3,5))`],
      ['ko', 'Ко — обеспеченность собственными оборотными', 5, x => `IF(${x}<0,0,IF(${x}<=0.2,3,5))`],
      ['kl', 'Кл — текущая ликвидность', 10, x => `IF(${x}<1,0,IF(${x}<=1.3,5,10))`],
      ['kpl', 'Кпл — платежеспособность (мес.)', 10, x => `IF(${x}>8,0,IF(${x}>=4,5,10))`],
      ['kob', 'Коб — оборачиваемость оборотных средств', 10, x => `IF(${x}<0.5,0,IF(${x}<=2,5,10))`],
      ['kobdz', 'Кобдз — оборачиваемость дебиторской задолженности', 10, x => `IF(${x}<0.8,0,IF(${x}<=2,5,10))`],
      ['rpr', 'Рпр — рентабельность продаж', 10, (x, i) => `IF(${x}<1,0,IF(${x}<=${c('hRpr', i)},5,10))`],
      ['ra', 'Ра — рентабельность активов', 10, (x, i) => `IF(${x}<1,0,IF(${x}<=${c('hRa', i)},5,10))`],
      ['rsk', 'Рск — рентабельность собственного капитала', 10, x => `IF(${x}<1,0,IF(${x}<=4,5,10))`],
    ];
    const rows = [{ sec: 'ЗНАЧЕНИЯ ПОКАЗАТЕЛЕЙ' }];
    defs.forEach(d => rows.push({ key: 'v_' + d[0], label: d[1], st: ST.num2, f: i => c(d[0], i) }));
    rows.push({ sec: 'БАЛЛЫ ПО ПОКАЗАТЕЛЯМ' });
    defs.forEach(d => rows.push({ key: 'p_' + d[0], label: d[1] + ' (макс. ' + d[2] + ')', st: ST.calc, f: sc(d[0], (x, i) => d[3](x, i)) }));
    rows.push({ key: 'p_na', label: 'Чистые активы (макс. 10)', f: i => R(ctx, 'Чистые активы', 'pt', i) });
    const pk = ['p_na'].concat(defs.map(d => 'p_' + d[0]));
    rows.push({ sec: 'ИТОГО' });
    rows.push({ key: 'n', label: 'Рассчитано показателей из 10', f: i => 'COUNT(' + pk.map(k => me(k, i)).join(',') + ')' });
    rows.push({ key: 'fin', label: 'Баллы по финансовым показателям (макс. 90)', bold: 1, st: ST.tot, f: i => 'SUM(' + pk.map(k => me(k, i)).join(',') + ')' });
    rows.push({ key: 'risk', label: 'Балл бизнес-риска (макс. 10; лист «Нефин показат»)', f: i => `'Нефин показат'!$D$20` });
    rows.push({ key: 'tot', label: 'ИТОГО БАЛЛОВ (макс. 100)', bold: 1, st: ST.tot, f: i => `IF(AND(ISNUMBER(${me('risk', i)}),${me('n', i)}=10),${me('fin', i)}+${me('risk', i)},"")` });
    const cls = x => `IF(${x}>=80,1,IF(${x}>=50,2,IF(${x}>=30,3,4)))`;
    rows.push({ key: 'cls', label: 'КЛАСС по сумме баллов', bold: 1, st: ST.tot, f: i => `IF(ISNUMBER(${me('tot', i)}),${cls(me('tot', i))},"")` });
    rows.push({ key: 'clsMin', label: 'Класс до оценки бизнес-риска: лучший случай (риск = 10)', f: i => `IF(${me('n', i)}=10,${cls(me('fin', i) + '+10')},"")` });
    rows.push({ key: 'clsMax', label: 'Класс до оценки бизнес-риска: худший случай (риск = 0)', f: i => `IF(${me('n', i)}=10,${cls(me('fin', i))},"")` });
    layout(ctx, N, rows); const sh = rowsSheetPre(ctx, N, 'Баллы по финансовым показателям (прил. 4.4.5–4.4.7)', rows);
    sh.set(2, 0, 'Показатель, который невозможно рассчитать (нет данных), не оценивается; итог и класс выводятся только при 10 рассчитанных показателях и заполненном бизнес-риске.', ST.note);
    const lastc = XW.colName(2 + ctx.np);
    sh.cf(`D${ctx.rows[N]['cls'] + 1}:${lastc}${ctx.rows[N]['cls'] + 1}`, { type: 'cellIs', op: 'equal', formula: ['4'], dxf: { color: 'FF9C0006', fill: 'FFFFC7CE' } });
    sh.cf(`D${ctx.rows[N]['cls'] + 1}:${lastc}${ctx.rows[N]['cls'] + 1}`, { type: 'cellIs', op: 'equal', formula: ['1'], dxf: { color: 'FF006100', fill: 'FFC6EFCE' } });
    sh.cf(`D${ctx.rows[N]['p_kfn'] + 1}:${lastc}${ctx.rows[N]['p_na'] + 1}`, { type: 'cellIs', op: 'equal', formula: ['0'], dxf: { color: 'FF9C0006', fill: 'FFFFC7CE' } });
    ctx.sc = { pk, defs };
  };

  E.buildNonFin = function (ctx) {
    const sh = ctx.wb.addSheet('Нефин показат', { tab: 'FFED7D31' });
    sh.set(0, 0, 'Нефинансовые показатели — бизнес-риск (прил. 4.4.2)', ST.title); sh.set(1, 0, E.subtitle(ctx), ST.sub);
    ['№', 'Фактор бизнес-риска (пункт методики)', 'Макс.', 'Балл', 'Допустимые значения и критерии', 'Основание / источник'].forEach((t, c) => sh.set(3, c, t, c === 1 || c === 4 || c === 5 ? ST.headL : ST.head));
    sh.width(0, 5); sh.width(1, 46); sh.width(2, 7); sh.width(3, 9); sh.width(4, 84); sh.width(5, 46); sh.height(3, 28); sh.freeze(4, 2);
    const fs = ctx.data.factors || {}; const last = ctx.np - 1;
    M.FACTORS.forEach((f, k) => {
      const r = 4 + k; const v = fs[k]; sh.set(r, 0, k + 1, ST.code); sh.set(r, 1, f.n + ' (п. ' + f.p + ')', ST.name); sh.set(r, 2, f.max, ST.code);
      let cellSt = { b: 1, h: 'center', bd: 1, color: ST.BLUE };
      if (k === 11 && last >= 0) { sh.set(r, 3, { f: `IF(ISNUMBER(${R(ctx, 'Чистые активы', 'dif', last)}),IF(${R(ctx, 'Чистые активы', 'dif', last)}<0,0,2),"")` }, Object.assign({}, cellSt, { fill: ST.ORG, color: 'FF000000' })); sh.set(r, 5, 'Авто: чистые активы на последнюю дату сравнены с уставным капиталом; сведения о ликвидации/банкротстве проверяются вручную.', ST.note); }
      else if (v && v.score !== undefined && v.score !== null) { sh.set(r, 3, v.score, Object.assign({}, cellSt, v.auto ? { fill: ST.ORG } : {})); sh.set(r, 5, v.note || '', ST.note); }
      else { sh.set(r, 3, null, Object.assign({}, cellSt, { fill: ST.YEL })); sh.set(r, 5, (v && v.note) || 'Требует суждения аналитика', ST.note); }
      sh.set(r, 4, f.opts.map(o => o[0] + ' — ' + o[1]).join('\n'), ST.name); sh.height(r, 14 * f.opts.length + 6);
      sh.validate('D' + (r + 1), { list: f.allowed.map(String), title: f.n, prompt: 'Допустимые значения: ' + f.allowed.join(', ') });
    });
    sh.set(17, 1, 'Сумма баллов (макс. 31)', ST.totName); sh.set(17, 3, { f: 'SUM(D5:D16)' }, ST.tot);
    sh.set(18, 1, 'Заполнено факторов из 12', ST.name); sh.set(18, 3, { f: 'COUNT(D5:D16)' }, ST.calc);
    sh.set(19, 1, 'БАЛЛ БИЗНЕС-РИСКА (<15 → 0; 15–20 → 5; >20 → 10)', ST.totName); sh.set(19, 3, { f: 'IF(D19=12,IF(D18<15,0,IF(D18<=20,5,10)),"")' }, ST.tot);
    sh.set(19, 4, 'Выводится только когда заполнены все 12 факторов. Жёлтые ячейки — не заполнены (нужно суждение), оранжевые — предложены автоматически, проверьте.', ST.note);
    sh.cf('D5:D16', { type: 'expression', formula: ['AND(NOT(ISNUMBER(D5)))'], dxf: { fill: 'FFFFF2CC' } });
  };

  E.buildNegatives = function (ctx) {
    const sh = ctx.wb.addSheet('Фин состояние', { tab: 'FFC00000' }); const last = ctx.np - 1; const N = 'Чистые активы';
    sh.set(0, 0, 'Негативные явления и финансовое положение (п. 7.2, 7.4)', ST.title); sh.set(1, 0, E.subtitle(ctx) + ' · оценка на ' + (last >= 0 ? ctx.periods[last].label : '—'), ST.sub);
    ['п.', 'Негативное явление', 'Проверка', 'Значение', 'Признак (1 — есть, 0 — нет)', 'Комментарий'].forEach((t, c) => sh.set(3, c, t, c === 1 || c === 5 ? ST.headL : ST.head));
    sh.width(0, 7); sh.width(1, 70); sh.width(2, 13); sh.width(3, 16); sh.width(4, 16); sh.width(5, 60); sh.height(3, 32); sh.freeze(4, 2);
    const rf = (c, i) => ctx.ref(c, i); const py = last >= 0 ? ctx.prevYearIdx(last) : -1; const fsn = ctx.data.negatives || {};
    const growth = (code, a, b) => `IFERROR(${rf(code, a)}/${rf(code, b)}-1,"")`;
    const auto = k => {
      if (last < 0) return null; const t = M.NEGATIVES[k].auto;
      if (t === 'na') return { v: R(ctx, N, 'na', last), f: `IF(ISNUMBER(D${5 + k}),IF(D${5 + k}<0,1,0),"")`, n: 'Чистые активы на последнюю дату' };
      if (t === 'nadrop') return { v: R(ctx, N, 'drop', last), f: `IF(ISNUMBER(D${5 + k}),IF(D${5 + k}>=0.25,1,0),"")`, n: 'Снижение ЧА от максимума за 12 мес.; 25% и более' , pct: 1 };
      if (t === 'zero') return { v: `${rf('1200', last)}+${rf('1500', last)}`, f: `IF(ISNUMBER(D${5 + k}),IF(AND(${rf('1200', last)}=0,${rf('1500', last)}=0),1,0),"")`, n: 'Оба раздела нулевые — сверьте с оборотами по счетам в Банке за 180 дней' };
      if (py < 0) return { v: null, f: null, n: 'Нет периода прошлого года для сравнения — заполните вручную' };
      if (t === 'rev') return { v: growth('2110', last, py), f: `IF(ISNUMBER(D${5 + k}),IF(D${5 + k}<-0.4,1,0),"")`, n: 'Рост/снижение выручки г/г; при снижении более 40% проверьте сезонность', pct: 1 };
      if (t === 'margin') return { v: `IFERROR(${R(ctx, 'Коэффициенты', 'rpr', last)}/${R(ctx, 'Коэффициенты', 'rpr', py)}-1,"")`, f: `IF(ISNUMBER(D${5 + k}),IF(AND(${R(ctx, 'Коэффициенты', 'rpr', py)}>0,D${5 + k}<-0.4),1,0),"")`, n: 'Изменение Рпр г/г (при положительной базе)', pct: 1 };
      if (t === 'ap') return { v: growth('1520', last, py), f: `IF(ISNUMBER(D${5 + k}),IF(AND(D${5 + k}>0.4,IFERROR(${rf('2120', last)}/${rf('2120', py)}-1,0)<0.4),1,0),"")`, n: 'Рост кредиторской задолженности г/г при росте себестоимости менее 40%', pct: 1 };
      if (t === 'ar') return { v: growth('1230', last, py), f: `IF(ISNUMBER(D${5 + k}),IF(AND(D${5 + k}>0.4,IFERROR(${rf('2110', last)}/${rf('2110', py)}-1,0)<0.4),1,0),"")`, n: 'Рост дебиторской задолженности г/г при росте выручки менее 40%', pct: 1 };
      return null;
    };
    M.NEGATIVES.forEach((n, k) => {
      const r = 4 + k; sh.set(r, 0, n.p, ST.code); sh.set(r, 1, n.t, ST.name); sh.height(r, 30); const a = n.auto ? auto(k) : null; const man = fsn[k];
      if (a && a.v) { sh.set(r, 2, 'авто', ST.code); sh.set(r, 3, { f: a.v }, a.pct ? ST.pct : { nf: '#,##0;[Red]-#,##0;0' }); sh.set(r, 4, { f: a.f }, { b: 1, h: 'center', bd: 1, fill: ST.ORG }); sh.set(r, 5, a.n, ST.note); }
      else {
        sh.set(r, 2, 'вручную', ST.code); sh.set(r, 3, null, ST.calc);
        const v = man && man.value !== undefined && man.value !== null ? man.value : null;
        sh.set(r, 4, v, { b: 1, h: 'center', bd: 1, color: ST.BLUE, fill: v === null ? ST.YEL : undefined }); sh.set(r, 5, (man && man.note) || (a && a.n) || 'Проверяется по документам и сведениям Банка', ST.note);
      }
      sh.validate('E' + (r + 1), { list: ['0', '1'], title: n.p, prompt: '1 — явление есть, 0 — нет' });
    });
    sh.set(17, 1, 'Количество негативных явлений', ST.totName); sh.set(17, 4, { f: 'SUM(E5:E16)' }, ST.tot);
    sh.set(18, 1, 'Оценено явлений из 12', ST.name); sh.set(18, 4, { f: 'COUNT(E5:E16)' }, ST.calc);
    sh.set(19, 1, 'Банкротство / иные признаки несостоятельности (1 — да, 0 — нет)', ST.name); sh.set(19, 4, (ctx.data.bankrupt === 1 || ctx.data.bankrupt === 0) ? ctx.data.bankrupt : null, { b: 1, h: 'center', bd: 1, color: ST.BLUE, fill: (ctx.data.bankrupt === 1 || ctx.data.bankrupt === 0) ? undefined : ST.YEL });
    sh.validate('E20', { list: ['0', '1'], title: 'Банкротство', prompt: '1 — есть, 0 — нет' });
    sh.set(20, 1, 'Класс по сумме баллов (лист «Фин показ», последний период)', ST.name); sh.set(20, 4, last >= 0 ? { f: R(ctx, 'Фин показ', 'cls', last) } : null, ST.tot);
    sh.set(21, 1, 'ФИНАНСОВОЕ ПОЛОЖЕНИЕ по матрице п. 7.4 (предварительное)', ST.totName);
    const cl = 'E21', n = 'E18', bk = 'E20';
    sh.set(21, 4, { f: `IF(NOT(ISNUMBER(${cl})),"укажите бизнес-риск",IF(${bk}=1,"плохое",IF(${cl}=1,IF(${n}=0,"хорошее","среднее"),IF(${cl}=2,IF(${n}<=1,"хорошее",IF(${n}<=3,"среднее","плохое")),IF(${cl}=3,IF(${n}<=2,"среднее","плохое"),IF(${n}<=1,"среднее","плохое"))))))` }, { b: 1, h: 'center', bd: 1, fill: ST.GREY });
    sh.set(22, 1, { f: `IF(AND(${n}>=1,E22="хорошее"),"Внимание: по п. 7.2 при наличии негативного явления положение не может быть хорошим — матрица п. 7.4 допускает ≤1 для 2 класса; примите решение по существу.",IF(COUNT(E5:E16)<12,"Оценены не все явления — положение предварительное.",""))` }, { color: 'FFC00000', wrap: 1 });
    sh.height(22, 30);
    sh.set(24, 1, 'Положение — предварительный расчетный результат. Итоговое суждение выносит аналитик с учетом обслуживания долга, достоверности отчетности и иной информации.', ST.note);
    sh.cf('E22', { type: 'cellIs', op: 'equal', formula: ['"плохое"'], dxf: { color: 'FF9C0006', fill: 'FFFFC7CE' } });
    sh.cf('E22', { type: 'cellIs', op: 'equal', formula: ['"хорошее"'], dxf: { color: 'FF006100', fill: 'FFC6EFCE' } });
    ctx.negSheet = sh;
  };
})(typeof window !== 'undefined' ? window : globalThis);
