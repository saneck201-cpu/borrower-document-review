/* export_c.js — Сводка, графики, Проверки, Источники, Документы и общий сборщик книги. */
(function (root) {
  'use strict';
  const AZ = root.AZ; const M = AZ.model; const XW = root.XW; const E = AZ.exp; const ST = E.ST; const q = XW.quoteSheet;

  E.buildCharts = function (ctx) {
    const sh = ctx.wb.addSheet('графики', { tab: 'FF5B9BD5' }); const K = Math.min(ctx.np, 12); const i0 = ctx.np - K;
    sh.set(0, 0, 'Графики (последние ' + K + ' периодов)', ST.title); sh.width(0, 34); for (let c = 1; c <= K; c++) sh.width(c, 12);
    sh.set(2, 0, 'Показатель', ST.headL); for (let k = 0; k < K; k++) sh.set(2, 1 + k, ctx.periods[i0 + k].label, ST.head);
    const qr = ctx.qBlock.rows; const series = [
      ['Выручка за квартал, тыс. руб.', i => `'Приб_Уб'!${XW.colName(3 + i)}${qr['2110'] + 1}`, '#,##0'],
      ['Прибыль от продаж за квартал', i => `'Приб_Уб'!${XW.colName(3 + i)}${qr['2200'] + 1}`, '#,##0'],
      ['Чистая прибыль за квартал', i => `'Приб_Уб'!${XW.colName(3 + i)}${qr['2400'] + 1}`, '#,##0'],
      ['Чистые активы', i => E.R(ctx, 'Чистые активы', 'na', i), '#,##0'],
      ['Уставный капитал', i => E.R(ctx, 'Чистые активы', 'uk', i), '#,##0'],
      ['Валюта баланса', i => ctx.ref('1600', i), '#,##0'],
      ['Собственный капитал (1300)', i => ctx.ref('1300', i), '#,##0'],
      ['Баллы по фин. показателям (из 90)', i => E.R(ctx, 'Фин показ', 'fin', i), '0'],
    ];
    series.forEach((s, n) => { sh.set(3 + n, 0, s[0], ST.name); for (let k = 0; k < K; k++) sh.set(3 + n, 1 + k, { f: s[1](i0 + k) }, { nf: s[2] }); });
    const cats = { sheet: 'графики', r1: 2, c1: 1, r2: 2, c2: K }; const sr = n => ({ name: series[n][0], vals: { sheet: 'графики', r1: 3 + n, c1: 1, r2: 3 + n, c2: K } });
    sh.addChart({ type: 'col', title: 'Выручка и прибыль по кварталам', cats, series: [sr(0), sr(1), sr(2)], anchor: { c1: 0, r1: 12, c2: 6, r2: 30 }, yFmt: '#,##0' });
    sh.addChart({ type: 'col', title: 'Чистые активы и уставный капитал', cats, series: [sr(3), sr(4)], anchor: { c1: 6, r1: 12, c2: 13, r2: 30 }, yFmt: '#,##0' });
    sh.addChart({ type: 'line', title: 'Валюта баланса и собственный капитал', cats, series: [sr(5), sr(6)], anchor: { c1: 0, r1: 31, c2: 6, r2: 49 }, yFmt: '#,##0' });
    sh.addChart({ type: 'line', title: 'Баллы по финансовым показателям', cats, series: [sr(7)], anchor: { c1: 6, r1: 31, c2: 13, r2: 49 } });
  };

  E.buildChecks = function (ctx) {
    const sh = ctx.wb.addSheet('Проверки', { tab: 'FF7F7F7F' }); sh.set(0, 0, 'Проверки достоверности и замечания к данным', ST.title);
    sh.width(0, 22); sh.width(1, 16); sh.width(2, 70); sh.width(3, 14); sh.width(4, 14);
    ['Лист / раздел', 'Период', 'Замечание', 'Значение 1', 'Значение 2'].forEach((t, c) => sh.set(2, c, t, c === 2 ? ST.headL : ST.head));
    let r = 3; const row = (a, b, c, d, e, st) => { sh.set(r, 0, a, ST.name); sh.set(r, 1, b, ST.code); sh.set(r, 2, c, st || ST.name); if (d !== undefined) sh.set(r, 3, d, ST.calc); if (e !== undefined) sh.set(r, 4, e, ST.calc); r++; };
    sh.set(r, 0, 'Арифметика баланса: 1600 (актив) = 1700 (пассив)', ST.sec); for (let c = 1; c < 5; c++) sh.set(r, c, null, ST.sec); r++;
    for (let i = 0; i < ctx.np; i++) {
      const a = ctx.ref('1600', i), b = ctx.ref('1700', i);
      sh.set(r, 0, 'Баланс', ST.name); sh.set(r, 1, ctx.periods[i].label, ST.code);
      sh.set(r, 2, { f: `IF(ABS(${a}-${b})<=MAX(1,0.001*ABS(${a})),"ОК","РАСХОЖДЕНИЕ: актив и пассив не равны")` }, ST.name);
      sh.set(r, 3, { f: a }, ST.calc); sh.set(r, 4, { f: b }, ST.calc); r++;
    }
    sh.cf('C4:C' + r, { type: 'expression', formula: ['LEFT(C4,3)="РАС"'], dxf: { color: 'FF9C0006', fill: 'FFFFC7CE' } });
    r++; sh.set(r, 0, 'Замечания при сборке', ST.sec); for (let c = 1; c < 5; c++) sh.set(r, c, null, ST.sec); r++;
    const need = ['1200', '1300', '1500', '1600', '2110', '2200', '2300'];
    ctx.periods.forEach((p, i) => { const miss = need.filter(c => !E.resolve(ctx, c, i).has); if (miss.length) row('Данные', p.label, 'Нет данных по строкам: ' + miss.join(', ') + ' — показатели периода неполные', undefined, undefined, { fill: ST.YEL, wrap: 1 }); });
    const seen = new Set();
    for (const f of ctx.flags) { const t = f.sheet + '|' + f.msg.replace(/\d+/g, '#'); if (f.level === 'info' && seen.has(t + f.code)) continue; seen.add(t + f.code); row(f.sheet + ' · ' + f.code, ctx.periods[f.i].label, f.msg, undefined, undefined, f.level === 'warn' ? { fill: ST.ORG, wrap: 1 } : ST.name); }
    (ctx.data.notes || []).forEach(n => row(n.where || 'Импорт', n.period || '', n.text, undefined, undefined, n.level === 'warn' ? { fill: ST.ORG, wrap: 1 } : ST.name));
    if (!ctx.np) row('Данные', '', 'Отчетность в документах не найдена — заполните листы «Актив», «Пассив», «Приб_Уб» вручную.', undefined, undefined, { fill: ST.YEL });
    row('Допущения', '', 'Расшифровки 5445, 5540, 5320 (залоги, просроченная ДЗ) в документах не найдены — приняты равными 0, Кл может быть завышен.', undefined, undefined, { fill: ST.ORG, wrap: 1 });
    sh.freeze(3, 0);
  };

  E.buildSources = function (ctx) {
    const sh = ctx.wb.addSheet('Источники', { tab: 'FF7F7F7F' }); sh.set(0, 0, 'Откуда взято каждое число', ST.title);
    ['Лист', 'Ячейка', 'Код', 'Показатель', 'Период', 'Значение, тыс. руб.', 'Источник (файл · лист/страница · строка)'].forEach((t, c) => sh.set(2, c, t, ST.head));
    [14, 9, 8, 44, 14, 16, 90].forEach((w, c) => sh.width(c, w)); let r = 3;
    for (const s of ctx.srcs) {
      sh.set(r, 0, s.sheet, ST.name); sh.set(r, 1, XW.colName(s.c) + (s.r + 1), ST.code); sh.link(r, 1, q(s.sheet) + '!' + XW.colName(s.c) + (s.r + 1), 'Перейти к ячейке'); sh.set(r, 2, s.code, ST.code);
      sh.set(r, 3, M.CODE_NAME[s.code] || '', ST.name); sh.set(r, 4, M.periodFromKey(s.key).label, ST.code); sh.set(r, 5, s.value, ST.calc); sh.set(r, 6, s.src || '', ST.note); r++;
    }
    sh.freeze(3, 0); sh.autofilter = 'A3:G' + Math.max(r, 4);
  };

  E.buildDocs = function (ctx) {
    const sh = ctx.wb.addSheet('Документы', { tab: 'FF7F7F7F' }); sh.set(0, 0, 'Реестр документов заемщика', ST.title);
    ['Файл', 'Раздел', 'Использовано для', 'Статус чтения'].forEach((t, c) => sh.set(2, c, t, ST.headL)); [56, 28, 40, 60].forEach((w, c) => sh.width(c, w)); let r = 3;
    for (const d of ctx.data.docs || []) { sh.set(r, 0, d.name, ST.name); sh.set(r, 1, d.category || '', ST.name); sh.set(r, 2, d.used || '', ST.name); sh.set(r, 3, d.status || '', d.manual ? { fill: ST.YEL, wrap: 1 } : ST.name); r++; }
    sh.freeze(3, 0); sh.autofilter = 'A3:D' + Math.max(r, 4);
  };

  E.buildSummary = function (ctx) {
    const sh = ctx.wb.addSheet('Сводка', { tab: 'FF102A24', grid: false }); const last = ctx.np - 1; const b = ctx.data.borrower;
    sh.width(0, 3); sh.width(1, 46); sh.width(2, 18); sh.width(3, 14); sh.width(4, 14); sh.width(5, 40);
    sh.set(0, 1, 'АНАЛИЗ ФИНАНСОВОГО ПОЛОЖЕНИЯ ЗАЕМЩИКА — ЮЛ', { b: 1, sz: 16, color: ST.G }); sh.set(1, 1, 'Методика ЮЛ 2025 · изменения № 1 от 01.06.2026 · предварительный расчет для аналитика', ST.sub);
    const kv = [['Заемщик', b.name || '— не определено —'], ['ИНН', b.inn || '—'], ['ОКВЭД', b.okved || '—'], ['Отрасль для порогов', { f: "'Коэффициенты'!C2" }], ['Последний период', last >= 0 ? ctx.periods[last].label : '—'], ['Периодов в анализе', ctx.np]];
    kv.forEach((x, k) => { sh.set(3 + k, 1, x[0], ST.nameB); sh.set(3 + k, 2, x[1], { b: 1 }); sh.merge(3 + k, 2, 3 + k, 4); });
    let r = 10; sh.set(r, 1, 'ИТОГ ПО ПОСЛЕДНЕМУ ПЕРИОДУ', ST.sec); for (let c = 2; c < 6; c++) sh.set(r, c, null, ST.sec); r++;
    const fp = k => (last >= 0 ? E.R(ctx, 'Фин показ', k, last) : '""');
    const kp = [['Баллы по фин. показателям (из 90)', fp('fin')], ['Балл бизнес-риска (из 10)', "'Нефин показат'!D20"], ['ИТОГО БАЛЛОВ (из 100)', fp('tot')], ['Класс по сумме баллов', fp('cls')], ['Класс до оценки бизнес-риска (лучший / худший)', `IF(ISNUMBER(${fp('clsMin')}),${fp('clsMin')}&" / "&${fp('clsMax')},"")`], ['Негативных явлений', "'Фин состояние'!E18"], ['Финансовое положение (п. 7.4)', "'Фин состояние'!E22"], ['Чистые активы, тыс. руб.', last >= 0 ? E.R(ctx, 'Чистые активы', 'na', last) : '""']];
    kp.forEach((x, k) => { sh.set(r, 1, x[0], ST.nameB); sh.set(r, 2, { f: x[1] }, { b: 1, sz: 12, h: 'center', bd: 1, fill: ST.GREY, nf: '#,##0' }); sh.merge(r, 2, r, 3); r++; });
    sh.cf('C14', { type: 'cellIs', op: 'equal', formula: ['4'], dxf: { color: 'FF9C0006', fill: 'FFFFC7CE' } }); sh.cf('C14', { type: 'cellIs', op: 'equal', formula: ['1'], dxf: { color: 'FF006100', fill: 'FFC6EFCE' } });
    r++; sh.set(r, 1, 'ПОКАЗАТЕЛИ ПО ПОСЛЕДНЕМУ ПЕРИОДУ', ST.sec); for (let c = 2; c < 6; c++) sh.set(r, c, null, ST.sec); r++;
    ['Показатель', 'Значение', 'Балл', 'Макс.'].forEach((t, c) => sh.set(r, 1 + c, t, ST.head)); r++;
    ctx.sc.defs.forEach(d => { sh.set(r, 1, d[1], ST.name); sh.set(r, 2, { f: last >= 0 ? E.R(ctx, 'Фин показ', 'v_' + d[0], last) : '""' }, ST.num2); sh.set(r, 3, { f: last >= 0 ? E.R(ctx, 'Фин показ', 'p_' + d[0], last) : '""' }, ST.calc); sh.set(r, 4, d[2], ST.code); r++; });
    sh.set(r, 1, 'Чистые активы', ST.name); sh.set(r, 2, { f: last >= 0 ? E.R(ctx, 'Чистые активы', 'na', last) : '""' }, ST.calc); sh.set(r, 3, { f: last >= 0 ? E.R(ctx, 'Фин показ', 'p_na', last) : '""' }, ST.calc); sh.set(r, 4, 10, ST.code); r += 2;
    sh.set(r, 1, 'ЧТО ЗАПОЛНИТЬ ВРУЧНУЮ', ST.sec); for (let c = 2; c < 6; c++) sh.set(r, c, null, ST.sec); r++;
    sh.set(r, 1, 'Факторов бизнес-риска не заполнено (из 12)', ST.name); sh.set(r, 2, { f: "12-'Нефин показат'!D19" }, { b: 1, h: 'center' }); sh.link(r, 1, "'Нефин показат'!D5", 'Перейти'); r++;
    sh.set(r, 1, 'Негативных явлений не оценено (из 12)', ST.name); sh.set(r, 2, { f: "12-'Фин состояние'!E19" }, { b: 1, h: 'center' }); sh.link(r, 1, "'Фин состояние'!E5", 'Перейти'); r++;
    sh.set(r, 1, 'Признак банкротства (п. 7.4) указан', ST.name); sh.set(r, 2, { f: "IF(ISNUMBER('Фин состояние'!E20),\"да\",\"нет — заполните\")" }, { b: 1, h: 'center' }); r += 2;
    sh.set(r, 1, 'СОДЕРЖАНИЕ КНИГИ (нажмите на название)', ST.sec); for (let c = 2; c < 6; c++) sh.set(r, c, null, ST.sec); r++;
    const toc = [['Актив', 'Форма 1, актив'], ['Пассив', 'Форма 1, пассив + расшифровки 5445/5540/5320'], ['Приб_Уб', 'Форма 2 + квартальный блок'], ['Удельный вес в ВБ', 'Структура баланса'], ['Изм.стр. Ф1', 'Изменения баланса г/г'], ['Изм.стр. Ф2', 'Изменения Ф2 г/г'], ['Чистые активы', 'Прил. 4.4.3'], ['Коэффициенты', 'Прил. 4.4.4, выбор отрасли'], ['Фин показ', 'Баллы, итог, класс'], ['Нефин показат', '12 факторов бизнес-риска'], ['Фин состояние', 'Негативные явления и матрица п. 7.4'], ['графики', 'Диаграммы'], ['СПРАВОЧНИК', 'Пороги и шкалы'], ['Проверки', 'Достоверность данных'], ['Источники', 'Происхождение чисел'], ['Документы', 'Реестр файлов']];
    toc.forEach(t => { sh.set(r, 1, t[0], { u: 1, color: 'FF0563C1' }); sh.link(r, 1, q(t[0]) + '!A1', t[1]); sh.set(r, 2, t[1], ST.sub); sh.merge(r, 2, r, 5); r++; });
    r++; sh.set(r, 1, 'Цвета: синий — данные из документов; оранжевый — автопредположение, проверьте; жёлтый — нужен ввод аналитика. Расчет — формулами Excel, можно менять исходные ячейки.', ST.note); sh.merge(r, 1, r, 5); sh.height(r, 30);
  };

  const ORDER = ['Сводка', 'Актив', 'Пассив', 'Приб_Уб', 'Удельный вес в ВБ', 'Изм.стр. Ф1', 'Изм.стр. Ф2', 'Чистые активы', 'Коэффициенты', 'Фин показ', 'Нефин показат', 'Фин состояние', 'графики', 'СПРАВОЧНИК', 'Проверки', 'Источники', 'Документы'];
  /** Собрать книгу по данным. data: {borrower,periods,vals,factors,negatives,bankrupt,docs,notes} */
  E.build = function (data) {
    const ctx = E.makeCtx(data);
    E.buildPrimary(ctx); E.buildShares(ctx); E.buildReference(ctx); E.buildNetAssets(ctx); E.buildCoefficients(ctx); E.buildScores(ctx); E.buildNonFin(ctx); E.buildNegatives(ctx);
    E.buildCharts(ctx); E.buildChecks(ctx); E.buildSources(ctx); E.buildDocs(ctx); E.buildSummary(ctx);
    ctx.wb.sheets.sort((a, b) => ORDER.indexOf(a.name) - ORDER.indexOf(b.name)); ctx.wb.invalidate();
    return ctx;
  };
})(typeof window !== 'undefined' ? window : globalThis);
