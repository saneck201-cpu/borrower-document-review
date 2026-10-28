/* export2.js — часть 2 сборки книги: баллы, бизнес-риск, негативные явления, сводка, графики, проверки, источники, документы. */
(function (root) {
  'use strict';
  const AZ = root.AZ = root.AZ || {}; const E = AZ.exp; const M = AZ.model, XW = root.XW;
  const { SH, ST, NF, C, FC, HR, R0 } = E;
  const q = XW.quoteSheet, ad = XW.addr;

  /* ── Нефинансовые показатели (бизнес-риск) ── */
  function nonFin(X) {
    const sh = X.sheets[SH.N]; const S = X.S;
    sh.set(0, 0, 'Нефинансовые показатели — бизнес-риск (п. 5.3 методики)', ST.title); sh.set(1, 0, 'Заемщик: ' + (S.borrower.name || '— не определен —'), ST.sub);
    ['№', 'Показатель (п. методики)', 'Макс.', 'Балл (выберите)', 'Основание / источник', 'Варианты оценки по методике'].forEach((t, c) => sh.set(HR, c, t, ST.hdr)); sh.height(HR, 30);
    [5, 38, 8, 14, 58, 90].forEach((w, c) => sh.width(c, w)); sh.freeze(HR + 1, 2);
    M.FACTORS.forEach((f, k) => {
      const r = R0 + k; const v = S.factors && S.factors[k] != null ? S.factors[k] : null;
      sh.set(r, 0, k + 1, ST.code); sh.set(r, 1, f.n + ' (п. ' + f.p + ')', ST.lab); sh.set(r, 2, f.max, ST.code);
      sh.set(r, 3, v, { b: 1, color: C.input, h: 'center', bd: { all: 'thin' }, fill: v == null ? C.yellow : (S.factorAuto && S.factorAuto[k] ? C.orange : undefined) });
      sh.set(r, 4, (S.factorNotes && S.factorNotes[k]) || '', ST.lab);
      sh.set(r, 5, f.opts.map(o => o[0] + ' — ' + o[1]).join('; '), ST.lab); sh.height(r, 32);
      sh.validate(ad(r, 3), { list: f.allowed.map(String), title: 'Балл: ' + f.n.slice(0, 22), prompt: f.opts.map(o => o[0] + ' — ' + o[1]).join('\n').slice(0, 250), error: 'Допустимые баллы: ' + f.allowed.join(', ') });
    });
    const r1 = R0 + 12; const rng = ad(R0, 3) + ':' + ad(R0 + 11, 3);
    sh.set(r1, 1, 'Сумма баллов (макс. 31)', ST.labB); sh.set(r1, 3, { f: 'SUM(' + rng + ')' }, ST.calc('0', 1, C.sect));
    sh.set(r1 + 1, 1, 'Заполнено факторов из 12', ST.lab); sh.set(r1 + 1, 3, { f: 'COUNT(' + rng + ')' }, ST.calc('0'));
    sh.set(r1 + 2, 1, 'БАЛЛ БИЗНЕС-РИСКА (0 / 5 / 10): сумма <15 → 0; 15…20 → 5; >20 → 10', ST.labB);
    const thr = X.K.thr.risk;
    sh.set(r1 + 2, 3, { f: 'IF(' + ad(r1 + 1, 3) + '<12,"",IF(' + ad(r1, 3) + '<' + thr.lo + ',0,IF(' + ad(r1, 3) + '<=' + thr.hi + ',5,10)))' }, ST.calc('0', 1, C.sect));
    sh.set(r1 + 3, 1, { f: 'IF(' + ad(r1 + 1, 3) + '<12,"Балл бизнес-риска не присваивается, пока не заполнены все 12 факторов — итоговый класс считается без него.","")' }, { color: 'FFB3261E', b: 1 });
    sh.set(r1 + 5, 1, 'Оранжевый балл подставлен автоматически (по реквизитам из документов) — проверьте. Желтая ячейка — нужна оценка аналитика: по документам эти факторы определить нельзя.', ST.note); sh.merge(r1 + 5, 1, r1 + 5, 4); sh.height(r1 + 5, 32);
    X.K.risk = q(SH.N) + '!$D$' + (r1 + 3); X.K.riskSum = q(SH.N) + '!$D$' + (r1 + 1); X.K.riskCnt = q(SH.N) + '!$D$' + (r1 + 2);
  }

  /* ── Фин показ: баллы по периодам ── */
  function scoring(X) {
    const sh = X.sheets[SH.S]; E.sheetHead(X, sh, 'Балльная оценка финансовых показателей и класс кредитоспособности (п. 5.4)', { c1: 'Макс.', c2: '', w0: 54 });
    const k = X.K.k; const kr = (key, i) => X.ref(SH.K, k[key], i); const T = X.K.thr; const ind = X.K.indCell, tbl = X.K.ind;
    const vl = col => 'VLOOKUP(' + ind + ',' + tbl + ',' + col + ',FALSE)';
    const R = {}; let r = R0;
    const line = (key, label, max, fn, o) => { o = o || {}; R[key] = r; sh.set(r, 0, label, o.b ? ST.labB : ST.lab); sh.set(r, 1, max, ST.code); sh.set(r, 2, '', ST.lab); X.periods.forEach((kk, i) => sh.set(r, FC + i, { f: fn(i) }, ST.calc(o.nf || '0', o.b, o.b ? C.sect : undefined))); r++; };
    const asc = (key, thr, s) => i => { const x = kr(key, i); return 'IF(ISNUMBER(' + x + '),IF(' + x + '<' + thr.lo + ',' + s[0] + ',IF(' + x + '<=' + thr.hi + ',' + s[1] + ',' + s[2] + ')),"")'; };
    line('na', '1. Чистые активы (> 0 → 10)', 10, i => X.ref(SH.NA, X.K.na.sc, i));
    line('ka', '2. Кфн — автономия', 5, asc('ka', T.ka, [0, 3, 5]));
    line('ko', '3. Ко — обеспеченность СОС', 5, asc('ko', T.ko, [0, 3, 5]));
    line('kl', '4. Кл — текущая ликвидность', 10, asc('kl', T.kl, [0, 5, 10]));
    line('kpl', '5. Кпл — степень платежеспособности (мес.)', 10, i => { const x = kr('kpl', i); return 'IF(ISNUMBER(' + x + '),IF(' + x + '>' + T.kpl.hi + ',0,IF(' + x + '>=' + T.kpl.lo + ',5,10)),"")'; });
    line('kob', '6. Коб — оборачиваемость оборотных средств', 10, asc('kob', T.kob, [0, 5, 10]));
    line('kobdz', '7. Кобдз — оборачиваемость дебиторской задолженности', 10, asc('kobdz', T.kobdz, [0, 5, 10]));
    line('rpr', '8. Рпр — рентабельность продаж (пороги по отрасли)', 10, i => { const x = kr('rpr', i); return 'IF(ISNUMBER(' + x + '),IF(' + x + '<' + vl(2) + ',0,IF(' + x + '>' + vl(5) + ',10,5)),"")'; });
    line('ra', '9. Ра — рентабельность активов (пороги по отрасли)', 10, i => { const x = kr('ra', i); return 'IF(ISNUMBER(' + x + '),IF(' + x + '<' + vl(2) + ',0,IF(' + x + '>' + vl(4) + ',10,5)),"")'; });
    line('rsk', '10. Рск — рентабельность собственного капитала', 10, asc('rsk', T.rsk, [0, 5, 10]));
    const a0 = R.na, a1 = R.rsk;
    line('fin', 'Сумма финансовых баллов (макс. 90)', 90, i => 'SUM(' + ad(a0, FC + i) + ':' + ad(a1, FC + i) + ')', { b: 1 });
    line('cnt', 'Показателей рассчитано из 10', 10, i => 'COUNT(' + ad(a0, FC + i) + ':' + ad(a1, FC + i) + ')');
    line('risk', '11. Бизнес-риск (0 / 5 / 10) — лист «Нефин показат»', 10, i => 'IF(ISNUMBER(' + X.K.risk + '),' + X.K.risk + ',"")');
    line('tot', 'ИТОГО БАЛЛОВ (макс. 100)', 100, i => ad(R.fin, FC + i) + '+IF(ISNUMBER(' + ad(R.risk, FC + i) + '),' + ad(R.risk, FC + i) + ',0)', { b: 1 });
    const cls = x => 'IF(' + x + '>=80,1,IF(' + x + '>=50,2,IF(' + x + '>=30,3,4)))';
    line('cls', 'КЛАСС КРЕДИТОСПОСОБНОСТИ (1 — лучший, 4 — худший)', 4, i => cls(ad(R.tot, FC + i)), { b: 1 });
    line('clsr', 'Диапазон класса, пока бизнес-риск не оценен (0…10 баллов)', '', i => 'IF(ISNUMBER(' + ad(R.risk, FC + i) + '),' + ad(R.cls, FC + i) + '&"","")&IF(ISNUMBER(' + ad(R.risk, FC + i) + '),"",' + cls(ad(R.fin, FC + i) + '+10') + '&"–"&' + cls(ad(R.fin, FC + i)) + ')', { nf: '@' });
    sh.cf(ad(R.cls, FC) + ':' + ad(R.cls, FC + Math.max(0, X.n - 1)), { type: 'cellIs', op: 'greaterThanOrEqual', formula: ['3'], dxf: { color: 'FF9C0006', fill: 'FFFAD4D4' } });
    sh.cf(ad(R.cls, FC) + ':' + ad(R.cls, FC + Math.max(0, X.n - 1)), { type: 'cellIs', op: 'lessThanOrEqual', formula: ['2'], dxf: { color: 'FF1B5E20', fill: 'FFD6EFDC' } });
    sh.set(r + 1, 0, 'Рейтинговый класс — расчетный показатель. Бизнес-риск единый для всех периодов (оценка на дату анализа), поэтому по прошлым периодам класс ориентировочный. Окончательное положение — на листе «Фин состояние» с учетом негативных явлений.', ST.note); sh.merge(r + 1, 0, r + 1, 8); sh.height(r + 1, 44);
    X.K.s = R;
  }

  /* ── Негативные явления и финансовое положение (п. 7.2, 7.4) ── */
  function state(X) {
    const sh = X.sheets[SH.F]; const S = X.S; E.sheetHead(X, sh, 'Финансовое положение: негативные явления (п. 7.2) и матрица (п. 7.4)', { c0: 'Негативное явление', c1: 'Пункт', c2: 'Ввод аналитика', w0: 70, w2: 16 });
    const pr = (code, i) => X.ref(SH.P, X.rows.P[code], i); const bl = (kind, code, i) => X.bal(kind, code, i); const na = (key, i) => X.ref(SH.NA, X.K.na[key], i);
    const last = X.n - 1; const rowOf = {}; let r = R0; const manualRows = [];
    M.NEGATIVES.forEach((n, k) => {
      rowOf[n.p] = r; sh.set(r, 0, n.t, ST.lab); sh.set(r, 1, n.p, ST.code); sh.height(r, 44);
      const man = !n.auto || n.auto === 'zero'; const manualOnly = !n.auto;
      if (manualOnly) { const v = S.negManual && S.negManual[k]; sh.set(r, 2, v || null, { b: 1, color: C.input, h: 'center', bd: { all: 'thin' }, fill: v ? undefined : C.yellow }); sh.validate(ad(r, 2), { list: ['Да', 'Нет'], title: n.p, prompt: 'Да — явление выявлено на дату анализа (по документам банка/заемщика).' }); manualRows.push(r); }
      else sh.set(r, 2, 'авто', ST.code);
      X.periods.forEach((key, i) => {
        const c = FC + i; const j = X.prevYear(i); const grey = { fill: C.grey, bd: { all: 'thin' } }; let f = null;
        switch (n.auto) {
          case 'na': f = 'IF(ISNUMBER(' + na('na', i) + '),IF(' + na('na', i) + '<0,1,0),"")'; break;
          case 'nadrop': f = 'IF(ISNUMBER(' + na('drop', i) + '),IF(' + na('drop', i) + '>=0.25,1,0),"")'; break;
          case 'zero': f = 'IF(AND(ISNUMBER(' + bl('A', '1200', i) + '),ISNUMBER(' + bl('L', '1500', i) + '),' + bl('A', '1600', i) + '>0),IF(AND(' + bl('A', '1200', i) + '=0,' + bl('L', '1500', i) + '=0),1,0),"")'; break;
          case 'rev': if (j >= 0) f = 'IF(' + pr('2110', j) + '<=0,"",IF(' + pr('2110', i) + '/' + pr('2110', j) + '-1<-0.4,1,0))'; break;
          case 'margin': if (j >= 0) { const mc = pr('2200', i) + '/' + pr('2110', i), mp = pr('2200', j) + '/' + pr('2110', j); f = 'IF(OR(' + pr('2110', i) + '<=0,' + pr('2110', j) + '<=0),"",IF(' + mp + '<=0,"",IF((' + mc + '-' + mp + ')/' + mp + '<-0.4,1,0)))'; } break;
          case 'ap': if (j >= 0) f = 'IF(OR(' + bl('L', '1520', j) + '<=0,' + pr('2120', j) + '=0),"",IF(AND(' + bl('L', '1520', i) + '/' + bl('L', '1520', j) + '-1>0.4,' + pr('2120', i) + '/' + pr('2120', j) + '-1<0.4),1,0))'; break;
          case 'ar': if (j >= 0) f = 'IF(OR(' + bl('A', '1230', j) + '<=0,' + pr('2110', j) + '<=0),"",IF(AND(' + bl('A', '1230', i) + '/' + bl('A', '1230', j) + '-1>0.4,' + pr('2110', i) + '/' + pr('2110', j) + '-1<0.4),1,0))'; break;
          default: if (i === last) f = 'IF(' + ad(r, 2, false, true) + '="Да",1,IF(' + ad(r, 2, false, true) + '="Нет",0,""))';
        }
        if (f) sh.set(r, c, { f }, { h: 'center', b: 1, bd: { all: 'thin' }, nf: '0' }); else sh.set(r, c, null, grey);
      }); r++;
    });
    const rng = i => ad(R0, FC + i) + ':' + ad(R0 + 11, FC + i);
    const mrng = ad(manualRows[0], 2) + ':' + ad(manualRows[manualRows.length - 1], 2);
    const res = {}; const line = (key, label, fn, o) => { o = o || {}; res[key] = r; sh.set(r, 0, label, o.b ? ST.labB : ST.lab); sh.set(r, 1, '', ST.code); sh.set(r, 2, o.c2 == null ? '' : o.c2, o.c2style || ST.lab); X.periods.forEach((k2, i) => sh.set(r, FC + i, { f: fn(i) }, { h: 'center', b: o.b, bd: { all: 'thin' }, nf: o.nf || '0', fill: o.b ? C.sect : undefined, wrap: 1 })); r++; };
    r++;
    line('n', 'Количество выявленных негативных явлений', i => 'COUNTIF(' + rng(i) + ',1)', { b: 1 });
    line('open', 'Не оценено аналитиком (пункты 7.2.6–7.2.10; считается только на дату анализа)', i => i === last ? String(manualRows.length) + '-COUNTIF(' + mrng + ',"Да")-COUNTIF(' + mrng + ',"Нет")' : '""');
    line('bk', 'Признан банкротом / устойчиво неплатежеспособен (п. 7.3)', i => i === last ? ad(r, 2, true, true) + '="Да"' : '""', { c2: S.bankrupt || 'Нет', c2style: { b: 1, color: C.input, h: 'center', bd: { all: 'thin' }, fill: C.yellow }, nf: 'General' });
    sh.validate(ad(res.bk, 2), { list: ['Да', 'Нет'], title: 'Банкротство', prompt: 'Да — если признан банкротом или устойчиво неплатежеспособен более 180 дней.' });
    // функции матрицы
    line('cls', 'Класс кредитоспособности (лист «Фин показ»)', i => X.ref(SH.S, X.K.s.cls, i), { b: 1 });
    const n = i => ad(res.n, FC + i), cl = i => ad(res.cls, FC + i), bk = i => ad(res.bk, FC + i);
    line('pos', 'Финансовое положение по матрице п. 7.4', i => 'IF(' + cl(i) + '="","",IF(' + (i === last ? bk(i) : 'FALSE') + ',"Плохое",IF(' + cl(i) + '=1,IF(' + n(i) + '=0,"Хорошее","Среднее"),IF(' + cl(i) + '=2,IF(' + n(i) + '<=1,"Хорошее",IF(' + n(i) + '<=3,"Среднее","Плохое")),IF(' + cl(i) + '=3,IF(' + n(i) + '<=2,"Среднее","Плохое"),IF(' + n(i) + '<=1,"Среднее","Плохое"))))))', { b: 1, nf: '@' });
    const ps = i => ad(res.pos, FC + i);
    line('fin', 'Итог (с учетом п. 7.2: при любом негативном явлении положение не «хорошее»)', i => 'IF(' + ps(i) + '="","",IF(AND(' + ps(i) + '="Хорошее",' + n(i) + '>=1),"Среднее",' + ps(i) + '))', { b: 1, nf: '@' });
    line('warn', 'Предупреждение', i => 'IF(AND(' + ps(i) + '="Хорошее",' + n(i) + '>=1),"Матрица допускает 1 явление для 2 класса, п. 7.2 — нет: принято «Среднее»","")', { nf: '@' });
    for (const k of ['pos', 'fin']) sh.cf(ad(res[k], FC) + ':' + ad(res[k], FC + Math.max(0, X.n - 1)), { type: 'cellIs', op: 'equal', formula: ['"Плохое"'], dxf: { color: 'FF9C0006', fill: 'FFFAD4D4' } });
    sh.cf(ad(R0, FC) + ':' + ad(R0 + 11, FC + Math.max(0, X.n - 1)), { type: 'cellIs', op: 'equal', formula: ['1'], dxf: { color: 'FF9C0006', fill: 'FFFAD4D4' } });
    sh.width(2, 16); sh.set(r + 1, 0, 'Пункты 7.2.1–7.2.5, 7.2.11, 7.2.12 считаются автоматически (сравнение — с тем же периодом прошлого года; сезонность и бизнес-план не учитываются — проверьте). 7.2.6–7.2.10 требуют данных банка: выберите «Да»/«Нет» в колонке C.', ST.note); sh.merge(r + 1, 0, r + 1, 6); sh.height(r + 1, 44);
    X.K.f = res;
  }

  /* ── Графики ── */
  function charts(X) {
    const sh = X.sheets[SH.G]; sh.set(0, 0, 'Данные для графиков и диаграммы', ST.title); sh.width(0, 40);
    X.periods.forEach((k, i) => { sh.set(2, 1 + i, X.lab(k), ST.hdr); sh.width(1 + i, 12); });
    const rows = [['Валюта баланса', i => X.bal('A', '1600', i)], ['Чистые активы', i => X.ref(SH.NA, X.K.na.na, i)], ['Выручка (нарастающим итогом)', i => X.ref(SH.P, X.rows.P['2110'], i)], ['Прибыль от продаж', i => X.ref(SH.P, X.rows.P['2200'], i)], ['Чистая прибыль', i => X.ref(SH.P, X.rows.P['2400'], i)], ['Итого баллов', i => X.ref(SH.S, X.K.s.tot, i)], ['Кл — текущая ликвидность', i => X.ref(SH.K, X.K.k.kl, i)], ['Кфн — автономия', i => X.ref(SH.K, X.K.k.ka, i)]];
    rows.forEach(([t, fn], k) => { sh.set(3 + k, 0, t, ST.lab); X.periods.forEach((kk, i) => sh.set(3 + k, 1 + i, { f: fn(i) }, ST.calc(k >= 6 ? NF.k2 : NF.num))); });
    const rg = (r) => ({ sheet: SH.G, r1: r, c1: 1, r2: r, c2: X.n });
    const cats = rg(2); const ser = (r, name) => ({ name, vals: rg(r) });
    if (X.n < 2) { sh.set(13, 0, 'Для графиков нужно не менее двух периодов.', ST.note); return; }
    sh.addChart({ type: 'col', title: 'Баланс и чистые активы, тыс. руб.', cats, series: [ser(3, 'Валюта баланса'), ser(4, 'Чистые активы')], anchor: { c1: 0, r1: 12, c2: 8, r2: 30 } });
    sh.addChart({ type: 'col', title: 'Выручка и прибыль (нарастающим итогом), тыс. руб.', cats, series: [ser(5, 'Выручка'), ser(6, 'Прибыль от продаж'), ser(7, 'Чистая прибыль')], anchor: { c1: 9, r1: 12, c2: 17, r2: 30 } });
    sh.addChart({ type: 'line', title: 'Итого баллов (макс. 100)', cats, series: [ser(8, 'Баллы')], anchor: { c1: 0, r1: 31, c2: 8, r2: 48 } });
    sh.addChart({ type: 'line', title: 'Кл и Кфн', cats, series: [ser(9, 'Кл'), ser(10, 'Кфн')], anchor: { c1: 9, r1: 31, c2: 17, r2: 48 } });
  }

  /* ── Проверки ── */
  function checks(X) {
    const sh = X.sheets[SH.C]; const S = X.S; sh.set(0, 0, 'Проверки достоверности и допущения (п. 6.4 методики)', ST.title);
    sh.width(0, 60); sh.width(1, 14); sh.width(2, 16); sh.width(3, 16); sh.width(4, 70);
    X.periods.forEach((k, i) => { sh.set(2, 1 + i, X.lab(k), ST.hdr); if (i > 2) sh.width(1 + i, 12); }); sh.set(2, 0, 'Арифметический контроль по периодам', ST.hdrL); sh.freeze(3, 1);
    let r = 3; const bad = [];
    const line = (label, fn) => { sh.set(r, 0, label, ST.lab); X.periods.forEach((k, i) => { const f = fn(i); if (f == null) sh.set(r, 1 + i, null, { fill: C.grey, bd: { all: 'thin' } }); else sh.set(r, 1 + i, { f }, { h: 'center', bd: { all: 'thin' }, nf: '0;[Red]-0;0' }); }); bad.push(r); r++; };
    line('Баланс: актив − пассив (1600 − 1700), тыс. руб. — должно быть 0', i => X.bal('A', '1600', i) + '-' + X.bal('L', '1700', i));
    line('Нарастающая выручка меньше, чем за предыдущий квартал того же года (1 — ошибка данных)', i => { const key = X.periods[i]; if (+key.split('-')[1] === 1) return null; const j = X.prevQ(i); if (j < 0) return null; return 'IF(' + X.ref(SH.P, X.rows.P['2110'], i) + '<' + X.ref(SH.P, X.rows.P['2110'], j) + ',1,0)'; });
    line('Отчетность на 31.12 прошлого года отсутствует (1 — среднее считается по концу периода)', i => X.startIdx(i) < 0 ? '1' : '0');
    line('Чистые активы меньше уставного капитала (1 — п. 5.3.12)', i => 'IF(' + X.ref(SH.NA, X.K.na.dif, i) + '<0,1,0)');
    sh.cf(ad(3, 1) + ':' + ad(r, X.n), { type: 'cellIs', op: 'notEqual', formula: ['0'], dxf: { color: 'FF9C0006', fill: 'FFFAD4D4' } });
    r += 1; const sect = t => { sh.set(r, 0, t, ST.labB); sh.merge(r, 0, r, 4); r++; };
    sect('Итоговые строки, не сходящиеся с суммой статей (в книге оставлена цифра из отчетности)');
    if (!X.checks.length) { sh.set(r, 0, 'Расхождений нет', ST.lab); r++; }
    X.checks.forEach(c => { sh.set(r, 0, c.sheet + ', код ' + c.code + ' (' + M.CODE_NAME[c.code] + '), ' + X.lab(c.key), ST.lab); sh.set(r, 1, c.rep, ST.inp()); sh.set(r, 2, c.sum, ST.calc()); sh.set(r, 3, c.rep - c.sum, ST.calc()); sh.set(r, 4, 'в отчетности / сумма статей / разница', ST.note); r++; });
    r++; sect('Допущения и недостающие данные');
    const fl = X.flags; if (!fl.length && !(S.notes || []).length) { sh.set(r, 0, 'Нет', ST.lab); r++; }
    const groups = new Map(); for (const f of fl) { if (!groups.has(f.text)) groups.set(f.text, []); groups.get(f.text).push(f.key); }
    for (const [text, keys] of groups) { sh.set(r, 0, text + ' — периодов: ' + keys.length + (keys.length > 1 ? ' (' + X.lab(keys[0]) + ' … ' + X.lab(keys[keys.length - 1]) + ')' : ' (' + X.lab(keys[0]) + ')'), ST.lab); sh.merge(r, 0, r, 4); sh.height(r, 30); r++; }
    X.K.flagGroups = groups.size;
    for (const n of (S.notes || [])) { sh.set(r, 0, n, ST.lab); sh.merge(r, 0, r, 4); sh.height(r, 30); r++; }
    r++; sect('Расхождения между документами (одна и та же статья, разные значения)');
    if (!(S.conflicts || []).length) { sh.set(r, 0, 'Расхождений нет', ST.lab); r++; }
    for (const c of (S.conflicts || [])) { sh.set(r, 0, X.lab(c.key) + ', ' + c.code + ' ' + (M.CODE_NAME[c.code] || ''), ST.lab); sh.set(r, 1, c.a, ST.inp()); sh.set(r, 2, c.b, ST.inp()); sh.set(r, 3, c.kept === 'a' ? 'принято 1-е' : 'принято 2-е', ST.lab); sh.set(r, 4, c.srcA + '  ⟷  ' + c.srcB, ST.note); r++; }
    X.K.checksSheetRows = bad;
    X.K.issueCount = X.checks.length + (S.conflicts || []).length;
  }

  /* ── Источники ── */
  function sources(X) {
    const sh = X.sheets[SH.SRC]; sh.set(0, 0, 'Источники: откуда взято каждое число', ST.title); ['Период', 'Код', 'Статья', 'Значение, тыс. руб.', 'Файл · лист/страница · строка'].forEach((t, c) => sh.set(2, c, t, ST.hdr));
    [14, 8, 52, 18, 90].forEach((w, c) => sh.width(c, w)); let r = 3;
    for (const key of X.periods) for (const [code, name] of M.CODES) { const v = X.v(key, code); if (v == null) continue; sh.set(r, 0, X.lab(key), ST.lab); sh.set(r, 1, code, ST.code); sh.set(r, 2, name, ST.lab); sh.set(r, 3, v, ST.inp()); sh.set(r, 4, (X.S.src[key] && X.S.src[key][code]) || '', ST.lab); r++; }
    sh.autofilter = 'A3:E' + Math.max(4, r); sh.freeze(3, 0);
  }

  /* ── Документы ── */
  function docs(X) {
    const sh = X.sheets[SH.DOC]; sh.set(0, 0, 'Реестр документов заемщика', ST.title); ['Файл', 'Раздел', 'Как использован', 'Статус чтения', 'Примечание'].forEach((t, c) => sh.set(2, c, t, ST.hdr));
    [48, 40, 46, 22, 60].forEach((w, c) => sh.width(c, w)); let r = 3;
    for (const d of (X.S.docs || [])) { sh.set(r, 0, d.path, ST.lab); sh.set(r, 1, d.kind, ST.lab); sh.set(r, 2, d.usedFor || '', ST.lab); sh.set(r, 3, d.status, { bd: { all: 'thin' }, fill: /не прочит|скан|ошибк/i.test(d.status) ? C.orange : undefined }); sh.set(r, 4, d.note || '', ST.lab); r++; }
    sh.autofilter = 'A3:E' + Math.max(4, r); sh.freeze(3, 0);
  }

  /* ── Сводка ── */
  function summary(X) {
    const sh = X.sheets[SH.SUM]; const S = X.S, b = S.borrower; const last = X.n - 1; const Kk = X.K;
    sh.set(0, 0, 'Анализ финансового положения заемщика', { b: 1, sz: 18, color: 'FF173B31' }); sh.set(1, 0, 'Методика ЮЛ 2025 · Изменения № 1 от 01.06.2026 · расчет выполнен формулами в этой книге', ST.sub);
    sh.width(0, 46); sh.width(1, 18); sh.width(2, 16); sh.width(3, 12); sh.width(4, 18); sh.width(5, 60);
    const info = [['Заемщик', b.name || '— не определен —'], ['ИНН', b.inn || ''], ['ОГРН', b.ogrn || ''], ['ОКВЭД', b.okved || ''], ['Дата регистрации', b.regDate || ''], ['Отрасль для порогов', { f: X.K.indCell }], ['Анализируемый период', X.lab(X.periods[last])], ['Периодов в книге', X.n]];
    info.forEach(([t, v], k) => { sh.set(3 + k, 0, t, ST.labB); sh.set(3 + k, 1, v, { b: 1, bd: { all: 'thin' } }); sh.merge(3 + k, 1, 3 + k, 3); });
    let r = 12; sh.set(r, 0, 'ИТОГ', ST.hdrL); ['Значение', '', '', '', 'Где смотреть'].forEach((t, c) => sh.set(r, 1 + c, t, ST.hdr)); r++;
    const res = [['Итого баллов (из 100)', X.ref(SH.S, Kk.s.tot, last), SH.S], ['Класс кредитоспособности', X.ref(SH.S, Kk.s.cls, last), SH.S], ['Если бизнес-риск не оценен — диапазон класса', X.ref(SH.S, Kk.s.clsr, last), SH.N], ['Негативных явлений выявлено', X.ref(SH.F, Kk.f.n, last), SH.F], ['Не оценено аналитиком (п. 7.2.6–7.2.10)', X.ref(SH.F, Kk.f.open, last), SH.F], ['Финансовое положение (матрица п. 7.4)', X.ref(SH.F, Kk.f.pos, last), SH.F], ['Финансовое положение — итог с учетом п. 7.2', X.ref(SH.F, Kk.f.fin, last), SH.F], ['Чистые активы, тыс. руб.', X.ref(SH.NA, Kk.na.na, last), SH.NA]];
    res.forEach(([t, f, to]) => { sh.set(r, 0, t, ST.lab); sh.set(r, 1, { f }, { b: 1, sz: 12, h: 'center', bd: { all: 'thin' }, nf: '#,##0' }); sh.merge(r, 1, r, 4); sh.set(r, 5, 'Лист «' + to + '»', { u: 1, color: C.link }); sh.link(r, 5, q(to) + '!A1', to); r++; });
    r++; sh.set(r, 0, 'Показатели на ' + X.lab(X.periods[last]), ST.hdrL); ['Значение', 'Балл', 'Макс.', 'Тот же период год назад', 'Норма'].forEach((t, c) => sh.set(r, 1 + c, t, ST.hdr)); sh.height(r, 30); r++;
    const j = X.prevYear(last); const kr = (key, i) => X.ref(SH.K, Kk.k[key], i); const sc = (key, i) => X.ref(SH.S, Kk.s[key], i);
    const kp = [['Чистые активы', X.ref(SH.NA, Kk.na.na, last), 'na', 10, j >= 0 ? X.ref(SH.NA, Kk.na.na, j) : null, '> 0', NF.num], ['Кфн — автономия', kr('ka', last), 'ka', 5, j >= 0 ? kr('ka', j) : null, '> 0,2', NF.k2], ['Ко — обеспеченность СОС', kr('ko', last), 'ko', 5, j >= 0 ? kr('ko', j) : null, '> 0,2', NF.k2], ['Кл — текущая ликвидность', kr('kl', last), 'kl', 10, j >= 0 ? kr('kl', j) : null, '> 1,3', NF.k2], ['Кпл — степень платежеспособности, мес.', kr('kpl', last), 'kpl', 10, j >= 0 ? kr('kpl', j) : null, '< 4', NF.k2], ['Коб — оборачиваемость оборотных средств', kr('kob', last), 'kob', 10, j >= 0 ? kr('kob', j) : null, '> 2', NF.k2], ['Кобдз — оборачиваемость дебиторской задолженности', kr('kobdz', last), 'kobdz', 10, j >= 0 ? kr('kobdz', j) : null, '> 2', NF.k2], ['Рпр — рентабельность продаж, %', kr('rpr', last), 'rpr', 10, j >= 0 ? kr('rpr', j) : null, 'по отрасли', NF.k1], ['Ра — рентабельность активов, %', kr('ra', last), 'ra', 10, j >= 0 ? kr('ra', j) : null, 'по отрасли', NF.k1], ['Рск — рентабельность собственного капитала, %', kr('rsk', last), 'rsk', 10, j >= 0 ? kr('rsk', j) : null, '> 4', NF.k1]];
    const k0 = r; kp.forEach(([t, f, key, mx, fp, norm, nf]) => { sh.set(r, 0, t, ST.lab); sh.set(r, 1, { f }, ST.calc(nf, 1)); sh.set(r, 2, { f: sc(key, last) }, ST.calc('0', 1)); sh.set(r, 3, mx, ST.code); if (fp) sh.set(r, 4, { f: fp }, ST.calc(nf)); else sh.set(r, 4, null, { fill: C.grey, bd: { all: 'thin' } }); sh.set(r, 5, norm, ST.lab); r++; });
    sh.cf(ad(k0, 2) + ':' + ad(r - 1, 2), { type: 'expression', formula: 'AND(ISNUMBER(' + ad(k0, 2) + '),' + ad(k0, 2) + '=0)', dxf: { color: 'FF9C0006', fill: 'FFFAD4D4' } });
    r++; sh.set(r, 0, 'ЧТО ТРЕБУЕТ ВНИМАНИЯ', ST.hdrL); sh.merge(r, 0, r, 5); r++;
    const todo = []; const risk = Kk.riskCnt;
    todo.push({ f: 'IF(' + risk + '<12,"• Бизнес-риск: заполнено факторов "&' + risk + '&" из 12 — лист «Нефин показат»","• Бизнес-риск заполнен полностью")', to: SH.N });
    todo.push({ f: 'IF(' + X.ref(SH.F, Kk.f.open, last) + '>0,"• Негативные явления 7.2.6–7.2.10: не оценено "&' + X.ref(SH.F, Kk.f.open, last) + '&" — лист «Фин состояние», колонка C","• Негативные явления 7.2.6–7.2.10 оценены")', to: SH.F });
    todo.push({ s: '• Проверки достоверности и допущения: ' + Kk.issueCount + ' расхождений в данных, ' + (Kk.flagGroups || 0) + ' допущений — лист «Проверки»', to: SH.C });
    (S.notes || []).forEach(n => todo.push({ s: '• ' + n, to: SH.C }));
    todo.forEach(t => { sh.set(r, 0, t.f ? { f: t.f } : t.s, { wrap: 1 }); sh.merge(r, 0, r, 4); sh.set(r, 5, 'перейти', { u: 1, color: C.link }); sh.link(r, 5, q(t.to) + '!A1', t.to); sh.height(r, 30); r++; });
    r++; sh.set(r, 0, 'ЛИСТЫ КНИГИ', ST.hdrL); sh.merge(r, 0, r, 5); r++;
    const desc = { [SH.A]: 'Бухгалтерский баланс, актив (данные из документов)', [SH.L]: 'Бухгалтерский баланс, пассив', [SH.P]: 'Отчет о финансовых результатах (нарастающим итогом и по кварталам)', [SH.W]: 'Структура баланса, % к валюте', [SH.D1]: 'Изменения баланса к прошлому году', [SH.D2]: 'Изменения финрезультатов к прошлому году', [SH.NA]: 'Расчет чистых активов', [SH.K]: 'Коэффициенты и отрасль', [SH.S]: 'Баллы и класс кредитоспособности', [SH.N]: 'Бизнес-риск: 12 факторов (ввод аналитика)', [SH.F]: 'Негативные явления и финансовое положение', [SH.G]: 'Графики', [SH.R]: 'Пороги методики', [SH.C]: 'Проверки и допущения', [SH.SRC]: 'Откуда взято каждое число', [SH.DOC]: 'Реестр документов' };
    for (const nm of Object.keys(desc)) { sh.set(r, 0, nm, { u: 1, color: C.link, bd: { all: 'thin' } }); sh.link(r, 0, q(nm) + '!A1', nm); sh.set(r, 1, desc[nm], ST.lab); sh.merge(r, 1, r, 5); r++; }
    sh.set(r + 1, 0, 'Синий шрифт — данные из документов, черный — формулы; желтая заливка — нужен ввод аналитика; оранжевая — допущение/автоподстановка; красная — нарушение критерия. Результат предварительный и требует профессионального суждения.', ST.note); sh.merge(r + 1, 0, r + 1, 5); sh.height(r + 1, 44);
  }

  /* ── Сборка ── */
  function build(S) {
    if (!S.periods || !S.periods.length) throw new Error('Нет ни одного периода отчетности');
    const X = E.newCtx(S);
    const order = [[SH.SUM, 'FF173B31'], [SH.A, 'FF2E7D5B'], [SH.L, 'FF2E7D5B'], [SH.P, 'FF2E7D5B'], [SH.W, 'FF7CB59A'], [SH.D1, 'FF7CB59A'], [SH.D2, 'FF7CB59A'], [SH.NA, 'FF1F6FB2'], [SH.K, 'FF1F6FB2'], [SH.S, 'FF1F6FB2'], [SH.N, 'FFE0A800'], [SH.F, 'FFE0A800'], [SH.G, 'FF7CB59A'], [SH.R, 'FF8A8F94'], [SH.C, 'FFB3261E'], [SH.SRC, 'FF8A8F94'], [SH.DOC, 'FF8A8F94']];
    for (const [n, tab] of order) X.sheet(n, { tab, grid: [SH.SUM].includes(n) ? false : true, zoom: n === SH.SUM ? 100 : 90 });
    E.reference(X);
    E.inputSheet(X, 'A', 'Бухгалтерский баланс — АКТИВ'); E.inputSheet(X, 'L', 'Бухгалтерский баланс — ПАССИВ'); E.inputSheet(X, 'P', 'Отчет о финансовых результатах');
    E.quarterBlock(X); E.structureSheet(X);
    E.dynSheet(X, SH.D1, 'Изменения баланса к аналогичному периоду прошлого года', [['A', M.CODES.filter(c => c[2] === 'A').map(c => c[0])], ['L', M.CODES.filter(c => c[2] === 'L').map(c => c[0])]]);
    E.dynSheet(X, SH.D2, 'Изменения финансовых результатов к аналогичному периоду прошлого года', [['P', M.CODES.filter(c => c[2] === 'P').map(c => c[0])]]);
    E.netAssets(X); E.coefficients(X); nonFin(X); scoring(X); state(X); charts(X); checks(X); sources(X); docs(X); summary(X);
    X.wb.definedNames.push && 0;
    return X;
  }
  AZ.exp.build = build;
  AZ.exp.toBlob = async S => { const X = build(S); return { blob: await X.wb.toBlob(), X }; };
})(typeof window !== 'undefined' ? window : globalThis);
