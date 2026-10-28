/* app.js — интерфейс: загрузка любых файлов → разбор → слияние → предпросмотр расчета → книга Excel. Всё в браузере. */
(function () {
  'use strict';
  const AZ = window.AZ, M = AZ.model, $ = id => document.getElementById(id);
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const st = {
    results: [], items: new Map(), edits: {}, extra: [], borrower: { name: '', inn: '', ogrn: '', okved: '', regDate: '' }, touched: {},
    industry: 'Прочее', factors: Array(12).fill(null), factorAuto: {}, factorTouched: {}, factorNotes: {}, negManual: Array(12).fill(null), bankrupt: 'Нет', merged: null, X: null
  };
  const KEYROWS = ['1100', '1200', '1230', '1300', '1400', '1500', '1520', '1600', '1700', '2110', '2120', '2200', '2300', '2400', '5445', '5540', '5320'];
  const nf = v => v == null ? '' : Math.round(v).toLocaleString('ru-RU');

  /* ── загрузка файлов ── */
  async function entriesFromDrop(dt) {
    const out = []; const walk = async (entry, path) => {
      if (entry.isFile) await new Promise(res => entry.file(f => { out.push({ file: f, name: f.name, path: path + f.name, size: f.size }); res(); }, res));
      else if (entry.isDirectory) { const rd = entry.createReader(); let batch; do { batch = await new Promise(res => rd.readEntries(res, () => res([]))); for (const e of batch) await walk(e, path + entry.name + '/'); } while (batch.length); }
    };
    const items = [...(dt.items || [])].map(i => i.webkitGetAsEntry && i.webkitGetAsEntry()).filter(Boolean);
    if (items.length) { for (const e of items) await walk(e, ''); } else for (const f of dt.files) out.push({ file: f, name: f.name, path: f.name, size: f.size });
    return out;
  }
  const fromInput = files => [...files].map(f => ({ file: f, name: f.name, path: f.webkitRelativePath || f.name, size: f.size }));
  const SKIP = /^(~\$|\.DS_Store|Thumbs\.db|desktop\.ini)/i;

  async function addFiles(list) {
    list = list.filter(x => !SKIP.test(x.name)); if (!list.length) return;
    const pr = $('progress'); pr.classList.remove('hidden');
    const queue = []; for (const it of list) { if (/\.zip$/i.test(it.name)) { try { const buf = await it.file.arrayBuffer(); queue.push(...await AZ.imp.expandZip(buf, it.path.replace(/\.zip$/i, ''))); } catch (e) { queue.push(it); } } else queue.push(it); }
    let n = 0;
    for (const it of queue) {
      if (st.items.has(it.path + '|' + it.size)) { n++; continue; }
      $('progressText').textContent = 'Читаю ' + it.path + ' (' + (++n) + ' из ' + queue.length + ')'; $('bar').style.width = (100 * n / queue.length) + '%';
      await new Promise(r => setTimeout(r, 0));
      const res = await AZ.imp.readItem(it); res.kind = AZ.imp.classify(res.path, res.text); res.ent = M.extractEntities(res.text);
      st.items.set(it.path + '|' + it.size, true); st.results.push(res);
    }
    pr.classList.add('hidden'); recompute(true);
  }

  /* ── пересчет данных из всех файлов ── */
  function entities() {
    const pick = f => { const reg = st.results.filter(r => /Регистрационные/.test(r.kind)); for (const g of [reg, st.results]) for (const r of g) if (r.ent && r.ent[f]) return r.ent[f]; return null; };
    return { name: pick('name'), inn: pick('inn'), ogrn: pick('ogrn'), okved: pick('okved'), regDate: pick('regDate') };
  }
  function build() {
    const recs = st.results.flatMap(r => r.records || []);
    const m = AZ.merge.merge(recs);
    for (const [k, v] of Object.entries(st.edits)) { const [key, code] = k.split('|'); (m.data[key] = m.data[key] || {})[code] = v; (m.src[key] = m.src[key] || {})[code] = 'введено вручную'; if (!m.periods.includes(key)) m.periods.push(key); }
    for (const key of st.extra) if (!m.periods.includes(key)) { m.periods.push(key); m.data[key] = m.data[key] || {}; m.src[key] = m.src[key] || {}; }
    m.periods.sort(M.keyCmp); if (m.periods.length > 41) m.periods = m.periods.slice(-41);
    return m;
  }
  function autoFactors(m) {
    const f = st.factorAuto; const note = st.factorNotes;
    const b = st.borrower; let months = null;
    const rm = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec(b.regDate || '');
    if (rm) { const d = new Date(+rm[3], +rm[2] - 1, +rm[1]); const last = m.periods.length ? M.periodFromKey(m.periods[m.periods.length - 1]) : null; const ref = last ? new Date(last.year, last.q * 3, 0) : new Date(); months = (ref.getFullYear() - d.getFullYear()) * 12 + ref.getMonth() - d.getMonth(); }
    if (!st.factorTouched[5] && months != null) { st.factors[5] = months > 12 ? 3 : months >= 6 ? 2 : 1; f[5] = 1; note[5] = 'Автоматически: дата регистрации ' + b.regDate + ' → ' + months + ' мес. на конец анализируемого периода'; }
    const lastKey = m.periods[m.periods.length - 1]; const d = lastKey && m.data[lastKey] || {};
    if (!st.factorTouched[11] && d['1600'] != null && d['1310'] != null) { const na = d['1600'] - ((d['1400'] || 0) + (d['1500'] || 0) - (d['1530'] || 0)); st.factors[11] = na < d['1310'] ? 0 : 2; f[11] = 1; note[11] = 'Автоматически: чистые активы ' + nf(na) + ' ' + (na < d['1310'] ? '<' : '≥') + ' уставного капитала ' + nf(d['1310']) + ' тыс. руб. (п. 5.3.12). Проверьте документально подтвержденные сведения о ликвидации/банкротстве.'; }
  }
  function recompute(fromFiles) {
    const e = entities();
    for (const k of ['name', 'inn', 'ogrn', 'okved']) if (!st.touched[k] && e[k]) st.borrower[k] = e[k];
    if (!st.touched.regDate && e.regDate) st.borrower.regDate = e.regDate.text;
    if (!st.touched.industry && st.borrower.okved) { const i = M.industryByOkved(st.borrower.okved); if (i) st.industry = i; }
    st.merged = build(); autoFactors(st.merged);
    $('app').classList.remove('hidden');
    if (fromFiles) { renderCard(); renderGrid(); renderFactors(); renderNegs(); renderFiles(); }
    renderPreview();
  }
  function state() {
    const m = st.merged; const docs = st.results.map(r => ({ path: r.path, kind: r.kind, usedFor: AZ.imp.USED_FOR[r.kind] || '', status: r.error ? 'ошибка: ' + r.error : r.scanned ? 'скан — не прочитан' : r.records.length ? 'прочитан: ' + r.records.length + ' значений' : r.text ? 'прочитан (текст)' : 'не прочитан', note: r.note }));
    return { borrower: st.borrower, industry: st.industry, industryNote: st.touched.industry ? 'выбрано вручную' : st.borrower.okved ? 'определено по ОКВЭД ' + st.borrower.okved : 'ОКВЭД не найден — выберите отрасль', periods: m.periods, data: m.data, src: m.src, conflicts: m.conflicts, notes: m.notes, factors: st.factors, factorAuto: st.factorAuto, factorNotes: st.factorNotes, negManual: st.negManual, bankrupt: st.bankrupt, docs };
  }

  /* ── карточка ── */
  function renderCard() {
    const b = st.borrower; $('f_name').value = b.name || ''; $('f_inn').value = b.inn || ''; $('f_ogrn').value = b.ogrn || ''; $('f_okved').value = b.okved || ''; $('f_reg').value = b.regDate || ''; $('f_ind').value = st.industry;
    $('indHint').textContent = b.inn && !M.innValid(b.inn) ? '⚠ ИНН не проходит проверку контрольной суммы — проверьте.' : (b.okved ? 'Отрасль определена по ОКВЭД ' + b.okved + ' (п. 5.1.11 методики: по виду деятельности с наибольшей выручкой — подтвердите).' : 'ОКВЭД не найден в документах — выберите отрасль вручную.');
  }
  for (const [id, key] of [['f_name', 'name'], ['f_inn', 'inn'], ['f_ogrn', 'ogrn'], ['f_okved', 'okved'], ['f_reg', 'regDate']]) $(id).addEventListener('input', () => { st.borrower[key] = $(id).value.trim(); st.touched[key] = true; if (key === 'okved' && !st.touched.industry) { const i = M.industryByOkved(st.borrower.okved); if (i) { st.industry = i; $('f_ind').value = i; } } if (key === 'regDate') { st.factorTouched[5] = false; } schedule(); });
  $('f_ind').addEventListener('change', () => { st.industry = $('f_ind').value; st.touched.industry = true; schedule(); });

  /* ── сетка данных ── */
  function renderGrid() {
    const m = st.merged; const P = m.periods; const tb = $('grid');
    let h = '<thead><tr><th>Статья</th><th>Код</th>' + P.map(k => '<th>' + esc(M.periodFromKey(k).label) + '</th>').join('') + '</tr></thead><tbody>';
    for (const code of KEYROWS) {
      h += '<tr><td>' + esc(M.CODE_NAME[code]) + '</td><td class="mono">' + code + '</td>';
      for (const key of P) { const v = m.data[key] && m.data[key][code]; const s = m.src[key] && m.src[key][code]; const manual = st.edits[key + '|' + code] != null; const req = ['1200', '1300', '1500', '1600', '1700', '2110', '2200', '2300'].includes(code); h += '<td class="' + (v == null && req ? 'need' : manual ? 'manual' : '') + '"><input data-k="' + key + '|' + code + '" value="' + (v == null ? '' : Math.round(v * 1000) / 1000) + '" title="' + esc(s || 'нет данных') + '" inputmode="decimal"></td>'; }
      h += '</tr>';
    }
    tb.innerHTML = h + '</tbody>'; const wrap = tb.parentElement; if (wrap) wrap.scrollLeft = wrap.scrollWidth;
    $('conflicts').innerHTML = m.conflicts.length ? '<div class="hint-box amber">Расхождения между документами (' + m.conflicts.length + '): принято значение из «текущего» столбца отчета; все случаи — на листе «Проверки». Пример: ' + esc(M.CODE_NAME[m.conflicts[0].code]) + ', ' + esc(M.periodFromKey(m.conflicts[0].key).label) + ': ' + nf(m.conflicts[0].a) + ' ≠ ' + nf(m.conflicts[0].b) + '.</div>' : '';
  }
  $('grid').addEventListener('change', e => { const k = e.target.dataset && e.target.dataset.k; if (!k) return; const raw = e.target.value.trim(); const v = M.readNum(raw.replace(/\s/g, '')); if (!raw) delete st.edits[k]; else if (v != null) st.edits[k] = k.endsWith('|1320') ? -Math.abs(v) : v; schedule(true); });
  $('ap_btn').addEventListener('click', () => { const y = +$('ap_year').value, q = +$('ap_q').value; if (y >= 2000 && y <= 2100) { const k = y + '-' + q; if (!st.extra.includes(k)) st.extra.push(k); recompute(true); } });

  /* ── факторы ── */
  function renderFactors() {
    $('factors').innerHTML = '<thead><tr><th>Показатель</th><th>Балл</th><th>Основание</th></tr></thead><tbody>' + M.FACTORS.map((f, i) => '<tr><td><b>' + (i + 1) + '. ' + esc(f.n) + '</b><br><small>п. ' + f.p + ' · макс. ' + f.max + '</small></td><td class="' + (st.factorAuto[i] ? 'auto' : st.factors[i] == null ? 'need' : '') + '"><select data-f="' + i + '"><option value="">—</option>' + f.opts.map(o => '<option value="' + o[0] + '"' + (st.factors[i] === o[0] ? ' selected' : '') + '>' + o[0] + ' — ' + esc(o[1]) + '</option>').join('') + '</select></td><td><small>' + esc(st.factorNotes[i] || '') + '</small></td></tr>').join('') + '</tbody>';
  }
  $('factors').addEventListener('change', e => { const i = e.target.dataset && e.target.dataset.f; if (i == null) return; const v = e.target.value; st.factors[i] = v === '' ? null : Number(v); st.factorTouched[i] = true; delete st.factorAuto[i]; st.factorNotes[i] = v === '' ? '' : 'Оценка аналитика'; e.target.parentElement.className = v === '' ? 'need' : ''; schedule(); });
  function renderNegs() {
    const rows = M.NEGATIVES.map((n, i) => ({ n, i })).filter(x => !x.n.auto);
    $('negs').innerHTML = '<thead><tr><th>Пункт</th><th>Явление</th><th>Выявлено?</th></tr></thead><tbody>' + rows.map(({ n, i }) => '<tr><td class="mono">' + n.p + '</td><td>' + esc(n.t) + '</td><td class="' + (st.negManual[i] ? '' : 'need') + '"><select data-n="' + i + '"><option value="">—</option><option' + (st.negManual[i] === 'Да' ? ' selected' : '') + '>Да</option><option' + (st.negManual[i] === 'Нет' ? ' selected' : '') + '>Нет</option></select></td></tr>').join('') + '<tr><td class="mono">7.3</td><td>Признан банкротом либо устойчиво неплатежеспособен более 180 дней</td><td><select id="bk"><option' + (st.bankrupt === 'Нет' ? ' selected' : '') + '>Нет</option><option' + (st.bankrupt === 'Да' ? ' selected' : '') + '>Да</option></select></td></tr></tbody>';
  }
  $('negs').addEventListener('change', e => { if (e.target.id === 'bk') st.bankrupt = e.target.value; else if (e.target.dataset.n != null) { st.negManual[e.target.dataset.n] = e.target.value || null; e.target.parentElement.className = e.target.value ? '' : 'need'; } schedule(); });

  /* ── реестр ── */
  function renderFiles() {
    $('fileCount').textContent = '· ' + st.results.length;
    $('fileRows').innerHTML = st.results.map(r => '<tr><td>' + esc(r.path) + '</td><td>' + esc(r.kind) + '</td><td><small>' + esc(AZ.imp.USED_FOR[r.kind] || '—') + '</small></td><td class="' + (r.error || r.scanned ? 'need' : '') + '"><small>' + esc(r.error ? 'ошибка: ' + r.error : r.scanned ? r.note : r.records.length ? 'значений: ' + r.records.length + (r.note ? '; ' + r.note : '') : (r.note || 'текст прочитан')) + '</small></td></tr>').join('');
  }

  /* ── предпросмотр (значения из книги) ── */
  let timer = null; function schedule(grid) { clearTimeout(timer); timer = setTimeout(() => { st.merged = build(); autoFactors(st.merged); if (grid) renderGrid(); renderPreview(); }, 150); }
  function renderPreview() {
    const box = $('preview'); const w = $('warnings');
    if (!st.merged.periods.length) { box.innerHTML = '<div class="hint-box amber">Периоды отчетности не найдены. Загрузите баланс и отчет о финрезультатах в Excel/PDF с текстом/Word либо добавьте период ниже и введите суммы вручную.</div>'; w.innerHTML = ''; $('download').disabled = true; return; }
    $('download').disabled = false; let X; try { X = AZ.exp.build(state()); } catch (e) { box.innerHTML = '<div class="hint-box amber">Ошибка расчета: ' + esc(e.message) + '</div>'; return; }
    st.X = X; const last = X.n - 1; const v = (sheet, row) => X.wb.value(sheet, row, 3 + last); const SH = AZ.exp.SH;
    const tot = v(SH.S, X.K.s.tot), cls = v(SH.S, X.K.s.cls), clsr = v(SH.S, X.K.s.clsr), pos = v(SH.F, X.K.f.fin, last), n = v(SH.F, X.K.f.n), open = v(SH.F, X.K.f.open);
    const cnt = st.factors.filter(x => x != null).length;
    const kp = [['na', 'Чистые активы, тыс. руб.', SH.NA, X.K.na.na, 'na'], ['ka', 'Кфн — автономия', SH.K, X.K.k.ka, 'ka'], ['ko', 'Ко — обеспеченность СОС', SH.K, X.K.k.ko, 'ko'], ['kl', 'Кл — текущая ликвидность', SH.K, X.K.k.kl, 'kl'], ['kpl', 'Кпл — платежеспособность, мес.', SH.K, X.K.k.kpl, 'kpl'], ['kob', 'Коб — оборач. оборотных средств', SH.K, X.K.k.kob, 'kob'], ['kobdz', 'Кобдз — оборач. дебиторки', SH.K, X.K.k.kobdz, 'kobdz'], ['rpr', 'Рпр — рентабельность продаж, %', SH.K, X.K.k.rpr, 'rpr'], ['ra', 'Ра — рентабельность активов, %', SH.K, X.K.k.ra, 'ra'], ['rsk', 'Рск — рентабельность СК, %', SH.K, X.K.k.rsk, 'rsk']];
    const fmt = x => x === '' || x == null ? '—' : typeof x === 'number' ? (Math.abs(x) >= 1000 ? nf(x) : String(x).replace('.', ',')) : esc(x);
    box.innerHTML = '<div class="summary-grid"><div class="summary-card"><strong>' + fmt(tot) + '</strong><span>баллов из 100 · период ' + esc(X.lab(X.periods[last])) + '</span></div><div class="summary-card"><strong>' + (cnt < 12 ? fmt(clsr) : fmt(cls)) + '</strong><span>класс кредитоспособности' + (cnt < 12 ? ' (диапазон: бизнес-риск не оценен — ' + cnt + ' из 12)' : '') + '</span></div><div class="summary-card"><strong>' + fmt(pos) + '</strong><span>финансовое положение · негативных явлений: ' + fmt(n) + (open > 0 ? ' (не оценено: ' + open + ')' : '') + '</span></div><div class="summary-card"><strong>' + X.n + '</strong><span>периодов в книге</span></div></div><div class="table-wrap"><table><thead><tr><th>Показатель</th><th>Значение</th><th>Балл</th></tr></thead><tbody>' + kp.map(([key, t, sh, row, sk]) => '<tr><td>' + esc(t) + '</td><td>' + fmt(v(sh, row)) + '</td><td>' + fmt(v(SH.S, X.K.s[sk])) + '</td></tr>').join('') + '</tbody></table></div>';
    const wl = []; if (st.merged.notes.length) wl.push(...st.merged.notes); if (X.checks.length) wl.push('Итоговые строки не сходятся с суммой статей: ' + X.checks.length + ' (см. лист «Проверки»)'); if (st.merged.conflicts.length) wl.push('Расхождения между документами: ' + st.merged.conflicts.length);
    const noStart = X.periods.filter((k, i) => X.startIdx(i) < 0).length; if (noStart) wl.push('Для ' + noStart + ' периодов нет баланса на 31.12 прошлого года — коэффициенты оборачиваемости и Ра считаются по остатку на конец периода (оранжевые ячейки).');
    const sc = st.results.filter(r => r.scanned).length; if (sc) wl.push('Сканов/изображений без текста: ' + sc + ' — суммы из них нужно ввести вручную.');
    w.innerHTML = wl.length ? '<div class="hint-box amber"><b>Обратите внимание</b><ul>' + wl.map(x => '<li>' + esc(x) + '</li>').join('') + '</ul></div>' : '';
  }

  /* ── скачивание ── */
  $('download').addEventListener('click', async () => {
    try { const { blob, X } = await AZ.exp.toBlob(state()); const a = document.createElement('a'); const nm = (st.borrower.name || 'заемщик').replace(/[\\/:*?"<>|«»]/g, '').replace(/\s+/g, '_').slice(0, 40); a.href = URL.createObjectURL(blob); a.download = 'Фин_анализ_' + nm + '_' + X.periods[X.n - 1] + '.xlsx'; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 5000); }
    catch (e) { alert('Не удалось собрать книгу: ' + e.message); }
  });

  /* ── события загрузки ── */
  const dz = $('dropzone');
  $('chooseFolder').addEventListener('click', () => $('dirInput').click()); $('chooseFiles').addEventListener('click', () => $('fileInput').click());
  for (const id of ['fileInput', 'dirInput']) $(id).addEventListener('change', e => { addFiles(fromInput(e.target.files)); e.target.value = ''; });
  ['dragenter', 'dragover'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add('drag'); }));
  ['dragleave', 'drop'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove('drag'); }));
  dz.addEventListener('drop', async e => { addFiles(await entriesFromDrop(e.dataTransfer)); });
  if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
  $('ap_year').value = new Date().getFullYear();
  let deferred; window.addEventListener('beforeinstallprompt', e => { e.preventDefault(); deferred = e; $('installApp').hidden = false; }); $('installApp').addEventListener('click', async () => { if (deferred) { deferred.prompt(); deferred = null; $('installApp').hidden = true; } });
  if ('serviceWorker' in navigator && location.protocol.startsWith('http')) navigator.serviceWorker.register('sw.js').catch(() => { });
  window.AZ.app = { st, addFiles, recompute, state };
})();
