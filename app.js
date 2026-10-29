/* app.js — интерфейс: загрузка любых файлов → разбор → слияние → Excel-книга по шаблону «фин анализ». Все вычисления в браузере. */
(function () {
  'use strict';
  const AZ = window.AZ, M = AZ.model, $ = id => document.getElementById(id);
  const st = { results: [], ctx: null, busy: false, manual: { name: '', inn: '', okved: '', industry: '' } };
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const fmt = (v, d) => (typeof v === 'number' ? v.toLocaleString('ru-RU', { maximumFractionDigits: d == null ? 0 : d }) : (v === '' || v == null ? '—' : esc(v)));

  /* ── загрузка ── */
  async function collect(fileList) {
    const out = [];
    for (const f of fileList) {
      const path = f.webkitRelativePath || f.name;
      if (/\.zip$/i.test(f.name)) { try { out.push(...await AZ.imp.expandZip(await f.arrayBuffer(), path.replace(/\.zip$/i, ''))); } catch (e) { out.push({ name: f.name, path, file: f }); } }
      else if (!/^(~\$|\.DS_Store|Thumbs\.db)/.test(f.name)) out.push({ name: f.name, path, file: f, size: f.size });
    }
    return out;
  }
  async function handle(fileList) {
    if (st.busy || !fileList.length) return; st.busy = true; setBusy(true);
    try {
      const items = await collect(fileList); let n = 0;
      for (const it of items) {
        n++; progress(`Читаю ${n} из ${items.length}: ${it.name}`);
        const r = await AZ.imp.readItem(it, {}); r.category = AZ.imp.classify(r.path, r.text);
        r.entities = /Методики/.test(r.category) ? {} : M.extractEntities(r.text); r.used = AZ.imp.USED_FOR[r.category] || ''; r.text = ''; r.buf = null; it.buf = null;
        st.results.push(r); if (n % 5 === 0) await new Promise(res => setTimeout(res));
      }
      progress(''); rebuild();
    } catch (e) { progress('Ошибка: ' + (e && e.message || e)); }
    st.busy = false; setBusy(false);
  }
  function setBusy(b) { $('dropzone').classList.toggle('busy', b); }
  function progress(t) { const p = $('progress'); p.textContent = t; p.hidden = !t; }

  /* ── сборка данных и книги ── */
  function statusOf(r) {
    if (r.error) return { t: 'ошибка чтения: ' + r.error, manual: true };
    if (r.records.length) { const ks = [...new Set(r.records.map(x => x.key))].sort(M.keyCmp).filter(k => k !== 'latest'); return { t: `прочитано: ${r.records.length} значений` + (ks.length ? `, периоды ${M.periodFromKey(ks[0]).label} — ${M.periodFromKey(ks[ks.length - 1]).label}` : '') + (r.note ? '; ' + r.note : '') }; }
    if (r.scanned || r.note) return { t: r.note || 'не прочитан', manual: !!r.scanned };
    return { t: r.text === '' && !r.records.length ? 'числовых данных не найдено (учтено как документ в деле)' : '' };
  }
  function gather() {
    const ent = AZ.core.mergeEntities(st.results); const fb = $('fallbackKey').value || null;
    const merged = AZ.core.mergeRecords(st.results.flatMap(r => r.records), { fallbackKey: fb });
    const af = AZ.core.autoFill(ent, new Date()); const man = st.manual;
    const industry = man.industry || af.industry || 'Прочее';
    const data = {
      borrower: { name: man.name || ent.name || '', inn: man.inn || ent.inn || '', okved: man.okved || ent.okved || '', industry },
      periods: merged.periods, vals: merged.vals, factors: af.factors, negatives: {}, bankrupt: null,
      notes: merged.notes.concat(af.notes), docs: st.results.map(r => { const s = statusOf(r); return { name: r.path, category: r.category, used: r.used, status: s.t, manual: s.manual }; })
    };
    const hasLatest = st.results.some(r => r.records.some(x => x.key === 'latest'));
    return { data, ent, af, merged, hasLatest };
  }
  function rebuild() {
    if (!st.results.length) { $('results').classList.add('hidden'); return; }
    const g = gather(); st.g = g; fillFallback(g);
    st.ctx = AZ.exp.build(g.data); render(g);
  }
  function fillFallback(g) {
    const box = $('fallbackBox'); box.hidden = !g.hasLatest; const sel = $('fallbackKey'); if (sel.options.length) return;
    sel.innerHTML = '<option value="">— выберите —</option>'; const y = new Date().getFullYear();
    for (let yy = y; yy >= y - 5; yy--) for (let q = 4; q >= 1; q--) { const p = M.mkPeriod(yy, q); sel.innerHTML += `<option value="${p.key}">${p.label}</option>`; }
  }

  /* ── вывод ── */
  function render(g) {
    const ctx = st.ctx, b = g.data.borrower, np = ctx.np, last = np - 1;
    $('results').classList.remove('hidden');
    for (const [id, k] of [['borrower', 'name'], ['inn', 'inn'], ['okved', 'okved']]) { const el = $(id); if (document.activeElement !== el) el.value = st.manual[k] || b[k] || ''; }
    $('industry').value = b.industry;
    $('fileCount').textContent = `(${st.results.length})`;
    // карточки
    const V = (sheet, key) => (last >= 0 ? ctx.wb.value(sheet, ctx.rows[sheet][key], 3 + last) : null);
    const cards = [['Период', last >= 0 ? ctx.periods[last].label : '—'], ['Периодов', np], ['Баллов (финансы, из 90)', V('Фин показ', 'n') === 10 ? fmt(V('Фин показ', 'fin')) : 'неполно'], ['Класс (до оценки бизнес-риска)', (() => { const a = V('Фин показ', 'clsMin'), c = V('Фин показ', 'clsMax'); return a === '' || a == null ? '—' : a === c ? a : a + '–' + c; })()], ['Чистые активы, тыс. руб.', fmt(V('Чистые активы', 'na'))]];
    $('cards').innerHTML = cards.map(c => `<div class="card"><span>${c[0]}</span><b>${c[1]}</b></div>`).join('');
    // таблица показателей
    if (last >= 0) {
      const rows = ctx.sc.defs.map(d => `<tr><td>${esc(d[1])}</td><td class="n">${fmt(V('Фин показ', 'v_' + d[0]), 2)}</td><td class="n">${fmt(V('Фин показ', 'p_' + d[0]))}</td><td class="n">${d[2]}</td></tr>`).join('') + `<tr><td>Чистые активы</td><td class="n">${fmt(V('Чистые активы', 'na'))}</td><td class="n">${fmt(V('Фин показ', 'p_na'))}</td><td class="n">10</td></tr>`;
      $('kpi').innerHTML = `<table><thead><tr><th>Показатель (${esc(ctx.periods[last].label)})</th><th>Значение</th><th>Балл</th><th>Макс.</th></tr></thead><tbody>${rows}</tbody></table>`;
    } else $('kpi').innerHTML = '<p class="hint">Отчетность в документах не найдена. Книга всё равно скачается — листы «Актив», «Пассив», «Приб_Уб» можно заполнить вручную, формулы пересчитаются.</p>';
    // что заполнить в Excel
    const miss = []; const fcount = Object.keys(g.data.factors).length; miss.push(`качественные факторы бизнес-риска: подставлено ${fcount + 1} из 12 (остальные — выпадающие списки в Excel)`); miss.push('негативные явления: 7 из 12 считаются автоматически, остальные — отметки в Excel');
    $('todo').innerHTML = miss.map(x => `<li>${esc(x)}</li>`).join('');
    // замечания
    const nts = g.data.notes.filter(n => n.level === 'warn').slice(0, 12); const fl = ctx.flags.filter(f => f.level === 'warn').length;
    $('notes').innerHTML = (nts.map(n => `<li>${esc((n.period ? n.period + ': ' : '') + n.text)}</li>`).join('') + (fl ? `<li>Расхождения итогов с суммой статей: ${fl} — см. лист «Проверки».</li>` : '')) || '<li class="ok">Существенных замечаний при слиянии нет.</li>';
    // файлы
    $('fileRows').innerHTML = st.results.map((r, i) => { const s = statusOf(r); return `<tr class="${s.manual ? 'warnrow' : ''}"><td>${esc(r.path)}</td><td>${esc(r.category)}</td><td>${esc(s.t)}</td><td><button class="link" data-del="${i}" title="Убрать из анализа">✕</button></td></tr>`; }).join('');
    // полнота данных по периодам
    const need = ['1200', '1300', '1500', '1600', '2110', '2200', '2300'];
    $('coverage').innerHTML = ctx.periods.map((p, i) => { const m = need.filter(c => !AZ.exp.resolve(ctx, c, i).has); return `<span class="chip ${m.length ? 'bad' : 'good'}" title="${m.length ? 'нет: ' + m.join(', ') : 'ключевые строки найдены'}">${esc(p.label)}</span>`; }).join('') || '—';
    $('download').disabled = false;
  }

  async function download() {
    const ctx = st.ctx; if (!ctx) return; $('download').disabled = true;
    try {
      const blob = await ctx.wb.toBlob(); const b = st.g.data.borrower; const last = ctx.np ? ctx.periods[ctx.np - 1].label : 'без_периода';
      const nm = (b.name || 'заемщик').replace(/[^\wА-Яа-яЁё.\- ]+/g, '').trim().slice(0, 50).replace(/\s+/g, '_');
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `Фин_анализ_${nm}_${last.replace(/[^\wА-Яа-я]+/g, '_')}.xlsx`; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    } finally { $('download').disabled = false; }
  }

  /* ── события ── */
  function init() {
    const inp = $('fileInput'), dz = $('dropzone');
    $('chooseFiles').onclick = () => { inp.removeAttribute('webkitdirectory'); inp.value = ''; inp.click(); };
    $('chooseFolder').onclick = () => { inp.setAttribute('webkitdirectory', ''); inp.value = ''; inp.click(); };
    inp.onchange = () => handle([...inp.files]);
    dz.ondragover = e => { e.preventDefault(); dz.classList.add('over'); }; dz.ondragleave = () => dz.classList.remove('over');
    dz.ondrop = async e => { e.preventDefault(); dz.classList.remove('over'); const files = await dropped(e.dataTransfer); handle(files); };
    $('download').onclick = download;
    $('reset').onclick = () => { st.results = []; st.ctx = null; st.manual = { name: '', inn: '', okved: '', industry: '' }; $('fallbackKey').innerHTML = ''; rebuild(); };
    for (const [id, k] of [['borrower', 'name'], ['inn', 'inn'], ['okved', 'okved']]) $(id).oninput = () => { st.manual[k] = $(id).value.trim(); clearTimeout(st.t); st.t = setTimeout(rebuild, 400); };
    $('industry').onchange = () => { st.manual.industry = $('industry').value; rebuild(); };
    $('fallbackKey').onchange = rebuild;
    $('fileRows').onclick = e => { const d = e.target.dataset && e.target.dataset.del; if (d != null) { st.results.splice(+d, 1); rebuild(); } };
    let deferred; window.addEventListener('beforeinstallprompt', e => { e.preventDefault(); deferred = e; $('installApp').hidden = false; }); $('installApp').onclick = () => deferred && deferred.prompt();
    if ('serviceWorker' in navigator && location.protocol.startsWith('http')) navigator.serviceWorker.register('sw.js').catch(() => { });
  }
  /** Перетаскивание папок: обходим дерево каталогов */
  async function dropped(dt) {
    const items = dt.items ? [...dt.items].map(i => i.webkitGetAsEntry && i.webkitGetAsEntry()) : []; if (!items.length || !items[0]) return [...dt.files];
    const files = []; const walk = async (en, base) => {
      if (en.isFile) await new Promise(res => en.file(f => { try { Object.defineProperty(f, 'webkitRelativePath', { value: base + f.name }); } catch (e) { } files.push(f); res(); }, res));
      else if (en.isDirectory) { const rd = en.createReader(); let batch; do { batch = await new Promise(res => rd.readEntries(res, () => res([]))); for (const c of batch) await walk(c, base + en.name + '/'); } while (batch.length); }
    };
    for (const en of items) if (en) await walk(en, ''); return files;
  }
  document.addEventListener('DOMContentLoaded', init);
})();
