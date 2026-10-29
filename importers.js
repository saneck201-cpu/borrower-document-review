/* importers.js — чтение документов заемщика: Excel/CSV, PDF (текстовый слой), Word, ZIP, RTF, старый DOC.
 * Чистые функции разбора (importGrid, parsePdfPage) не зависят от браузера и тестируются отдельно. */
(function (root) {
  'use strict';
  const AZ = root.AZ = root.AZ || {};
  const M = AZ.model;

  const cellText = v => (v == null ? '' : v instanceof Date ? '' : String(v).replace(/[ \s]+/g, ' ').trim());
  const rowCodeOf = cells => {
    for (let i = 0; i < Math.min(cells.length, 8); i++) {
      const x = cells[i]; if (x == null || x === '') continue;
      let s = typeof x === 'number' ? (Number.isInteger(x) ? String(x) : '') : String(x).trim();
      const m = /^(\d{4})(?:\s*[.:\-)]?\s*)?$/.exec(s) || /^(\d{4})(?=\s)/.exec(s);
      if (m) { let c = m[1]; if (M.CODE_ALIAS[c]) c = M.CODE_ALIAS[c]; if (M.CODE_SET.has(c)) return { code: c, idx: i }; }
    }
    return null;
  };

  /** Сколько ячеек строки — подписи периодов */
  function headerRow(grid, r) {
    const row = grid[r] || []; const found = new Map(); const prev = grid[r - 1] || [];
    for (let c = 0; c < row.length; c++) {
      const v = row[c]; let pi = null;
      if (v instanceof Date) pi = M.periodInfo(v);
      else if (typeof v === 'number') { if (Number.isInteger(v) && v >= 2000 && v <= 2100) pi = M.periodInfo(v); else if (v > 36500 && v < 73000 && root.XLSX && XLSX.SSF) { const d = XLSX.SSF.parse_date_code(v); if (d) pi = M.periodInfo(new Date(d.y, d.m - 1, d.d)); } }
      else { const t = cellText(v); if (t) { pi = M.periodInfo(t); if (!pi && cellText(prev[c])) pi = M.periodInfo(cellText(prev[c]) + ' ' + t); } }
      if (pi) found.set(c, pi);
    }
    return found;
  }

  /** Разбор таблицы (массив строк) в записи {code,key,label,value,src,role}. */
  function importGrid(grid, ctx) {
    ctx = ctx || {}; const out = []; let hdr = new Map(); let units = ctx.units || null; let titlePeriod = null; const seen = new Set();
    const headText = []; for (let r = 0; r < Math.min(grid.length, 40); r++) headText.push((grid[r] || []).map(cellText).join(' '));
    const u0 = M.detectUnits(headText.join(' ')); if (u0) units = u0;
    let scale = units ? units.f : 1; let current = null;
    for (let r = 0; r < grid.length; r++) {
      const row = grid[r] || []; const rc = rowCodeOf(row);
      if (!rc) {
        const h = headerRow(grid, r);
        if (h.size) { hdr = h; const mx = [...h.values()].sort((a, b) => M.keyCmp(b.key, a.key))[0]; current = mx.key; }
        else { const tx = row.map(cellText).join(' '); const un = M.detectUnits(tx); if (un && r > 3) scale = un.f; const pi = tx.length < 200 ? M.periodInfo(tx) : null; if (pi) titlePeriod = pi; }
        continue;
      }
      const { code, idx } = rc; const vals = [];
      for (let c = 0; c < row.length; c++) {
        if (c === idx) continue; let n = M.readNum(row[c]); if (n === null && M.isDash(row[c]) && hdr.has(c)) n = 0; if (n === null) continue;
        vals.push([c, n]);
      }
      if (!vals.length) continue;
      let used = 0;
      for (const [c, n] of vals) {
        let pi = hdr.get(c);
        if (!pi) continue; // столбцы без периода (код, примечания) пропускаем
        const dk = code + '|' + c + '|' + r; if (seen.has(dk)) continue; seen.add(dk);
        // повтор кода в том же столбце ниже (например, квартальный блок шаблона) игнорируем
        const key2 = code + '|' + pi.key + '|' + ctx.sheet; if (seen.has(key2)) continue; seen.add(key2);
        const value = (code === '1320' ? -Math.abs(n) : n) * scale;
        out.push({ code, key: pi.key, label: pi.label, value, src: (ctx.src || '') + (ctx.sheet ? ' · ' + ctx.sheet : '') + ' · стр. ' + (r + 1), role: pi.key === current ? 'cur' : 'cmp', units: units ? units.label : null, unitsKnown: !!units, kind: pi.kind });
        used++;
      }
      if (!used && !hdr.size) {
        // таблица без заголовков периодов: единственное число → «последний период»
        const nums = vals.filter(([c]) => !(c === idx)); const pi = titlePeriod;
        if (nums.length === 1) {
          const n = nums[0][1]; const value = (code === '1320' ? -Math.abs(n) : n) * scale;
          out.push({ code, key: pi ? pi.key : 'latest', label: pi ? pi.label : 'Последний период (дата не распознана)', value, src: (ctx.src || '') + (ctx.sheet ? ' · ' + ctx.sheet : '') + ' · стр. ' + (r + 1), role: 'cur', units: units ? units.label : null, unitsKnown: !!units, kind: pi ? pi.kind : null });
        }
      }
    }
    return out;
  }

  /* ── Листы Excel: определяем, какие листы содержат первичные данные ── */
  const DERIVED_SHEET = /удельн|изм\.?\s*стр|коэфф|график|справочн|фин\s*показ|нефин|рейтинг|чистые\s*актив|резерв|расчет|субъект|провер|источник|документ|сводка|фин\s*состоян/i;
  const PRIMARY_SHEET = /актив|пассив|баланс|приб|убыт|финрез|отчет|ф1|ф2|форма|опу|бух/i;
  function importWorkbookObj(wb, file) {
    const recs = []; const names = wb.SheetNames.slice().sort((a, b) => (PRIMARY_SHEET.test(a) ? 0 : 1) - (PRIMARY_SHEET.test(b) ? 0 : 1));
    const skipped = [];
    for (const name of names) {
      if (DERIVED_SHEET.test(name)) { skipped.push(name); continue; }
      const ws = wb.Sheets[name]; if (!ws) continue;
      let grid; try { grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' }); } catch (e) { continue; }
      if (!grid.length) continue;
      recs.push(...importGrid(grid, { src: file.path || file.name, sheet: name }));
    }
    return { records: dropEmptyPeriods(recs), skipped };
  }
  /** Периоды, где все значения нулевые (пустые колонки шаблонов), отбрасываем */
  function dropEmptyPeriods(recs) {
    const nz = new Set(); for (const r of recs) if (r.value) nz.add(r.key);
    return recs.filter(r => nz.has(r.key) || r.key === 'latest');
  }

  /* ── PDF: строки по координатам ── */
  function parsePdfPage(items, ctx, carryAnchors) {
    ctx = ctx || {}; const lines = new Map();
    for (const z of items) { if (!z.str || !z.str.trim()) continue; const y = Math.round((z.transform ? z.transform[5] : 0) / 3) * 3; if (!lines.has(y)) lines.set(y, []); lines.get(y).push({ x: z.transform ? z.transform[4] : 0, str: z.str.trim(), w: z.width || 0 }); }
    const ordered = [...lines.entries()].sort((a, b) => b[0] - a[0]).map(e => e[1].sort((p, q) => p.x - q.x));
    const pageText = ordered.map(l => l.map(c => c.str).join(' ')).join('\n');
    const un = M.detectUnits(pageText); const scale = ctx.units ? ctx.units.f : (un ? un.f : 1); const units = ctx.units || un;
    // якоря периодов: слева направо
    let anchors = []; for (const line of ordered) { if (line.some(c => rowCodeOf([c.str]))) continue; for (let i = 0; i < line.length; i++) { let pi = M.periodInfo(line[i].str); if (!pi && i + 1 < line.length) pi = M.periodInfo(line[i].str + ' ' + line[i + 1].str); if (pi) anchors.push({ x: line[i].x + (line[i].w || 0) / 2, pi }); } }
    const uniq = []; for (const a of anchors) if (!uniq.some(u => u.pi.key === a.pi.key && Math.abs(u.x - a.x) < 40)) uniq.push(a);
    let useAnch = uniq.length ? uniq : (carryAnchors || []);
    const out = []; const curKey = useAnch.length ? [...useAnch].sort((a, b) => M.keyCmp(b.pi.key, a.pi.key))[0].pi.key : null;
    for (const line of ordered) {
      const ix = line.findIndex(c => rowCodeOf([c.str])); if (ix < 0) continue; const rc = rowCodeOf([line[ix].str]);
      // PDF-экстракторы дробят «1 114 303» на куски: склеиваем соседние токены в пределах одного столбца
      const groups = []; for (const cell of line.slice(ix + 1)) { const g = groups[groups.length - 1]; if (!g || cell.x - (g.x + g.w) > 7) groups.push({ x: cell.x, w: cell.w || 0, parts: [cell.str] }); else { g.parts.push(cell.str); g.w = cell.x + (cell.w || 0) - g.x; } }
      const vals = groups.map(g => { const t = g.parts.join(' ').trim(); return { x: g.x, n: M.isDash(t) ? 0 : M.readNum(t) }; }).filter(v => v.n !== null);
      if (!vals.length) continue;
      if (useAnch.length) { for (const v of vals) { const near = useAnch.reduce((a, b) => Math.abs(b.x - v.x) < Math.abs(a.x - v.x) ? b : a); if (Math.abs(near.x - v.x) > 120) continue; const n = rc.code === '1320' ? -Math.abs(v.n) : v.n; out.push({ code: rc.code, key: near.pi.key, label: near.pi.label, value: n * scale, src: (ctx.src || '') + ' · PDF стр. ' + (ctx.page || 1), role: near.pi.key === curKey ? 'cur' : 'cmp', units: units ? units.label : null, unitsKnown: !!units, kind: near.pi.kind }); } }
      else if (vals.length === 1) { const n = rc.code === '1320' ? -Math.abs(vals[0].n) : vals[0].n; out.push({ code: rc.code, key: 'latest', label: 'Последний период (дата не распознана)', value: n * scale, src: (ctx.src || '') + ' · PDF стр. ' + (ctx.page || 1), role: 'cur', units: units ? units.label : null, unitsKnown: !!units }); }
    }
    return { records: out, text: pageText, anchors: uniq.length ? uniq : (carryAnchors || []), units };
  }

  /* ── DOCX (через JSZip): текст + таблицы ── */
  async function readDocx(buf) {
    if (!root.JSZip) return { text: '', grids: [] };
    const zip = await JSZip.loadAsync(buf); const f = zip.file('word/document.xml'); if (!f) return { text: '', grids: [] };
    const xml = await f.async('string'); const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const textOf = el => { let s = ''; const walk = n => { for (const ch of n.childNodes) { if (ch.nodeType !== 1) continue; if (ch.localName === 't') s += ch.textContent; else if (ch.localName === 'tab') s += ' '; else if (ch.localName === 'br') s += ' '; else walk(ch); } }; walk(el); return s; };
    const grids = []; const parts = [];
    const body = doc.getElementsByTagNameNS(W, 'body')[0]; if (!body) return { text: '', grids: [] };
    for (const ch of body.childNodes) {
      if (ch.nodeType !== 1) continue;
      if (ch.localName === 'p') { parts.push(textOf(ch)); }
      else if (ch.localName === 'tbl') {
        const grid = [];
        for (const tr of ch.getElementsByTagNameNS(W, 'tr')) {
          const row = [];
          for (const tc of tr.childNodes) { if (tc.nodeType !== 1 || tc.localName !== 'tc') continue; const ps = [...tc.getElementsByTagNameNS(W, 'p')].map(textOf).filter(Boolean).join(' ').trim(); const gs = tc.getElementsByTagNameNS(W, 'gridSpan')[0]; const span = gs ? Number(gs.getAttribute('w:val') || gs.getAttributeNS(W, 'val') || 1) : 1; row.push(ps); for (let k = 1; k < span; k++) row.push(''); }
          grid.push(row); parts.push(row.join(' | '));
        }
        grids.push(grid);
      }
    }
    return { text: parts.join('\n'), grids };
  }

  /* ── старый DOC / RTF: только текст для классификации и реквизитов ── */
  function legacyDocText(buf) {
    const u8 = new Uint8Array(buf); const chunks = [];
    try { const u16 = new TextDecoder('utf-16le', { fatal: false }).decode(u8); const m = u16.match(/[А-Яа-яЁёA-Za-z0-9№"«»().,;:\-–/\s]{40,}/g); if (m) chunks.push(...m.filter(s => /[А-Яа-я]{3}/.test(s))); } catch (e) { }
    try { const w = new TextDecoder('windows-1251', { fatal: false }).decode(u8); const m = w.match(/[А-Яа-яЁёA-Za-z0-9№"«»().,;:\-–/ ]{40,}/g); if (m) chunks.push(...m.filter(s => /[А-Яа-я]{3}/.test(s))); } catch (e) { }
    return [...new Set(chunks.map(s => s.replace(/\s+/g, ' ').trim()))].join('\n').slice(0, 300000);
  }
  function rtfText(s) {
    return s.replace(/\\'([0-9a-f]{2})/gi, (_, h) => { try { return new TextDecoder('windows-1251').decode(new Uint8Array([parseInt(h, 16)])); } catch (e) { return ''; } })
      .replace(/\\u(-?\d+)\??/g, (_, n) => String.fromCharCode(n < 0 ? 65536 + +n : +n)).replace(/\\par[d]?/g, '\n').replace(/\\[a-z]+-?\d* ?/gi, '').replace(/[{}]/g, '');
  }

  /* ── ZIP ── */
  async function expandZip(buf, basePath, depth) {
    depth = depth || 0; const out = []; if (!root.JSZip || depth > 3) return out;
    let zip; try { zip = await JSZip.loadAsync(buf, { decodeFileName: bytes => { try { return new TextDecoder('utf-8', { fatal: true }).decode(bytes); } catch (e) { return new TextDecoder('ibm866').decode(bytes); } } }); } catch (e) { return out; }
    for (const [name, entry] of Object.entries(zip.files)) {
      if (entry.dir) continue; const base = name.split('/').pop(); if (/^(~\$|\.DS_Store|Thumbs\.db)/.test(base) || /__MACOSX\//.test(name)) continue;
      const data = await entry.async('arraybuffer'); const path = basePath + '/' + name;
      if (/\.zip$/i.test(base)) out.push(...await expandZip(data, path.replace(/\.zip$/i, ''), depth + 1)); else out.push({ name: base, path, buf: data, size: data.byteLength });
    }
    return out;
  }

  /* ── Классификация ── */
  const KINDS = [
    ['Бухгалтерский баланс (форма 1)', ['бухгалтерский баланс', 'форма 1', 'ф1', 'баланс', 'форма по окуд 0710001', 'форме 1']],
    ['Отчет о финансовых результатах (форма 2)', ['финансовых результатах', 'форма 2', 'ф2', 'финрез', 'прибылях и убытках', 'опу', 'окуд 0710002']],
    ['Пояснения к отчетности / расшифровки', ['пояснени', 'расшифровк', 'расшифровка', 'дебиторская задолженность', 'кредиторская задолженность', 'основные средства']],
    ['Аудиторское заключение', ['аудитор']],
    ['Налоговая отчетность', ['декларац', 'ндс', 'усн', 'налог на прибыль', 'справка об исполнении', 'расчет по страховым']],
    ['Регистрационные документы / ЕГРЮЛ', ['егрюл', 'выписка из единого', 'устав', 'огрн', 'учредител', 'свидетельство', 'лист записи']],
    ['Банковские выписки / обороты', ['выписка по счет', 'выписка по расчетному', 'выписка банка', 'оборот по счет', 'обороты по счет', 'оборотно-сальдов', 'движение средств по счет']],
    ['Картотека / справки банка', ['картотек', 'справка об отсутствии', 'справка о наличии', 'ссудная задолженность']],
    ['Судебные дела', ['арбитраж', 'кад.арбитр', 'судебн', 'иск', 'банкрот']],
    ['Кредитные договоры и обязательства', ['кредитный договор', 'график погашения', 'займ', 'заем', 'лизинг', 'договор']],
    ['Бизнес-план / ТЭО', ['бизнес-план', 'бизнес план', 'тэо', 'прогноз', 'презентац']],
    ['Обеспечение / залог', ['залог', 'поручител', 'гарант', 'оценка имущества', 'отчет об оценке']],
    ['Анкета / заявка', ['анкета', 'заявка', 'заявление', 'опросник', 'карточка предприятия']],
    ['Методики и внутренние документы (не анализируются)', ['методик', 'положение о', 'регламент', 'приказ', 'политика', 'правила']]
  ];
  function classify(path, text) {
    const nameL = String(path).toLowerCase(); const head = String(text || '').slice(0, 5000).toLowerCase();
    for (const [k, terms] of KINDS) if (terms.some(t => nameL.includes(t))) return k;
    for (const [k, terms] of KINDS) if (terms.some(t => head.includes(t))) return k;
    return 'Не определён';
  }
  const USED_FOR = {
    'Бухгалтерский баланс (форма 1)': 'листы «Актив», «Пассив» → все коэффициенты', 'Отчет о финансовых результатах (форма 2)': 'лист «Приб_Уб» → рентабельность, оборачиваемость', 'Пояснения к отчетности / расшифровки': 'коды 5445/5540/5320 → текущая ликвидность',
    'Регистрационные документы / ЕГРЮЛ': 'реквизиты, ОКВЭД → отрасль, дата регистрации → длительность работы', 'Банковские выписки / обороты': 'вручную: поток средств, п. 7.2.12 (обороты не разбираются)', 'Картотека / справки банка': 'вручную: п. 7.2.6–7.2.7', 'Судебные дела': 'вручную: показатель № 8 бизнес-риска', 'Налоговая отчетность': 'сверка отчетности (п. 6.4)', 'Аудиторское заключение': 'достоверность отчетности (п. 6.3)'
  };

  /* ── Чтение одного файла → результат ── */
  async function readItem(item, opts) {
    opts = opts || {}; const ext = (item.name.split('.').pop() || '').toLowerCase(); const res = { name: item.name, path: item.path || item.name, size: item.size || (item.buf ? item.buf.byteLength : 0), ext, text: '', records: [], note: '', scanned: false, error: '' };
    const src = res.path;
    try {
      const buf = item.buf || await item.file.arrayBuffer(); item.buf = buf;
      if (['xlsx', 'xls', 'xlsm', 'xlsb', 'ods'].includes(ext)) {
        if (!root.XLSX) { res.note = 'Библиотека Excel не загрузилась (нужен интернет)'; return res; }
        const wb = XLSX.read(buf, { type: 'array', cellDates: true }); const r = importWorkbookObj(wb, { path: src, name: item.name }); res.records = r.records;
        res.text = wb.SheetNames.map(n => { try { return XLSX.utils.sheet_to_csv(wb.Sheets[n]); } catch (e) { return ''; } }).join('\n').slice(0, 250000); res.note = 'листов: ' + wb.SheetNames.length + (r.skipped.length ? '; расчетные листы пропущены: ' + r.skipped.length : '');
      } else if (ext === 'csv' || ext === 'txt' || ext === 'tsv') {
        let text; try { text = new TextDecoder('utf-8', { fatal: true }).decode(buf); } catch (e) { text = new TextDecoder('windows-1251').decode(buf); } res.text = text.slice(0, 300000);
        if (root.XLSX && (ext !== 'txt' || /[;\t]/.test(text.slice(0, 2000)))) { try { const wb = XLSX.read(text, { type: 'string', FS: text.includes(';') ? ';' : text.includes('\t') ? '\t' : ',' }); const g = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: '' }); res.records = importGrid(g, { src }); } catch (e) { } }
      } else if (ext === 'pdf') {
        if (!root.pdfjsLib) { res.note = 'Библиотека PDF не загрузилась (нужен интернет)'; return res; }
        const pdf = await pdfjsLib.getDocument({ data: buf.slice(0) }).promise; let t = ''; let carry = []; let units = null; const maxPages = Math.min(pdf.numPages, /бухгалтер|баланс|отчет|финанс|ф1|ф2/i.test(item.name) ? 40 : 12);
        for (let p = 1; p <= maxPages; p++) { const pg = await pdf.getPage(p), tc = await pg.getTextContent(); const r = parsePdfPage(tc.items, { src, page: p, units }, carry); if (r.anchors.length) carry = r.anchors; if (r.units && !units) units = r.units; res.records.push(...r.records); t += r.text + '\n'; }
        res.text = t; if (t.replace(/\s/g, '').length < 40) { res.scanned = true; res.note = 'скан без текстового слоя: нужна загрузка Excel/PDF с текстом или ручной ввод'; }
      } else if (ext === 'docx') {
        const d = await readDocx(buf); res.text = d.text; for (const g of d.grids) res.records.push(...importGrid(g, { src })); res.note = d.grids.length ? 'таблиц: ' + d.grids.length : '';
      } else if (ext === 'doc') { res.text = legacyDocText(buf); res.note = 'старый формат DOC: прочитан только текст (сохраните как DOCX/XLSX для таблиц)'; }
      else if (ext === 'rtf') { res.text = rtfText(new TextDecoder('windows-1251').decode(buf)).slice(0, 300000); }
      else if (ext === 'xml') { try { res.text = new TextDecoder('utf-8').decode(buf).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').slice(0, 300000); } catch (e) { } res.note = 'XML: взяты реквизиты и текст; суммы по тегам не разбираются'; }
      else if (['jpg', 'jpeg', 'png', 'tif', 'tiff', 'bmp', 'gif', 'webp'].includes(ext)) { res.scanned = true; res.note = 'изображение: текст не распознается (OCR отключен)'; }
      else res.note = 'формат не поддерживается';
    } catch (e) { res.error = String(e && e.message || e); }
    return res;
  }

  AZ.imp = { dropEmptyPeriods, importGrid, importWorkbookObj, parsePdfPage, readDocx, legacyDocText, rtfText, expandZip, classify, USED_FOR, readItem, rowCodeOf };
})(typeof window !== 'undefined' ? window : globalThis);
