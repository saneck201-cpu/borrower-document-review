/* xw.js — компактный генератор XLSX для браузера (без внешних библиотек).
 *
 * Возможности: стили, формулы с КЭШИРОВАННЫМИ значениями (считает встроенный
 * вычислитель), объединения, проверки данных (выпадающие списки), условное
 * форматирование, внутренние гиперссылки, закрепление областей, ширина колонок,
 * диаграммы (столбчатые / линейные), печать по ширине страницы.
 *
 * Кэшированные значения нужны, чтобы книга показывала числа и в «Защищённом
 * просмотре» Excel, в превью почты и на телефоне, где формулы не пересчитываются.
 */
(function (root) {
  'use strict';

  /* ───────────── утилиты ───────────── */
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const clean = s => String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '');
  const X = s => esc(clean(s));

  function colName(i) { let s = ''; for (i++; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + (i - 1) % 26) + s; return s; }
  function colIndex(name) { let n = 0; for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; }
  const addr = (r, c, absR, absC) => (absC ? '$' : '') + colName(c) + (absR ? '$' : '') + (r + 1); // r,c — 0-based
  function parseAddr(a) { const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(a); if (!m) return null; return { r: Number(m[2]) - 1, c: colIndex(m[1]) }; }
  const quoteSheet = n => (/^[A-Za-zА-Яа-яЁё_][A-Za-zА-Яа-яЁё0-9_]*$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'");

  /* ───────────── стили ───────────── */
  const BUILTIN_NF = { 'General': 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10 };
  class Styles {
    constructor() {
      this.fonts = []; this.fills = []; this.borders = []; this.numFmts = []; this.xfs = []; this.dxfs = [];
      this.fontMap = new Map(); this.fillMap = new Map(); this.borderMap = new Map(); this.nfMap = new Map(); this.xfMap = new Map(); this.dxfMap = new Map();
      this.fonts.push('<font><sz val="10"/><color rgb="FF1F2A26"/><name val="Calibri"/><family val="2"/></font>'); this.fontMap.set('{}', 0);
      this.fills.push('<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>');
      this.borders.push('<border><left/><right/><top/><bottom/><diagonal/></border>'); this.borderMap.set('{}', 0);
      this.xfs.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'); this.xfMap.set('{}', 0);
    }
    font(f) {
      const k = JSON.stringify(f || {}); if (this.fontMap.has(k)) return this.fontMap.get(k);
      const x = '<font>' + (f.b ? '<b/>' : '') + (f.i ? '<i/>' : '') + (f.u ? '<u/>' : '') + `<sz val="${f.sz || 10}"/><color rgb="${f.color || 'FF1F2A26'}"/><name val="${f.name || 'Calibri'}"/><family val="2"/></font>`;
      this.fonts.push(x); this.fontMap.set(k, this.fonts.length - 1); return this.fonts.length - 1;
    }
    fill(rgb) {
      if (!rgb) return 0; if (this.fillMap.has(rgb)) return this.fillMap.get(rgb);
      this.fills.push(`<fill><patternFill patternType="solid"><fgColor rgb="${rgb}"/><bgColor indexed="64"/></patternFill></fill>`);
      this.fillMap.set(rgb, this.fills.length - 1); return this.fills.length - 1;
    }
    border(b) {
      if (!b) return 0; const k = JSON.stringify(b); if (this.borderMap.has(k)) return this.borderMap.get(k);
      const side = (n, v) => { if (!v) return `<${n}/>`; const st = typeof v === 'string' ? v : v.s || 'thin'; const col = (typeof v === 'object' && v.color) || 'FFBFCBC5'; return `<${n} style="${st}"><color rgb="${col}"/></${n}>`; };
      const all = b.all; const x = '<border>' + side('left', b.l || all) + side('right', b.r || all) + side('top', b.t || all) + side('bottom', b.b || all) + '<diagonal/></border>';
      this.borders.push(x); this.borderMap.set(k, this.borders.length - 1); return this.borders.length - 1;
    }
    numFmt(code) {
      if (!code) return 0; if (code in BUILTIN_NF) return BUILTIN_NF[code]; if (this.nfMap.has(code)) return this.nfMap.get(code);
      const id = 164 + this.numFmts.length; this.numFmts.push(`<numFmt numFmtId="${id}" formatCode="${esc(code)}"/>`); this.nfMap.set(code, id); return id;
    }
    xf(s) {
      s = s || {}; const k = JSON.stringify(s); if (this.xfMap.has(k)) return this.xfMap.get(k);
      const fontId = this.font({ b: s.b, i: s.i, u: s.u, sz: s.sz, color: s.color, name: s.fontName });
      const fillId = this.fill(s.fill), borderId = this.border(s.bd), nf = this.numFmt(s.nf);
      const al = (s.h || s.v || s.wrap || s.indent) ? `<alignment${s.h ? ` horizontal="${s.h}"` : ''}${s.v ? ` vertical="${s.v}"` : ' vertical="center"'}${s.wrap ? ' wrapText="1"' : ''}${s.indent ? ` indent="${s.indent}"` : ''}/>` : '';
      const x = `<xf numFmtId="${nf}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${nf ? ' applyNumberFormat="1"' : ''} applyFont="1"${fillId ? ' applyFill="1"' : ''}${borderId ? ' applyBorder="1"' : ''}${al ? ' applyAlignment="1">' + al + '</xf>' : '/>'}`;
      this.xfs.push(x); this.xfMap.set(k, this.xfs.length - 1); return this.xfs.length - 1;
    }
    dxf(d) {
      const k = JSON.stringify(d); if (this.dxfMap.has(k)) return this.dxfMap.get(k);
      const x = '<dxf>' + ((d.color || d.b || d.i) ? `<font>${d.b ? '<b/>' : ''}${d.i ? '<i/>' : ''}${d.color ? `<color rgb="${d.color}"/>` : ''}</font>` : '') + (d.fill ? `<fill><patternFill patternType="solid"><fgColor rgb="${d.fill}"/><bgColor rgb="${d.fill}"/></patternFill></fill>` : '') + '</dxf>';
      this.dxfs.push(x); this.dxfMap.set(k, this.dxfs.length - 1); return this.dxfs.length - 1;
    }
    xml() {
      return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        (this.numFmts.length ? `<numFmts count="${this.numFmts.length}">${this.numFmts.join('')}</numFmts>` : '') +
        `<fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts><fills count="${this.fills.length}">${this.fills.join('')}</fills><borders count="${this.borders.length}">${this.borders.join('')}</borders>` +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        `<cellXfs count="${this.xfs.length}">${this.xfs.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
        `<dxfs count="${this.dxfs.length}">${this.dxfs.join('')}</dxfs></styleSheet>`;
    }
  }

  /* ───────────── вычислитель формул (подмножество Excel) ───────────── */
  const ERR = (e) => ({ err: e });
  const isErr = v => v && typeof v === 'object' && 'err' in v;
  function roundHalfAway(x, n) { const f = Math.pow(10, n); const s = x < 0 ? -1 : 1; return s * Math.round(Math.abs(x) * f + 1e-9 * (Math.abs(x) * f > 1e6 ? 0 : 1)) / f; }
  function xlRound(x, n) {
    // «как в Excel»: округление половины от нуля, с поправкой на двоичную погрешность
    if (!Number.isFinite(x)) return x; const f = Math.pow(10, n); const y = Math.abs(x) * f; const r = Math.round(Number(y.toPrecision(15)));
    return (x < 0 ? -1 : 1) * r / f;
  }

  function tokenize(src) {
    const t = []; let i = 0; const n = src.length;
    while (i < n) {
      const ch = src[i];
      if (/\s/.test(ch)) { i++; continue; }
      if (ch === '"') { let j = i + 1, s = ''; while (j < n) { if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; } s += src[j++]; } t.push({ k: 'str', v: s }); i = j + 1; continue; }
      if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(src[i + 1] || ''))) { let j = i; while (j < n && /[0-9.]/.test(src[j])) j++; if (/[eE]/.test(src[j] || '') && /[0-9+\-]/.test(src[j + 1] || '')) { j++; if (/[+\-]/.test(src[j])) j++; while (j < n && /[0-9]/.test(src[j])) j++; } t.push({ k: 'num', v: Number(src.slice(i, j)) }); i = j; continue; }
      // ссылка с именем листа
      let m = /^'((?:[^']|'')+)'!/.exec(src.slice(i));
      let sheet = null, len = 0;
      if (m) { sheet = m[1].replace(/''/g, "'"); len = m[0].length; }
      else { m = /^([A-Za-zА-Яа-яЁё_][A-Za-zА-Яа-яЁё0-9_.]*)!/.exec(src.slice(i)); if (m) { sheet = m[1]; len = m[0].length; } }
      const rest = src.slice(i + len);
      const rm = /^(\$?[A-Za-z]{1,3}\$?\d+)(?::(\$?[A-Za-z]{1,3}\$?\d+))?/.exec(rest);
      if (rm && (sheet !== null || !/^[A-Za-z]{1,3}\d+\(/.test(rest))) {
        // не путать с функцией вроде LOG10(
        const after = rest[rm[0].length];
        if (!(after && /[A-Za-zА-Яа-я_(]/.test(after))) { t.push({ k: 'ref', sheet, a: rm[1], b: rm[2] || null }); i += len + rm[0].length; continue; }
      }
      if (/[A-Za-zА-Яа-я_]/.test(ch)) { let j = i; while (j < n && /[A-Za-zА-Яа-я0-9_.]/.test(src[j])) j++; t.push({ k: 'id', v: src.slice(i, j).toUpperCase() }); i = j; continue; }
      const two = src.slice(i, i + 2);
      if (two === '<>' || two === '<=' || two === '>=') { t.push({ k: 'op', v: two }); i += 2; continue; }
      if ('+-*/^&=<>(),%'.includes(ch)) { t.push({ k: 'op', v: ch }); i++; continue; }
      throw new Error('Неизвестный символ в формуле: ' + ch + ' [' + src + ']');
    }
    return t;
  }

  function parseFormula(src) {
    const toks = tokenize(src); let p = 0;
    const peek = () => toks[p], next = () => toks[p++];
    const isOp = v => peek() && peek().k === 'op' && peek().v === v;
    function expr() { return cmp(); }
    function cmp() { let l = concat(); while (peek() && peek().k === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(peek().v)) { const op = next().v; l = { t: 'bin', op, l, r: concat() }; } return l; }
    function concat() { let l = add(); while (isOp('&')) { next(); l = { t: 'bin', op: '&', l, r: add() }; } return l; }
    function add() { let l = mul(); while (peek() && peek().k === 'op' && (peek().v === '+' || peek().v === '-')) { const op = next().v; l = { t: 'bin', op, l, r: mul() }; } return l; }
    function mul() { let l = pow(); while (peek() && peek().k === 'op' && (peek().v === '*' || peek().v === '/')) { const op = next().v; l = { t: 'bin', op, l, r: pow() }; } return l; }
    function pow() { let l = unary(); while (isOp('^')) { next(); l = { t: 'bin', op: '^', l, r: unary() }; } return l; }
    function unary() { if (isOp('-')) { next(); return { t: 'neg', e: unary() }; } if (isOp('+')) { next(); return unary(); } return post(); }
    function post() { let e = primary(); while (isOp('%')) { next(); e = { t: 'bin', op: '/', l: e, r: { t: 'num', v: 100 } }; } return e; }
    function primary() {
      const tk = next(); if (!tk) throw new Error('Неожиданный конец формулы: ' + src);
      if (tk.k === 'num') return { t: 'num', v: tk.v };
      if (tk.k === 'str') return { t: 'str', v: tk.v };
      if (tk.k === 'ref') return { t: 'ref', sheet: tk.sheet, a: tk.a, b: tk.b };
      if (tk.k === 'id') {
        if (isOp('(')) { next(); const args = []; if (!isOp(')')) { for (;;) { if (isOp(',') || isOp(')')) args.push({ t: 'empty' }); else args.push(expr()); if (isOp(',')) { next(); continue; } break; } } if (!isOp(')')) throw new Error('Ожидалась ) в: ' + src); next(); return { t: 'fn', name: tk.v, args }; }
        if (tk.v === 'TRUE') return { t: 'bool', v: true }; if (tk.v === 'FALSE') return { t: 'bool', v: false };
        throw new Error('Неизвестный идентификатор ' + tk.v + ' в: ' + src);
      }
      if (tk.k === 'op' && tk.v === '(') { const e = expr(); if (!isOp(')')) throw new Error('Ожидалась ) в: ' + src); next(); return e; }
      throw new Error('Неожиданный токен ' + JSON.stringify(tk) + ' в: ' + src);
    }
    const e = expr(); if (p < toks.length) throw new Error('Лишние символы в формуле: ' + src);
    return e;
  }

  const toNum = v => { if (isErr(v)) return v; if (v === null || v === undefined || v === '') return v === '' ? ERR('#VALUE!') : 0; if (typeof v === 'number') return v; if (typeof v === 'boolean') return v ? 1 : 0; if (typeof v === 'string') { const n = Number(v.replace(',', '.')); return v.trim() !== '' && Number.isFinite(n) ? n : ERR('#VALUE!'); } return ERR('#VALUE!'); };
  const toStr = v => { if (isErr(v)) return v; if (v === null || v === undefined) return ''; if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'; if (typeof v === 'number') return String(Number(v.toPrecision(15))); return String(v); };
  const typeRank = v => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
  function compare(a, b) {
    if (a === null || a === undefined) a = (typeof b === 'string') ? '' : (typeof b === 'boolean' ? false : 0);
    if (b === null || b === undefined) b = (typeof a === 'string') ? '' : (typeof a === 'boolean' ? false : 0);
    if (typeRank(a) !== typeRank(b)) return typeRank(a) - typeRank(b);
    if (typeof a === 'string') { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
    return a < b ? -1 : a > b ? 1 : 0;
  }

  class Evaluator {
    constructor(wb) { this.wb = wb; this.cache = new Map(); this.stack = new Set(); this.parsed = new Map(); }
    cellValue(sheet, r, c) {
      const sh = sheet; const cell = sh.cells.get(r + ',' + c);
      if (!cell) return null;
      if (cell.f === undefined || cell.f === null) return cell.v === undefined ? null : cell.v;
      const key = sh.name + '!' + r + ',' + c;
      if (this.cache.has(key)) return this.cache.get(key);
      if (this.stack.has(key)) return ERR('#REF!');
      this.stack.add(key);
      let val;
      try {
        let ast = this.parsed.get(cell.f); if (!ast) { ast = parseFormula(cell.f); this.parsed.set(cell.f, ast); }
        val = this.ev(ast, sh);
        if (Array.isArray(val)) val = val[0][0]; // диапазон на месте скаляра
      } catch (e) { val = ERR('#NAME?'); this.lastError = e; }
      this.stack.delete(key);
      if (val === undefined) val = null;
      this.cache.set(key, val); return val;
    }
    rangeVals(refNode, sh) {
      const target = refNode.sheet ? this.wb.sheetByName(refNode.sheet) : sh;
      if (!target) return ERR('#REF!');
      const a = parseAddr(refNode.a), b = refNode.b ? parseAddr(refNode.b) : a;
      const r1 = Math.min(a.r, b.r), r2 = Math.max(a.r, b.r), c1 = Math.min(a.c, b.c), c2 = Math.max(a.c, b.c);
      const rows = []; for (let r = r1; r <= r2; r++) { const row = []; for (let c = c1; c <= c2; c++) row.push(this.cellValue(target, r, c)); rows.push(row); }
      return rows;
    }
    scalar(v) { if (Array.isArray(v)) return v[0][0] === undefined ? null : v[0][0]; return v; }
    flat(args, sh) { const out = []; for (const a of args) { const v = a.t === 'ref' ? this.rangeVals(a, sh) : this.ev(a, sh); if (Array.isArray(v)) for (const row of v) for (const x of row) out.push({ v: x, ref: true }); else out.push({ v, ref: false }); } return out; }
    ev(n, sh) {
      switch (n.t) {
        case 'num': return n.v; case 'str': return n.v; case 'bool': return n.v; case 'empty': return null;
        case 'ref': { const v = this.rangeVals(n, sh); if (isErr(v)) return v; if (!n.b) return v[0][0]; return v; }
        case 'neg': { const v = toNum(this.scalar(this.ev(n.e, sh))); return isErr(v) ? v : -v; }
        case 'bin': {
          const l = this.scalar(this.ev(n.l, sh)), r = this.scalar(this.ev(n.r, sh));
          if (isErr(l)) return l; if (isErr(r)) return r;
          if (n.op === '&') return toStr(l) + toStr(r);
          if (['=', '<>', '<', '>', '<=', '>='].includes(n.op)) { const c = compare(l, r); return n.op === '=' ? c === 0 : n.op === '<>' ? c !== 0 : n.op === '<' ? c < 0 : n.op === '>' ? c > 0 : n.op === '<=' ? c <= 0 : c >= 0; }
          const a = toNum(l), b = toNum(r); if (isErr(a)) return a; if (isErr(b)) return b;
          switch (n.op) { case '+': return a + b; case '-': return a - b; case '*': return a * b; case '/': return b === 0 ? ERR('#DIV/0!') : a / b; case '^': return Math.pow(a, b); }
          return ERR('#VALUE!');
        }
        case 'fn': return this.fn(n, sh);
      }
      return ERR('#VALUE!');
    }
    fn(n, sh) {
      const name = n.name, args = n.args;
      const sc = i => this.scalar(this.ev(args[i], sh));
      switch (name) {
        case 'IF': { const c = sc(0); if (isErr(c)) return c; const t = typeof c === 'string' ? ERR('#VALUE!') : (c === null ? false : (typeof c === 'number' ? c !== 0 : c)); if (isErr(t)) return t; if (t) return args[1] ? this.scalar(this.ev(args[1], sh)) : true; return args[2] ? this.scalar(this.ev(args[2], sh)) : false; }
        case 'IFERROR': { const v = sc(0); if (isErr(v)) return this.scalar(this.ev(args[1], sh)); return v; }
        case 'ISNUMBER': { const v = sc(0); return typeof v === 'number'; }
        case 'ISBLANK': { const v = sc(0); return v === null; }
        case 'ISERROR': return isErr(sc(0));
        case 'AND': case 'OR': { const xs = this.flat(args, sh); let res = name === 'AND'; let any = false; for (const x of xs) { if (isErr(x.v)) return x.v; if (x.v === null || (x.ref && typeof x.v === 'string')) continue; let b; if (typeof x.v === 'string') return ERR('#VALUE!'); b = typeof x.v === 'number' ? x.v !== 0 : x.v; any = true; if (name === 'AND') res = res && b; else res = res || b; } return any ? res : ERR('#VALUE!'); }
        case 'NOT': { const v = sc(0); if (isErr(v)) return v; return !(typeof v === 'number' ? v !== 0 : v); }
        case 'ABS': { const v = toNum(sc(0)); return isErr(v) ? v : Math.abs(v); }
        case 'ROUND': { const v = toNum(sc(0)), d = toNum(sc(1)); if (isErr(v)) return v; if (isErr(d)) return d; return xlRound(v, d); }
        case 'SUM': case 'MAX': case 'MIN': case 'AVERAGE': case 'COUNT': {
          const xs = this.flat(args, sh); const nums = [];
          for (const x of xs) { if (isErr(x.v)) { if (name !== 'COUNT') return x.v; continue; } if (typeof x.v === 'number') nums.push(x.v); else if (!x.ref && name !== 'COUNT') { if (typeof x.v === 'boolean') nums.push(x.v ? 1 : 0); else if (typeof x.v === 'string') { const q = toNum(x.v); if (isErr(q)) return q; nums.push(q); } } }
          if (name === 'COUNT') return nums.length; if (name === 'SUM') return nums.reduce((a, b) => a + b, 0);
          if (!nums.length) return name === 'AVERAGE' ? ERR('#DIV/0!') : 0;
          if (name === 'MAX') return Math.max(...nums); if (name === 'MIN') return Math.min(...nums); return nums.reduce((a, b) => a + b, 0) / nums.length;
        }
        case 'COUNTIF': { const rng = this.ev(args[0], sh); const crit = this.scalar(this.ev(args[1], sh)); if (isErr(rng)) return rng; let cnt = 0; const rows = Array.isArray(rng) ? rng : [[rng]]; for (const row of rows) for (const v of row) { if (typeof crit === 'string') { if (typeof v === 'string' && v.toLowerCase() === crit.toLowerCase()) cnt++; } else if (typeof crit === 'number') { if (v === crit) cnt++; } } return cnt; }
        case 'VLOOKUP': { const key = sc(0); const tbl = this.ev(args[1], sh); const col = toNum(sc(2)); if (isErr(key)) return key; if (!Array.isArray(tbl)) return ERR('#N/A'); for (const row of tbl) { if (compare(row[0], key) === 0) return row[col - 1] === undefined ? ERR('#REF!') : row[col - 1]; } return ERR('#N/A'); }
        case 'INDEX': { const tbl = this.ev(args[0], sh); const r = toNum(sc(1)); const c = args[2] ? toNum(sc(2)) : 1; if (!Array.isArray(tbl)) return tbl; const rr = tbl.length === 1 && !args[2] ? 1 : r; const cc = tbl.length === 1 && !args[2] ? r : c; const row = tbl[rr - 1]; if (!row || row[cc - 1] === undefined) return ERR('#REF!'); return row[cc - 1]; }
        case 'MATCH': { const key = sc(0); const rng = this.ev(args[1], sh); if (!Array.isArray(rng)) return ERR('#N/A'); const flat = rng.flat(); for (let i = 0; i < flat.length; i++) if (compare(flat[i], key) === 0) return i + 1; return ERR('#N/A'); }
        case 'LEN': { const v = sc(0); return isErr(v) ? v : toStr(v).length; }
        case 'CONCATENATE': { let s = ''; for (let i = 0; i < args.length; i++) { const v = sc(i); if (isErr(v)) return v; s += toStr(v); } return s; }
        default: return ERR('#NAME?');
      }
    }
  }

  /* ───────────── Лист / Книга ───────────── */
  class Sheet {
    constructor(wb, name, opts) {
      this.wb = wb; this.name = name; this.cells = new Map(); this.merges = []; this.colW = new Map(); this.rowH = new Map(); this.validations = []; this.cfs = []; this.links = [];
      this.freezeAt = null; this.tab = (opts && opts.tab) || null; this.grid = !(opts && opts.grid === false); this.zoom = (opts && opts.zoom) || 100; this.landscape = true; this.charts = []; this.printTitles = null; this.hiddenCols = new Set();
      this.maxR = 0; this.maxC = 0; this.defaultStyle = null; this.autofilter = null;
    }
    /** value: number | string | null | {f:'формула'}; style: объект стиля */
    set(r, c, value, style) {
      const cell = { s: style ? this.wb.styles.xf(style) : 0 };
      if (value && typeof value === 'object' && value.f !== undefined) { cell.f = String(value.f).replace(/^=/, ''); }
      else if (value === null || value === undefined || value === '') cell.v = null;
      else if (typeof value === 'number') cell.v = Number.isFinite(value) ? value : null;
      else if (typeof value === 'boolean') cell.v = value;
      else cell.v = String(value);
      this.cells.set(r + ',' + c, cell); if (r > this.maxR) this.maxR = r; if (c > this.maxC) this.maxC = c; return cell;
    }
    style(r, c, style) { const k = r + ',' + c; const cell = this.cells.get(k); if (cell) cell.s = this.wb.styles.xf(style); else this.set(r, c, null, style); }
    get(r, c) { return this.wb.evaluator().cellValue(this, r, c); }
    merge(r1, c1, r2, c2) { this.merges.push([r1, c1, r2, c2]); }
    width(c, w) { this.colW.set(c, w); }
    height(r, h) { this.rowH.set(r, h); }
    freeze(r, c) { this.freezeAt = { r, c }; }
    validate(range, spec) { this.validations.push({ range, ...spec }); }
    cf(range, rule) { this.cfs.push({ range, rule }); }
    link(r, c, location, tooltip) { this.links.push({ r, c, location, tooltip }); }
    addChart(spec) { this.charts.push(spec); }
    ref(r, c, absR, absC) { return quoteSheet(this.name) + '!' + addr(r, c, absR, absC); }
  }

  class Workbook {
    constructor(props) { this.sheets = []; this.styles = new Styles(); this._ev = null; this.props = props || {}; this.definedNames = []; }
    addSheet(name, opts) { if (name.length > 31 || /[\[\]:*?\/\\]/.test(name)) throw new Error('Недопустимое имя листа: ' + name); const s = new Sheet(this, name, opts); this.sheets.push(s); this._ev = null; return s; }
    sheetByName(n) { return this.sheets.find(s => s.name.toLowerCase() === String(n).toLowerCase()); }
    evaluator() { if (!this._ev) this._ev = new Evaluator(this); return this._ev; }
    invalidate() { this._ev = null; }
    /** значение ячейки после расчёта */
    value(sheetName, r, c) { const sh = this.sheetByName(sheetName); return sh ? this.evaluator().cellValue(sh, r, c) : null; }

    /* ── сериализация ── */
    async toBlob() {
      const files = this.files(); const zipped = await zip(files);
      return new Blob([zipped], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    }
    files() {
      const ev = this.evaluator(); this.evalErrors = [];
      const sst = []; const sstMap = new Map(); const sid = s => { if (!sstMap.has(s)) { sstMap.set(s, sst.length); sst.push(s); } return sstMap.get(s); };
      const out = []; const ct = []; const wbRels = []; const sheetNodes = []; const drawingsFor = [];
      let chartNo = 0; const chartFiles = [];
      this.sheets.forEach((sh, si) => {
        const n = si + 1; let rowsXml = '';
        const byRow = new Map(); for (const [k, cell] of sh.cells) { const [r, c] = k.split(',').map(Number); if (!byRow.has(r)) byRow.set(r, []); byRow.get(r).push([c, cell]); }
        const rowKeys = [...new Set([...byRow.keys(), ...sh.rowH.keys()])].sort((a, b) => a - b);
        for (const r of rowKeys) {
          const cs = (byRow.get(r) || []).sort((a, b) => a[0] - b[0]); const ht = sh.rowH.get(r);
          let cx = '';
          for (const [c, cell] of cs) {
            const ref = colName(c) + (r + 1), s = cell.s ? ` s="${cell.s}"` : '';
            if (cell.f !== undefined) {
              let v = ev.cellValue(sh, r, c);
              let t = '', vx = '';
              if (isErr(v)) { t = ' t="e"'; vx = `<v>${v.err}</v>`; this.evalErrors.push(sh.name + '!' + ref + ' ' + v.err); }
              else if (typeof v === 'number') { vx = `<v>${Number.isFinite(v) ? String(v) : 0}</v>`; }
              else if (typeof v === 'boolean') { t = ' t="b"'; vx = `<v>${v ? 1 : 0}</v>`; }
              else if (typeof v === 'string') { t = ' t="str"'; vx = `<v>${X(v)}</v>`; }
              else { t = ' t="str"'; vx = '<v></v>'; }
              cx += `<c r="${ref}"${s}${t}><f>${X(cell.f)}</f>${vx}</c>`;
            } else if (cell.v === null || cell.v === undefined) cx += `<c r="${ref}"${s}/>`;
            else if (typeof cell.v === 'number') cx += `<c r="${ref}"${s}><v>${cell.v}</v></c>`;
            else if (typeof cell.v === 'boolean') cx += `<c r="${ref}"${s} t="b"><v>${cell.v ? 1 : 0}</v></c>`;
            else cx += `<c r="${ref}"${s} t="s"><v>${sid(clean(cell.v))}</v></c>`;
          }
          rowsXml += `<row r="${r + 1}"${ht ? ` ht="${ht}" customHeight="1"` : ''}>${cx}</row>`;
        }
        const dim = `A1:${colName(Math.max(0, sh.maxC))}${Math.max(1, sh.maxR + 1)}`;
        let pane = '';
        if (sh.freezeAt && (sh.freezeAt.r || sh.freezeAt.c)) {
          const { r, c } = sh.freezeAt; const tl = colName(c) + (r + 1);
          const act = r && c ? 'bottomRight' : r ? 'bottomLeft' : 'topRight';
          pane = `<pane${c ? ` xSplit="${c}"` : ''}${r ? ` ySplit="${r}"` : ''} topLeftCell="${tl}" activePane="${act}" state="frozen"/>` +
            (r && c ? '<selection pane="topRight"/><selection pane="bottomLeft"/>' : '') + `<selection pane="${act}" activeCell="${tl}" sqref="${tl}"/>`;
        }
        const cols = [...sh.colW.keys()].sort((a, b) => a - b).map(c => `<col min="${c + 1}" max="${c + 1}" width="${sh.colW.get(c)}" customWidth="1"${sh.hiddenCols.has(c) ? ' hidden="1"' : ''}/>`).join('');
        const merges = sh.merges.length ? `<mergeCells count="${sh.merges.length}">${sh.merges.map(m => `<mergeCell ref="${colName(m[1])}${m[0] + 1}:${colName(m[3])}${m[2] + 1}"/>`).join('')}</mergeCells>` : '';
        let prio = 1;
        const cfx = sh.cfs.map(({ range, rule }) => {
          let inner = '';
          if (rule.type === 'cellIs') inner = `<cfRule type="cellIs" dxfId="${this.styles.dxf(rule.dxf)}" priority="${prio++}" operator="${rule.op}">${rule.formula.map(f => `<formula>${X(f)}</formula>`).join('')}</cfRule>`;
          else if (rule.type === 'expression') inner = `<cfRule type="expression" dxfId="${this.styles.dxf(rule.dxf)}" priority="${prio++}"><formula>${X(rule.formula)}</formula></cfRule>`;
          else if (rule.type === 'colorScale') inner = `<cfRule type="colorScale" priority="${prio++}"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="${rule.min || 'FFFFFFFF'}"/><color rgb="${rule.max || 'FF63BE7B'}"/></colorScale></cfRule>`;
          else if (rule.type === 'dataBar') inner = `<cfRule type="dataBar" priority="${prio++}"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="${rule.color || 'FF7CB59A'}"/></dataBar></cfRule>`;
          return `<conditionalFormatting sqref="${range}">${inner}</conditionalFormatting>`;
        }).join('');
        const dvs = sh.validations.length ? `<dataValidations count="${sh.validations.length}">${sh.validations.map(v => {
          const list = v.list ? `<formula1>"${esc(v.list.join(','))}"</formula1>` : v.formula ? `<formula1>${X(v.formula)}</formula1>` : '';
          return `<dataValidation type="${v.type || 'list'}" allowBlank="1" showInputMessage="1" showErrorMessage="${v.strict === false ? 0 : 1}" sqref="${v.range}"${v.title ? ` promptTitle="${X(v.title).slice(0, 32)}"` : ''}${v.prompt ? ` prompt="${X(v.prompt).slice(0, 255)}"` : ''}${v.error ? ` errorTitle="Недопустимое значение" error="${X(v.error).slice(0, 255)}"` : ''}>${list}</dataValidation>`;
        }).join('')}</dataValidations>` : '';
        const links = sh.links.length ? `<hyperlinks>${sh.links.map(l => `<hyperlink ref="${colName(l.c)}${l.r + 1}" location="${X(l.location)}" display="${X(l.tooltip || '')}"/>`).join('')}</hyperlinks>` : '';
        let drawing = '';
        if (sh.charts.length) {
          const dn = drawingsFor.length + 1; drawingsFor.push({ sheet: sh, dn, n });
          drawing = '<drawing r:id="rId1"/>';
        }
        const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
          `<sheetPr>${sh.tab ? `<tabColor rgb="${sh.tab}"/>` : ''}<pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="${dim}"/>` +
          `<sheetViews><sheetView${sh.grid ? '' : ' showGridLines="0"'}${si === 0 ? ' tabSelected="1"' : ''} zoomScale="${sh.zoom}" zoomScaleNormal="${sh.zoom}" workbookViewId="0">${pane}</sheetView></sheetViews>` +
          '<sheetFormatPr defaultRowHeight="15" customHeight="0"/>' + (cols ? `<cols>${cols}</cols>` : '') + `<sheetData>${rowsXml}</sheetData>` +
          (sh.autofilter ? `<autoFilter ref="${sh.autofilter}"/>` : '') + merges + cfx + dvs + links +
          '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>' + `<pageSetup paperSize="9" orientation="${sh.landscape ? 'landscape' : 'portrait'}" fitToWidth="1" fitToHeight="0"/>` + drawing + '</worksheet>';
        out.push([`xl/worksheets/sheet${n}.xml`, xml]);
        ct.push(`<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
        wbRels.push(`<Relationship Id="rId${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${n}.xml"/>`);
        sheetNodes.push(`<sheet name="${X(sh.name)}" sheetId="${n}" r:id="rId${n}"/>`);
      });
      // диаграммы
      for (const d of drawingsFor) {
        const { sheet, dn, n } = d; let anchors = '', rels = '';
        sheet.charts.forEach((ch, ci) => {
          chartNo++; chartFiles.push([`xl/charts/chart${chartNo}.xml`, chartXml(ch, this)]);
          ct.push(`<Override PartName="/xl/charts/chart${chartNo}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);
          rels += `<Relationship Id="rId${ci + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartNo}.xml"/>`;
          const a = ch.anchor; // {c1,r1,c2,r2}
          anchors += `<xdr:twoCellAnchor><xdr:from><xdr:col>${a.c1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${a.r1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${a.c2}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${a.r2}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
            `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${ci + 2}" name="Диаграмма ${ci + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId${ci + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
        });
        out.push([`xl/drawings/drawing${dn}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${anchors}</xdr:wsDr>`]);
        out.push([`xl/drawings/_rels/drawing${dn}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`]);
        out.push([`xl/worksheets/_rels/sheet${n}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${dn}.xml"/></Relationships>`]);
        ct.push(`<Override PartName="/xl/drawings/drawing${dn}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
      }
      out.push(...chartFiles);
      const names = this.sheets.map((sh, i) => sh.printTitles ? `<definedName name="_xlnm.Print_Titles" localSheetId="${i}">${X(sh.printTitles)}</definedName>` : '').join('') + this.definedNames.map(d => `<definedName name="${X(d.name)}">${X(d.ref)}</definedName>`).join('');
      const wbXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr defaultThemeVersion="124226"/><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="15000" activeTab="0"/></bookViews>' +
        `<sheets>${sheetNodes.join('')}</sheets>${names ? `<definedNames>${names}</definedNames>` : ''}<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`;
      const n1 = this.sheets.length;
      const files = [
        ['[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>${ct.join('')}</Types>`],
        ['_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'],
        ['docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${X(this.props.title || 'Анализ финансового положения заемщика')}</dc:title><dc:creator>${X(this.props.creator || 'Анализ заемщика')}</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().replace(/\.\d+Z$/, 'Z')}</dcterms:created></cp:coreProperties>`],
        ['docProps/app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>'],
        ['xl/workbook.xml', wbXml],
        ['xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${wbRels.join('')}<Relationship Id="rId${n1 + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId${n1 + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`],
        ['xl/styles.xml', this.styles.xml()],
        ['xl/sharedStrings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sst.length}" uniqueCount="${sst.length}">${sst.map(s => `<si><t xml:space="preserve">${esc(s)}</t></si>`).join('')}</sst>`],
        ...out
      ];
      return files;
    }
  }

  /* ───────────── диаграммы ───────────── */
  function chartXml(ch, wb) {
    // ch: {type:'col'|'colStacked'|'colPct'|'line', title, cats:{sheet, r1,c1,r2,c2}, series:[{name:{sheet,r,c}|string, vals:{sheet,r1,c1,r2,c2}, color}], yFmt, anchor}
    const sh = n => wb.sheetByName(n);
    const rangeRef = rg => `${quoteSheet(rg.sheet)}!${addr(rg.r1, rg.c1, true, true)}${(rg.r1 !== rg.r2 || rg.c1 !== rg.c2) ? ':' + addr(rg.r2, rg.c2, true, true) : ''}`;
    const cells = rg => { const out = []; for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) out.push(wb.value(rg.sheet, r, c)); return out; };
    const strCache = arr => `<c:strCache><c:ptCount val="${arr.length}"/>${arr.map((v, i) => `<c:pt idx="${i}"><c:v>${X(v === null ? '' : toStr(v))}</c:v></c:pt>`).join('')}</c:strCache>`;
    const numCache = arr => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${arr.length}"/>${arr.map((v, i) => typeof v === 'number' ? `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>` : '').join('')}</c:numCache>`;
    const palette = ['FF176B4E', 'FF4F9FD1', 'FFE0A030', 'FFB5527F', 'FF7A8A83', 'FF6F6BC0', 'FFC96A3E', 'FF3F8F8F', 'FF9AB84A', 'FF8E5BA6'];
    const isLine = ch.type === 'line';
    const cats = ch.cats ? `<c:cat><c:strRef><c:f>${X(rangeRef(ch.cats))}</c:f>${strCache(cells(ch.cats))}</c:strRef></c:cat>` : '';
    const ser = ch.series.map((s, i) => {
      const color = s.color || palette[i % palette.length];
      const nm = typeof s.name === 'string' ? `<c:tx><c:v>${X(s.name)}</c:v></c:tx>` : `<c:tx><c:strRef><c:f>${X(quoteSheet(s.name.sheet) + '!' + addr(s.name.r, s.name.c, true, true))}</c:f>${strCache([wb.value(s.name.sheet, s.name.r, s.name.c)])}</c:strRef></c:tx>`;
      const sp = isLine ? `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${color.slice(2)}"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="${color.slice(2)}"/></a:solidFill></c:spPr></c:marker>` : `<c:spPr><a:solidFill><a:srgbClr val="${color.slice(2)}"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/>`;
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${nm}${sp}${cats}<c:val><c:numRef><c:f>${X(rangeRef(s.vals))}</c:f>${numCache(cells(s.vals))}</c:numRef></c:val>${isLine ? '<c:smooth val="0"/>' : ''}</c:ser>`;
    }).join('');
    const grouping = ch.type === 'colStacked' ? 'stacked' : ch.type === 'colPct' ? 'percentStacked' : 'clustered';
    const plot = isLine ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${ser}<c:marker val="1"/><c:axId val="111"/><c:axId val="222"/></c:lineChart>`
      : `<c:barChart><c:barDir val="col"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${ser}<c:gapWidth val="60"/>${grouping !== 'clustered' ? '<c:overlap val="100"/>' : ''}<c:axId val="111"/><c:axId val="222"/></c:barChart>`;
    const txPr = sz => `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${sz}"><a:solidFill><a:srgbClr val="445851"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="ru-RU"/></a:p></c:txPr>`;
    const title = ch.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="1"/></a:pPr><a:r><a:rPr lang="ru-RU" sz="1200" b="1"><a:solidFill><a:srgbClr val="17352C"/></a:solidFill></a:rPr><a:t>${X(ch.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/>' +
      `<c:chart>${title}<c:plotArea><c:layout/>${plot}` +
      `<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="low"/>${txPr(900)}<c:crossAx val="222"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
      `<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="E1E8E4"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="${esc(ch.yFmt || '#,##0')}" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${txPr(900)}<c:crossAx val="111"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>` +
      `</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/>${txPr(900)}</c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
  }

  /* ───────────── ZIP (deflate через CompressionStream, иначе без сжатия) ───────────── */
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1); t[n] = c >>> 0; } return t; })();
  function crc32(b) { let c = 0xffffffff; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }
  async function deflateRaw(bytes) {
    if (typeof CompressionStream === 'undefined') return null;
    try {
      const cs = new CompressionStream('deflate-raw'); const w = cs.writable.getWriter(); w.write(bytes); w.close();
      const buf = await new Response(cs.readable).arrayBuffer(); return new Uint8Array(buf);
    } catch (e) { return null; }
  }
  async function zip(entries) {
    const enc = new TextEncoder(); const locals = [], centrals = []; let offset = 0;
    const now = new Date(); const dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1); const dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
    for (const [name, content] of entries) {
      const nb = enc.encode(name); const raw = typeof content === 'string' ? enc.encode(content) : content; const crc = crc32(raw);
      let data = raw, method = 0; const d = raw.length > 200 ? await deflateRaw(raw) : null; if (d && d.length < raw.length) { data = d; method = 8; }
      const lh = new Uint8Array(30 + nb.length), lv = new DataView(lh.buffer);
      lv.setUint32(0, 0x04034b50, true); lv.setUint16(4, 20, true); lv.setUint16(6, 0x0800, true); lv.setUint16(8, method, true); lv.setUint16(10, dosTime, true); lv.setUint16(12, dosDate, true); lv.setUint32(14, crc, true); lv.setUint32(18, data.length, true); lv.setUint32(22, raw.length, true); lv.setUint16(26, nb.length, true); lh.set(nb, 30);
      locals.push(lh, data);
      const ch = new Uint8Array(46 + nb.length), cv = new DataView(ch.buffer);
      cv.setUint32(0, 0x02014b50, true); cv.setUint16(4, 20, true); cv.setUint16(6, 20, true); cv.setUint16(8, 0x0800, true); cv.setUint16(10, method, true); cv.setUint16(12, dosTime, true); cv.setUint16(14, dosDate, true); cv.setUint32(16, crc, true); cv.setUint32(20, data.length, true); cv.setUint32(24, raw.length, true); cv.setUint16(28, nb.length, true); cv.setUint32(42, offset, true); ch.set(nb, 46);
      centrals.push(ch); offset += lh.length + data.length;
    }
    const csize = centrals.reduce((a, b) => a + b.length, 0), end = new Uint8Array(22), ev = new DataView(end.buffer);
    ev.setUint32(0, 0x06054b50, true); ev.setUint16(8, entries.length, true); ev.setUint16(10, entries.length, true); ev.setUint32(12, csize, true); ev.setUint32(16, offset, true);
    const total = offset + csize + 22; const outBuf = new Uint8Array(total); let p = 0; for (const part of [...locals, ...centrals, end]) { outBuf.set(part, p); p += part.length; }
    return outBuf;
  }

  root.XW = { Workbook, colName, colIndex, addr, parseAddr, quoteSheet, isErr, xlRound, zip, crc32, parseFormula };
})(typeof window !== 'undefined' ? window : globalThis);
