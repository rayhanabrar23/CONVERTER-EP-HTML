/* PEI Porto — rekap portofolio & pendanaan per nasabah (format SUMM_TRX_PEI: sheet COMP + satu sheet per CID).
 * Alur harian: baca porto kemarin → tambah transaksi hari ini dari RP/LR Final → tulis workbook baru (rapi + formula).
 * Murni logika: workbook masuk/keluar berupa objek SheetJS, library penulis (xlsx-js-style) dioper dari luar. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.PEIPorto = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const FIRST = 4; // baris transaksi pertama di sheet nasabah (baris 3 = header)
  const HEAD = ['CLIENT ID', 'NAME', 'B/S', 'TRX DATE', 'DUE DATE', 'MATURITY', 'ACTIVITY', 'STOCK', 'HC', 'COLLATERAL (VOL)', 'PRICE',
    'COLLATERAL (IDR-HC)', 'AMOUNT TRX', 'FUNDING', 'OUTSTANDING', 'LN', 'INTEREST', 'RATIO', 'PEI', 'Days Left'];
  const COLS = 'ABCDEFGHIJKLMNOPQRST'.split('');
  const INPUT = new Set(['C', 'D', 'E', 'G', 'H', 'J', 'M', 'P']); // kolom isian; kolom lain berformula
  const SELL = ['S', 'DR', 'CR', 'R', 'D'];

  const str = (v) => (v == null ? '' : String(v)).trim();
  const num = (v) => (typeof v === 'number' ? v : (v == null || v === '' ? 0 : Number(String(v).replace(/,/g, '')) || 0));
  const cellOf = (ws, a) => (ws && ws[a]) || null;
  const valOf = (ws, a) => { const c = cellOf(ws, a); return c ? c.v : undefined; };

  // ── tanggal (serial Excel, tanpa zona waktu) ──
  const EPOCH = Date.UTC(1899, 11, 30);
  const toSerial = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - EPOCH) / 864e5);
  const fromSerial = (s) => { const t = new Date(EPOCH + s * 864e5); return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() }; };
  const isoOf = (s) => { const { y, m, d } = fromSerial(s); return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`; };
  const serialOfIso = (iso) => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(str(iso)); return m ? toSerial(+m[1], +m[2], +m[3]) : NaN; };
  /** n hari kerja setelah serial (Sabtu/Minggu dilewati; libur bursa tidak diketahui → bisa diubah manual) */
  function addWorkdays(s, n) {
    let t = s;
    while (n > 0) { t += 1; const wd = new Date(EPOCH + t * 864e5).getUTCDay(); if (wd !== 0 && wd !== 6) n -= 1; }
    return t;
  }

  // ── tranche: A..Z, AA..AZ, BA.. ──
  const trancheKey = (p) => [p.length, p];
  const cmpTranche = (a, b) => (a.length - b.length) || (a < b ? -1 : a > b ? 1 : 0);
  function nextTranche(used) {
    let last = ''; for (const p of used) if (/^[A-Z]+$/.test(p) && cmpTranche(p, last) > 0) last = p;
    if (!last) return 'A';
    const ch = last.split(''); let i = ch.length - 1;
    while (i >= 0 && ch[i] === 'Z') { ch[i] = 'A'; i -= 1; }
    if (i < 0) ch.unshift('A'); else ch[i] = String.fromCharCode(ch[i].charCodeAt(0) + 1);
    return ch.join('');
  }

  // ── baca porto ──
  const rateOf = (f, compRate) => {
    const m = /\*\s*(\d+(?:\.\d+)?)%\s*\/\s*360/.exec(f || '');
    if (m) return Number(m[1]);
    return /COMP!\$?M\$?1/i.test(f || '') && compRate != null ? Math.round(compRate * 1e6) / 1e4 : null;
  };
  function headerRow(ws) {
    for (let r = 1; r <= 8; r++) if (str(valOf(ws, 'A' + r)).toUpperCase() === 'CLIENT ID' && str(valOf(ws, 'C' + r)).toUpperCase() === 'B/S') return r;
    return 0;
  }
  function lastRow(ws) { const ref = ws && ws['!ref']; if (!ref) return 0; const m = /:[A-Z]+(\d+)$/.exec(ref) || /^[A-Z]+(\d+)$/.exec(ref); return m ? +m[1] : 0; }

  function parseClientSheet(ws, sheetName, compRate) {
    const hr = headerRow(ws); if (!hr) return null;
    let cid, name;
    if (str(valOf(ws, 'A1')).toUpperCase() === 'CLIENT ID') { cid = str(valOf(ws, 'B1')); name = str(valOf(ws, 'E1')); }
    else { cid = str(valOf(ws, 'A1')); name = str(valOf(ws, 'D1')); }
    cid = cid || sheetName; const code = cid.slice(0, 7);
    const rows = []; const end = lastRow(ws);
    for (let r = hr + 1; r <= end; r++) {
      if (str(valOf(ws, 'F' + r)) === cid && !str(valOf(ws, 'A' + r))) break; // baris total
      if (!str(valOf(ws, 'A' + r)).startsWith(code)) continue;
      const row = { man: {}, cache: {} };
      for (const c of COLS) {
        const cell = cellOf(ws, c + r); if (!cell) continue;
        if (cell.t === 'z' && !cell.f) continue; // stub kosong
        const has = cell.t !== 'z' && cell.v !== undefined && cell.v !== '';
        if (c === 'A') { row.A = str(cell.v); continue; }
        if (INPUT.has(c)) { if (has) row[c] = cell.t === 's' ? str(cell.v) : cell.v; continue; }
        if (cell.f) {
          if (has) row.cache[c] = cell.v; // nilai terhitung terakhir (tidak ada bila file belum pernah dibuka di Excel)
          const adj = /\)\s*([+-]\s*\d+(?:\.\d+)?)\s*$/.exec(cell.f); // koreksi manual di ujung formula, mis. "…))))+25000"
          if (adj) (row.adj = row.adj || {})[c] = adj[1].replace(/\s+/g, '');
        }
        else if (has && c !== 'B') row.man[c] = cell.v; // nilai diketik manual (mis. LR Final)
      }
      const q = cellOf(ws, 'Q' + r);
      row.rate = q && q.f ? rateOf(q.f, compRate) : null;
      rows.push(row);
    }
    return { cid, code, name, rows };
  }

  function parseComp(ws) {
    const out = { group: [], reguler: [], maxGroup: null, maxNonGroup: null, rate: null };
    if (!ws) return out;
    const m1 = valOf(ws, 'M1'); if (typeof m1 === 'number' && m1 > 0 && m1 < 1) out.rate = m1;
    let sec = 'group'; const end = lastRow(ws);
    for (let r = 1; r <= end; r++) {
      const b = str(valOf(ws, 'B' + r)).toUpperCase(), c = str(valOf(ws, 'C' + r)), d = str(valOf(ws, 'D' + r));
      const e = str(valOf(ws, 'E' + r)).toUpperCase(), f = str(valOf(ws, 'F' + r)).toUpperCase();
      if (b.startsWith('GROUP CLIENT')) { sec = 'group'; continue; }
      if (b.startsWith('REGUL')) { sec = 'reguler'; continue; }
      if (e.startsWith('MAX FUNDING') && f === 'GROUP' && out.maxGroup == null) { out.maxGroup = num(valOf(ws, 'G' + r)); continue; }
      if (e.startsWith('MAX FUNDING') && f === 'NON GROUP' && out.maxNonGroup == null) { out.maxNonGroup = num(valOf(ws, 'G' + r)); continue; }
      if (!c || c.toUpperCase() === 'CID' || !/^\d{4,}/.test(c)) continue;
      const pl = valOf(ws, 'M' + r);
      out[sec].push({ cid: c, name: d, flag: str(valOf(ws, 'L' + r)), note: str(valOf(ws, 'K' + r)), plafond: typeof pl === 'number' ? pl : 0, p: str(valOf(ws, 'P' + r)) });
    }
    return out;
  }

  /** wb: workbook SheetJS (dibaca dengan cellFormula:true, sheetStubs:true, tanpa cellDates). */
  function parsePorto(wb) {
    const comp = parseComp(wb.Sheets.COMP || wb.Sheets.COMPILE);
    const clients = new Map();
    for (const name of wb.SheetNames) {
      if (/^(COMP|COMPILE|HARGA|RISK P|REQ)$/i.test(name)) continue;
      const c = parseClientSheet(wb.Sheets[name], name, comp.rate);
      if (c) clients.set(c.cid, c);
    }
    if (!clients.size && !comp.group.length && !comp.reguler.length) throw new Error('Bukan file porto: tidak ada sheet COMP maupun sheet nasabah (header CLIENT ID / B/S).');
    // nasabah yang punya sheet tapi tidak tercantum di COMP → masuk reguler
    const listed = new Set([...comp.group, ...comp.reguler].map((x) => x.cid));
    for (const c of clients.values()) if (!listed.has(c.cid)) comp.reguler.push({ cid: c.cid, name: c.name, flag: '', note: '', plafond: 0, p: '' });
    return { comp, clients };
  }

  // ── nilai FUNDING per baris (cache Excel / manual / baris baru) ──
  function fundingOf(row) {
    if (row.man.N !== undefined) return num(row.man.N);
    if (row.cache.N !== undefined) return num(row.cache.N);
    if (row.newN !== undefined) return row.newN;
    return null;
  }
  function trancheBalances(client) {
    const bal = new Map(); const order = []; let unknown = 0;
    for (const row of client.rows) {
      const p = str(row.P); if (!p) continue;
      const n = fundingOf(row); if (n == null) { unknown += 1; continue; }
      if (!bal.has(p)) { bal.set(p, 0); order.push(p); }
      bal.set(p, bal.get(p) + n);
    }
    return { bal, order, unknown };
  }
  const outstandingOf = (client) => { let s = 0; for (const v of trancheBalances(client).bal.values()) s += v; return s; };
  const maxTrxDate = (model) => { let m = 0; for (const c of model.clients.values()) for (const r of c.rows) if (typeof r.D === 'number' && r.D > m) m = r.D; return m; };

  /** Bagi total Rupiah bulat ke beberapa bagian sebanding bobot; jumlah bagian = total persis. */
  function splitRupiah(total, weights) {
    const t = Math.round(total); const w = weights.map((x) => Math.max(num(x), 0)); const sw = w.reduce((a, b) => a + b, 0);
    if (!w.length) return [];
    if (sw <= 0) return w.map((_, i) => (i === w.length - 1 ? t : 0));
    const out = w.map((x) => Math.floor((t * x) / sw)); let rest = t - out.reduce((a, b) => a + b, 0);
    for (let i = 0; rest > 0; i = (i + 1) % out.length) { out[i] += 1; rest -= 1; }
    return out;
  }

  /** Tambahkan transaksi satu hari.
   * postings: [{ cid, name, rp: Rupiah, sells:[{stock, vol, value}], lr: Rupiah, buys:[{stock, vol, value}] }]
   * opts: { trxDate (serial), dueDate (serial), rate (persen, mis. 9.5) } */
  function applyDay(model, postings, opts) {
    const trx = opts.trxDate, due = opts.dueDate, rate = opts.rate;
    if (!(trx > 0) || !(due >= trx)) throw new Error('Tanggal transaksi / jatuh tempo tidak valid.');
    const last = maxTrxDate(model);
    if (last >= trx) throw new Error(`Porto yang diunggah sudah memuat transaksi tanggal ${isoOf(last)} (≥ ${isoOf(trx)}). Unggah porto hari sebelumnya.`);
    const log = [], warn = [];
    const listed = new Set([...model.comp.group, ...model.comp.reguler].map((x) => x.cid));
    for (const p of postings) {
      const known = model.clients.get(p.cid);
      const client = known || { cid: p.cid, code: p.cid.slice(0, 7), name: p.name || p.cid, rows: [], isNew: true };
      const before = client.rows.length;
      const base = { A: client.code, D: trx, E: due, man: {}, cache: {}, rate, isNew: true };
      // RP: pelunasan tranche terlama dulu (pokok dulu, sesuai aturan v5)
      if (p.rp > 0) {
        const { bal, order, unknown } = trancheBalances(client);
        if (unknown) warn.push(`${p.cid}: ${unknown} baris FUNDING tanpa nilai tersimpan — buka & simpan porto di Excel dulu agar sisa pinjaman per tranche terbaca.`);
        let rest = Math.round(p.rp); const pieces = [];
        for (const t of order) { // lunas penuh = saldo persis (termasuk sisa pecahan Rupiah), sebagian = Rupiah bulat
          const o = bal.get(t); if (Math.round(o) <= 0 || rest <= 0) continue;
          const x = rest >= Math.round(o) ? o : rest; pieces.push([t, x]); rest = Math.round(rest - x);
        }
        for (const [t, x] of pieces) client.rows.push(Object.assign({}, base, { C: 'S', G: 'RP', P: t, man: { N: -x }, newN: -x }));
        const paid = pieces.reduce((a, b) => a + b[1], 0);
        log.push({ cid: p.cid, kind: 'RP', text: `RP Final ${Math.round(p.rp)} → pokok ${paid} (${pieces.map(([t, x]) => t + ': ' + x).join(', ') || '-'})${rest > 0 ? `; sisa ${rest} = bunga/accrued` : ''}` });
        if (!pieces.length) warn.push(`${p.cid}: RP Final ${Math.round(p.rp)} tidak tercatat — tidak ada tranche terbuka di porto${known ? '' : ' (nasabah belum ada di porto kemarin)'}.`);
      }
      for (const s of p.sells || []) if (s.vol > 0) client.rows.push(Object.assign({}, base, { C: 'S', G: 'TRX', H: s.stock, J: -s.vol, M: -Math.round(s.value), man: {} }));
      // LR: satu tranche baru per nasabah per hari, FUNDING = LR Final (dibagi sebanding nilai beli)
      if (p.lr > 0 && (p.buys || []).length) {
        const t = nextTranche(client.rows.map((r) => str(r.P)));
        const parts = splitRupiah(p.lr, p.buys.map((b) => b.value));
        p.buys.forEach((b, i) => client.rows.push(Object.assign({}, base, { C: 'B', G: 'TRX', H: b.stock, J: b.vol, M: Math.round(b.value), P: t, man: { N: parts[i] }, newN: parts[i] })));
        log.push({ cid: p.cid, kind: 'LR', text: `LR Final ${Math.round(p.lr)} → tranche ${t} (${p.buys.map((b) => b.stock).join(', ')})` });
      }
      if (!known && client.rows.length > before) { // nasabah baru hanya dibuat bila ada baris yang tercatat
        model.clients.set(p.cid, client);
        if (!listed.has(p.cid)) { model.comp.reguler.push({ cid: p.cid, name: client.name, flag: '', note: '', plafond: 0, p: '', isNew: true }); listed.add(p.cid); }
        log.unshift({ cid: p.cid, kind: 'NASABAH BARU', text: `${client.name}: sheet & baris COMP baru (plafond 0 — isi manual)` });
      }
    }
    return { log, warn };
  }

  // ── formula (sama dengan SUMM_TRX_PEI) ──
  const F = {
    F: (r) => `IF(C${r}="B",IF(WEEKDAY(DATE(YEAR(E${r}),MONTH(E${r})+3,DAY(E${r})))>6,DATE(YEAR(E${r}),MONTH(E${r})+3,DAY(E${r}))-1,IF(WEEKDAY(DATE(YEAR(E${r}),MONTH(E${r})+3,DAY(E${r})))<2,DATE(YEAR(E${r}),MONTH(E${r})+3,DAY(E${r}))+1,DATE(YEAR(E${r}),MONTH(E${r})+3,DAY(E${r})))),"")`,
    B: (r) => `IFERROR(VLOOKUP(A${r},COMP!C:D,2,0),"")`,
    I: (r) => `IFERROR(VLOOKUP(H${r},HARGA!$A:$C,2,0),0)`,
    K: (r) => `IFERROR(VLOOKUP(H${r},HARGA!$A:$C,3,0),"")`,
    L: (r) => `IFERROR(J${r}*K${r}*(100-I${r})/100,0)`,
    N: (r) => {
      const p = r - 1; const S = `{"S","DR","D"}`;
      const sumM = `SUM(SUMIFS($M$3:M${r},$A$3:A${r},A${r},$C$3:C${r},${S},$E$3:E${r},E${r}))`;
      const prevN = `SUMIFS($N$3:N${p},A$3:A${p},A${r},$P$3:P${p},P${r})`, prevQ = `SUMIFS($Q$3:Q${p},A$3:A${p},A${r},$P$3:P${p},P${r})`;
      const sN = `ABS(SUM(SUMIFS($N$3:N${p},$A$3:A${p},A${r},$C$3:C${p},${S},$E$3:E${p},E${r})))`;
      const sQ = `ABS(SUM(SUMIFS($Q$3:Q${p},$A$3:A${p},A${r},$C$3:C${p},${S},$E$3:E${p},E${r})))`;
      return `IF(OR(C${r}="CD",C${r}="CW"),M${r},IF(P${r}="",0,IF(OR(C${r}="B",C${r}="SW",C${r}="SD"),M${r},IF(ABS(${sumM})<${prevN}+${prevQ}+${sN}+${sQ},IF(${sumM}+${prevQ}+${sN}+${sQ}>0,${sumM}+${sN}+${sQ},${sumM}+${prevQ}+${sN}+${sQ}),-${prevN}))))`;
    },
    O: (r) => `SUMIFS($N$3:N${r},$A$3:A${r},A${r},$P$3:P${r},P${r})`,
    Q: (r, tot, rate) => {
      const p = r - 1; const k = `${rate}%/360`;
      const cum = `SUMIFS($N$4:N${r},$A$4:$A${r},A${r},$P$4:$P${r},P${r})`;
      const seg = `IFERROR((INDEX($E${r + 1}:$E$${tot},MATCH(P${r},$P${r + 1}:$P$${tot},0),1)-E${r})*${cum}*${k},($E$${tot}-E${r})*${cum}*${k})`;
      const prevQ = `SUMIFS($Q$3:Q${p},$A$3:A${p},A${r},$P$3:P${p},P${r})`;
      return `IF($E$${tot}-E${r}<0,0,IF(OR(${prevQ}>ABS(N${r}),M${r}=N${r}),${seg},-SUMIFS($Q$3:Q${p},A$3:A${p},A${r},$P$3:P${p},P${r})+${seg}))+SUM(R${r})`;
    },
    R: (r) => `IF(H${r}="","",IF(SUMIFS(J:J,A:A,A${r},H:H,H${r})>=0,"","SHORT"))`,
    S: (r) => `IF(OR(C${r}="S",E${r}<=COMP!$M$2),"","N")`,
    T: (r) => `IF(OR(SUMIF(P:P,P${r},N:N)=0,F${r}="",F${r}<COMP!$M$2),"",IF((F${r}-COMP!$M$2)>30,"",IF((F${r}-COMP!$M$2)=0,"DUE",F${r}-COMP!$M$2)))`,
  };

  // ── gaya ──
  const RGB = { head: '8A1230', headTx: 'FFFFFF', sub: 'F8EFE6', total: 'FBE6EA', line: 'D9CFC2', manual: 'FFF3CD', newRow: 'EAF6EC', title: '740721' };
  const side = { style: 'thin', color: { rgb: RGB.line } };
  const BOX = { top: side, bottom: side, left: side, right: side };
  const font = (o) => Object.assign({ name: 'Arial', sz: 9 }, o || {});
  const ST = {
    title: { font: font({ sz: 14, bold: true, color: { rgb: RGB.title } }) },
    label: { font: font({ bold: true }) },
    head: { font: font({ bold: true, color: { rgb: RGB.headTx } }), fill: { fgColor: { rgb: RGB.head } }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: BOX },
    sub: { font: font({ bold: true, color: { rgb: RGB.head } }), fill: { fgColor: { rgb: RGB.sub } }, border: BOX },
    cell: { font: font(), border: BOX },
    total: { font: font({ bold: true }), fill: { fgColor: { rgb: RGB.total } }, border: BOX },
    input: { font: font({ bold: true, color: { rgb: '0000FF' } }), fill: { fgColor: { rgb: RGB.manual } }, border: BOX },
  };
  const Z = { money: '#,##0;(#,##0);"-"', vol: '#,##0;(#,##0);"-"', date: 'dd-mmm-yy', pct: '0.00%', hc: '0"%"' };
  const withFill = (s, rgb) => Object.assign({}, s, { fill: { fgColor: { rgb } } });

  function sheetWriter() {
    const ws = {}; let maxR = 1, maxC = 0;
    const put = (addr, v, s, z) => {
      const m = /^([A-Z]+)(\d+)$/.exec(addr); const c = m[1].split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1; const r = +m[2];
      maxR = Math.max(maxR, r); maxC = Math.max(maxC, c);
      const cell = {};
      if (v && typeof v === 'object' && 'f' in v) { cell.f = v.f; if (v.v !== undefined) { cell.v = v.v; cell.t = typeof v.v === 'number' ? 'n' : 's'; } else { cell.t = 'n'; } }
      else if (v === null || v === undefined || v === '') { cell.t = 's'; cell.v = ''; }
      else if (typeof v === 'number') { cell.t = 'n'; cell.v = v; }
      else { cell.t = 's'; cell.v = String(v); }
      if (s) cell.s = Object.assign({}, s); // salinan: penulis xlsx-js-style menambah numFmt ke objek gaya
      if (z) cell.z = z;
      ws[addr] = cell;
    };
    const finish = (cols) => { ws['!ref'] = `A1:${colName(maxC)}${maxR}`; if (cols) ws['!cols'] = cols.map((w) => ({ wch: w })); return ws; };
    return { ws, put, finish };
  }
  function colName(i) { let s = ''; i += 1; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }

  // ── tulis sheet nasabah ──
  function writeClient(client, asOf) {
    const { put, finish, ws } = sheetWriter();
    put('A1', client.cid, ST.title); put('D1', client.name, ST.title);
    put('A2', 'Posisi per', ST.label); put('D2', { f: 'COMP!$M$2' }, ST.label, Z.date);
    HEAD.forEach((h, i) => put(COLS[i] + 3, h, ST.head));
    const rows = client.rows; const n = rows.length; const last = FIRST + Math.max(n, 1) - 1; const tot = last + 1;
    const zOf = { D: Z.date, E: Z.date, F: Z.date, I: Z.hc, J: Z.vol, K: Z.money, L: Z.money, M: Z.money, N: Z.money, O: Z.money, Q: Z.money };
    const running = new Map();
    rows.forEach((row, i) => {
      const r = FIRST + i; const base = row.isNew ? withFill(ST.cell, RGB.newRow) : ST.cell;
      const nVal = fundingOf(row); const p = str(row.P);
      if (p && nVal != null) running.set(p, (running.get(p) || 0) + nVal);
      for (const c of COLS) {
        if (c === 'A') { put('A' + r, row.A || client.code, base); continue; }
        if (INPUT.has(c)) { put(c + r, row[c] === undefined ? '' : row[c], base, zOf[c]); continue; }
        if (row.man[c] !== undefined) { put(c + r, row.man[c], ST.input, zOf[c]); continue; } // manual dipertahankan (ditandai kuning)
        let f;
        if (c === 'Q') f = row.rate == null ? null : F.Q(r, tot, row.rate);
        else f = F[c](r);
        if (f == null) { put(c + r, row.cache[c] === undefined ? '' : row.cache[c], ST.input, zOf[c]); continue; }
        if (row.adj && row.adj[c]) f = `${f}${row.adj[c]}`;
        const cache = c === 'N' ? (nVal == null ? undefined : nVal) : c === 'O' && p && nVal != null ? running.get(p) : undefined;
        put(c + r, { f, v: cache }, base, zOf[c]);
      }
    });
    if (!n) COLS.forEach((c) => put(c + FIRST, '', ST.cell));
    // total
    const rng = (c) => `${c}${FIRST}:${c}${last}`;
    for (const c of COLS) put(c + tot, '', ST.total);
    const lastDue = `LOOKUP(2,1/(NOT(ISBLANK(E$${FIRST}:$E${last}))),E$${FIRST}:$E${last})`;
    put('E' + tot, { f: `IFERROR(IF(COMP!$M$2<=${lastDue},${lastDue},COMP!$M$2),COMP!$M$2)` }, ST.total, Z.date);
    put('F' + tot, { f: 'A1' }, ST.total); put('G' + tot, { f: 'D1' }, ST.total);
    for (const c of ['J', 'L', 'N', 'Q']) put(c + tot, { f: `SUM(${rng(c)})` }, ST.total, zOf[c]);
    put('O' + tot, 'LAST LOAN', ST.total);
    put('P' + tot, { f: `IFERROR(LOOKUP(2,1/(COUNTIF(P$${FIRST}:P$${last},">"&P$${FIRST}:P$${last}&"*")=0),P$${FIRST}:P$${last}),"")` }, ST.total);
    put('R' + tot, { f: `IF(ABS(N${tot})<1,0,IF(L${tot}=0,"NO COLL",(N${tot}+Q${tot})/L${tot}))` }, ST.total, Z.pct);
    // blok COLLATERAL SAHAM (posisi sekarang)
    const vol = new Map(); for (const row of rows) { const h = str(row.H); if (h) vol.set(h, (vol.get(h) || 0) + num(row.J)); }
    const stocks = [...vol].filter(([, v]) => Math.abs(v) > 0).map(([s]) => s).sort();
    let r = tot + 2;
    put('H' + r, 'COLLATERAL SAHAM (posisi saat ini)', ST.label); r += 1;
    ['STOCK', 'HC', 'COLLATERAL (VOL)', 'PRICE', 'COLLATERAL (IDR-HC)'].forEach((h, i) => put('HIJKL'[i] + r, h, ST.head)); r += 1;
    const c0 = r;
    for (const s of stocks) {
      put('H' + r, s, ST.cell); put('I' + r, { f: F.I(r) }, ST.cell, Z.hc);
      put('J' + r, { f: `SUMIF(H$${FIRST}:H$${last},H${r},J$${FIRST}:J$${last})` }, ST.cell, Z.vol);
      put('K' + r, { f: F.K(r) }, ST.cell, Z.money); put('L' + r, { f: F.L(r) }, ST.cell, Z.money); r += 1;
    }
    put('H' + r, 'PORTOFOLIO', ST.total); ['I', 'J', 'K'].forEach((c) => put(c + r, '', ST.total));
    put('L' + r, { f: stocks.length ? `SUM(L${c0}:L${r - 1})` : '0' }, ST.total, Z.money); const portRow = r;
    // blok PELUNASAN FUNDING (tranche yang masih terbuka)
    const { bal, order } = trancheBalances(client);
    const open = order.filter((t) => Math.abs(bal.get(t)) >= 1);
    r += 2; put('H' + r, 'PELUNASAN FUNDING (tranche terbuka)', ST.label); r += 1;
    ['TRANCHE', 'TGL PINJAM', 'JATUH TEMPO', 'OUTSTANDING', 'INTEREST'].forEach((h, i) => put('HIJKL'[i] + r, h, ST.head)); r += 1;
    const t0 = r;
    for (const t of open) {
      put('H' + r, t, ST.cell);
      put('I' + r, { f: `IFERROR(INDEX($D$${FIRST}:$D$${last},MATCH(H${r},$P$${FIRST}:$P$${last},0)),"")` }, ST.cell, Z.date);
      put('J' + r, { f: `IFERROR(INDEX($F$${FIRST}:$F$${last},MATCH(H${r},$P$${FIRST}:$P$${last},0)),"")` }, ST.cell, Z.date);
      put('K' + r, { f: `SUMIF($P$${FIRST}:$P$${last},H${r},$N$${FIRST}:$N$${last})` }, ST.cell, Z.money);
      put('L' + r, { f: `SUMIF($P$${FIRST}:$P$${last},H${r},$Q$${FIRST}:$Q$${last})` }, ST.cell, Z.money); r += 1;
    }
    put('H' + r, 'TOTAL', ST.total); ['I', 'J'].forEach((c) => put(c + r, '', ST.total));
    put('K' + r, { f: open.length ? `SUM(K${t0}:K${r - 1})` : '0' }, ST.total, Z.money);
    put('L' + r, { f: open.length ? `SUM(L${t0}:L${r - 1})` : '0' }, ST.total, Z.money);
    finish([10, 26, 5, 10, 10, 10, 12, 12, 10, 14, 14, 16, 16, 16, 16, 6, 14, 9, 5, 8]);
    ws['!rows'] = [{}, {}, { hpt: 28 }];
    return { ws, tot, portRow };
  }

  // ── COMP ──
  const CH = ['No.', 'CID', 'Client', 'Original Funding', 'Repayment', 'Outstanding', 'Accrued Interest', 'Portofolio', 'Ratio', 'Note', '', 'PLAFOND', 'USAGE', 'SHORTFALL'];
  const q = (cid) => `'${String(cid).replace(/'/g, "''")}'`;
  function writeComp(model, sheetInfo, opts) {
    const { put, finish } = sheetWriter();
    put('B1', 'LIST FUNDING PEI (Internal only)', ST.title);
    put('K1', 'Bunga transaksi baru (p.a.)', ST.label); put('M1', opts.rate / 100, ST.input, Z.pct);
    put('K2', 'Posisi per', ST.label); put('M2', opts.asOf, ST.input, Z.date);
    let r = 3; const totals = {};
    const section = (title, list, key) => {
      put('B' + r, title, ST.label); r += 1;
      CH.forEach((h, i) => { if (h) put(colName(i + 1) + r, h, ST.head); }); r += 1;
      const r0 = r;
      list.forEach((x, i) => {
        const info = sheetInfo.get(x.cid); const s = x.isNew ? withFill(ST.cell, RGB.newRow) : ST.cell;
        put('B' + r, i + 1, s); put('C' + r, x.cid, s); put('D' + r, x.name, s);
        if (info) {
          const sh = q(x.cid);
          put('E' + r, { f: `SUM(SUMIFS(${sh}!$N:$N,${sh}!$C:$C,{"B","L"}))` }, s, Z.money);
          put('F' + r, { f: `SUM(SUMIFS(${sh}!$N:$N,${sh}!$C:$C,{"S","DR","CR","R","D"}))` }, s, Z.money);
          put('G' + r, { f: `${sh}!$N$${info.tot}` }, s, Z.money);
          put('H' + r, { f: `${sh}!$Q$${info.tot}` }, s, Z.money);
          put('I' + r, { f: `${sh}!$L$${info.tot}` }, s, Z.money);
          put('J' + r, { f: `IF(ABS(G${r})<1,0,${sh}!$R$${info.tot})` }, s, Z.pct);
        } else for (const c of 'EFGHI') put(c + r, 0, s, Z.money), put('J' + r, 0, s, Z.pct);
        put('K' + r, x.note || '', s); put('L' + r, x.flag || '', s);
        put('M' + r, x.plafond || 0, x.isNew ? withFill(ST.input, RGB.manual) : ST.input, Z.money);
        put('N' + r, { f: `IF(M${r}=0,0,(H${r}+G${r})/M${r})` }, s, Z.pct);
        put('O' + r, { f: `IF(ABS(G${r})<1,0,IF(AND(ISNUMBER(J${r}),J${r}<65%),0,MAX(G${r}+H${r}-65%*I${r},0)))` }, s, Z.money);
        put('P' + r, x.p || '', s); r += 1;
      });
      const r1 = Math.max(r - 1, r0);
      if (!list.length) r += 1;
      for (const c of 'BCDEFGHIJKLMNO') put(c + r, '', ST.total);
      put('D' + r, 'TOTAL', ST.total);
      for (const c of 'EFGHIMO') put(c + r, { f: `SUM(${c}${r0}:${c}${r1})` }, ST.total, Z.money);
      put('N' + r, { f: `IF(M${r}=0,0,(H${r}+G${r})/M${r})` }, ST.total, Z.pct);
      totals[key] = r; r += 2;
    };
    section('GROUP CLIENT', model.comp.group, 'g');
    const blk = (label, max, tr) => {
      put('E' + r, 'MAX FUNDING', ST.label); put('F' + r, label, ST.label); put('G' + r, max || 0, ST.input, Z.money); r += 1;
      put('F' + r, 'TOTAL FUNDING', ST.label); put('G' + r, { f: `G${tr}+H${tr}` }, ST.cell, Z.money); r += 1;
      put('F' + r, 'FREE', ST.label); put('G' + r, { f: `G${r - 2}-G${r - 1}` }, ST.cell, Z.money); r += 1;
      put('F' + r, 'USAGE', ST.label); put('G' + r, { f: `IF(G${r - 3}=0,0,G${r - 2}/G${r - 3})` }, ST.cell, Z.pct); r += 2;
    };
    blk('GROUP', model.comp.maxGroup, totals.g);
    section('REGULER CLIENT', model.comp.reguler, 'n');
    blk('NON GROUP', model.comp.maxNonGroup, totals.n);
    return finish([4, 6, 12, 32, 17, 17, 17, 15, 17, 9, 14, 4, 17, 9, 17, 4]);
  }

  function writeHarga(prices) {
    const { put, finish } = sheetWriter();
    ['STOCK', 'HC (%)', 'PRICE'].forEach((h, i) => put('ABC'[i] + 1, h, ST.head));
    prices.forEach((p, i) => { put('A' + (i + 2), p.stock, ST.cell); put('B' + (i + 2), p.hc, ST.cell, '0.##'); put('C' + (i + 2), p.price, ST.cell, Z.money); });
    return finish([10, 8, 12]);
  }

  /** XS: library penulis (xlsx-js-style). prices: [{stock, hc(%), price}] dari Closing Price + Risk Parameter. */
  function buildWorkbook(model, XS, opts) {
    const wb = XS.utils.book_new(); const info = new Map(); const sheets = [];
    for (const x of [...model.comp.group, ...model.comp.reguler]) {
      const c = model.clients.get(x.cid); if (!c || info.has(x.cid)) continue;
      const w = writeClient(c, opts.asOf); info.set(x.cid, w); sheets.push([x.cid, w.ws]);
    }
    XS.utils.book_append_sheet(wb, writeComp(model, info, opts), 'COMP');
    XS.utils.book_append_sheet(wb, writeHarga(opts.prices || []), 'HARGA');
    for (const [name, ws] of sheets) XS.utils.book_append_sheet(wb, ws, name.slice(0, 31));
    return wb;
  }

  /** Data harga dari konteks PEICore: Closing Price + haircut Risk Parameter (dalam %). */
  function pricesFromCtx(ctx) {
    const all = new Set([...ctx.cp.keys(), ...ctx.hc.keys()]);
    return [...all].sort().map((s) => ({ stock: s, hc: ctx.hc.has(s) ? Math.round(ctx.hc.get(s) * 10000) / 100 : 0, price: ctx.cp.has(s) ? ctx.cp.get(s) : 0 }));
  }

  /** Posting harian dari hasil validasi (sumber kebenaran = yang masuk file iFast). */
  function postingsFromResults(vres, ctx, lolosRp, lolosLr) {
    const out = [], skipped = [];
    for (const [sid, d] of vres) {
      if (!d.is_pei) continue;
      const rpOk = lolosRp(d) && d.rp_final > 0, lrOk = lolosLr(d) && d.lr_final > 0;
      if (!rpOk && !lrOk) continue;
      const cid = ctx.sidToCid.get(sid);
      if (!cid) { skipped.push(sid); continue; }
      const net = ctx.netting.get(sid) || {};
      const amt = (stock) => Math.abs(net[stock] ? net[stock].net_value : 0);
      out.push({
        cid, sid, name: d.name,
        rp: rpOk ? d.rp_final : 0,
        sells: rpOk ? d.rp_detail.filter((x) => x.ada_di_op && x.lot_keluar > 0).map((x) => ({ stock: x.stock, vol: x.lot_keluar, value: amt(x.stock) * (x.lot_sell ? x.lot_keluar / x.lot_sell : 1) })) : [],
        lr: lrOk ? d.lr_final : 0,
        buys: lrOk ? Object.entries(d.buy_stocks).map(([stock, b]) => ({ stock, vol: b.lot, value: amt(stock) || b.value })) : [],
      });
    }
    return { postings: out, skipped };
  }

  /** Bandingkan outstanding porto dengan Loan Existing OP (per SID) sebelum transaksi hari ini. */
  function compareWithOp(model, ctx) {
    const diffs = [];
    for (const [sid, op] of ctx.op) {
      const cid = ctx.sidToCid.get(sid); if (!cid) continue;
      const c = model.clients.get(cid); const o = c ? outstandingOf(c) : 0;
      if (Math.abs(o - op.loan_existing) >= 1) diffs.push({ sid, cid, porto: o, op: op.loan_existing });
    }
    return diffs;
  }

  return { parsePorto, applyDay, buildWorkbook, pricesFromCtx, postingsFromResults, compareWithOp, outstandingOf, trancheBalances, nextTranche, splitRupiah,
    addWorkdays, toSerial, isoOf, serialOfIso, maxTrxDate, F, HEAD, FIRST };
});
