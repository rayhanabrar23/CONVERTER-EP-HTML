// Tes porto harian (public/porto.js). Jalankan: node tests/porto.test.js
const fs = require('fs'), path = require('path'), assert = require('assert');
const X = require('xlsx'), XS = require('xlsx-js-style');
const C = require('../public/core.js'), P = require('../public/porto.js');
const D = path.join(__dirname, 'data');
const xrows = (f) => { const w = X.readFile(D + '/' + f); return X.utils.sheet_to_json(w.Sheets[w.SheetNames[0]]); };
const txt = (f) => fs.readFileSync(D + '/' + f, 'utf8');
const ctx = C.buildContext({ closingRows: xrows('closing.xlsx'), riskText: txt('risk.txt'), opText: txt('op.txt'), nettingRows: xrows('netting.xlsx'),
  clText: txt('cl.txt'), sidClientRows: xrows('sidclient.xlsx') });
let n = 0; const t = (name, fn) => { fn(); n++; console.log('  ✓', name); };
const roundTrip = (wb) => X.read(XS.write(wb, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer', cellFormula: true, sheetStubs: true });
const empty = () => ({ comp: { group: [], reguler: [], maxGroup: null, maxNonGroup: null, rate: null }, clients: new Map() });
const D1 = P.toSerial(2026, 10, 26), D2 = P.toSerial(2026, 10, 30);
const prices = P.pricesFromCtx(ctx);

t('tranche berikutnya: A → B, Z → AA, AZ → BA, kosong → A', () => {
  assert.strictEqual(P.nextTranche([]), 'A'); assert.strictEqual(P.nextTranche(['A', '', 'C', 'B']), 'D');
  assert.strictEqual(P.nextTranche(['Z', 'Y']), 'AA'); assert.strictEqual(P.nextTranche(['AZ', 'Z']), 'BA'); assert.strictEqual(P.nextTranche(['ZZ']), 'AAA');
});
t('LR Final dibagi ke saham sebanding nilai beli, jumlahnya tepat (Rupiah bulat)', () => {
  const p = P.splitRupiah(1000001, [3, 3, 3]); assert.strictEqual(p.reduce((a, b) => a + b, 0), 1000001); assert.ok(p.every(Number.isInteger));
  assert.deepStrictEqual(P.splitRupiah(500, [0, 0]), [0, 500]);
});
t('jatuh tempo T+2 hari kerja (Jumat → Selasa)', () => { const fri = P.toSerial(2026, 10, 23); assert.strictEqual(P.isoOf(P.addWorkdays(fri, 2)), '2026-10-27'); });

// hari 1: dua kali LR (tranche A lalu B di hari berikutnya), hari 2: RP melunasi tranche terlama dulu
let model = empty();
t('hari 1: LR Final → nasabah baru, satu tranche, FUNDING = LR Final (manual), harga dari HARGA', () => {
  const r = P.applyDay(model, [{ cid: '0000001', name: 'BUDI', lr: 1000000, buys: [{ stock: 'BBCA', vol: 100, value: 900000 }, { stock: 'TLKM', vol: 200, value: 300000 }], rp: 0, sells: [] }],
    { trxDate: D1, dueDate: D1 + 2, rate: 9.5 });
  assert.ok(r.log.some((x) => x.kind === 'NASABAH BARU')); assert.strictEqual(model.comp.reguler[0].cid, '0000001');
  const rows = model.clients.get('0000001').rows; assert.strictEqual(rows.length, 2);
  assert.deepStrictEqual(rows.map((x) => x.P), ['A', 'A']); assert.deepStrictEqual(rows.map((x) => x.man.N), [750000, 250000]);
  const wb = roundTrip(P.buildWorkbook(model, XS, { asOf: D1, rate: 9.5, prices }));
  assert.deepStrictEqual(wb.SheetNames.slice(0, 3), ['COMP', 'HARGA', '0000001']);
  const ws = wb.Sheets['0000001']; assert.strictEqual(ws.N4.v, 750000); assert.ok(!ws.N4.f, 'LR Final = nilai, bukan formula');
  assert.match(ws.Q4.f, /9\.5%\/360/); assert.match(ws.K4.f, /HARGA!/); assert.strictEqual(ws.C3.v, 'B/S');
  const comp = wb.Sheets.COMP; const cr = Object.keys(comp).find((a) => /^C\d+$/.test(a) && comp[a].v === '0000001').slice(1);
  assert.match(comp['G' + cr].f, /'0000001'!\$N\$6/); assert.strictEqual(comp['M' + cr].v, 0, 'plafond nasabah baru 0 (isi manual)');
  model = P.parsePorto(wb); assert.strictEqual(P.outstandingOf(model.clients.get('0000001')), 1000000, 'nilai FUNDING terbaca lagi dari file');
});
t('hari 1b: LR lagi → tranche B; porto tanggal yang sama/lebih lama ditolak', () => {
  assert.throws(() => P.applyDay(model, [], { trxDate: D1, dueDate: D1 + 2, rate: 9.5 }), /sudah memuat transaksi tanggal 2026-10-26/);
  P.applyDay(model, [{ cid: '0000001', name: 'BUDI', lr: 500000, buys: [{ stock: 'BBCA', vol: 50, value: 600000 }], rp: 0, sells: [] }], { trxDate: D1 + 1, dueDate: D1 + 3, rate: 9.5 });
  const rows = model.clients.get('0000001').rows; assert.strictEqual(rows[2].P, 'B');
  model = P.parsePorto(roundTrip(P.buildWorkbook(model, XS, { asOf: D1 + 1, rate: 9.5, prices })));
  assert.strictEqual(P.outstandingOf(model.clients.get('0000001')), 1500000);
});
t('hari 2: RP Final melunasi tranche terlama dulu (A penuh, B sebagian); kelebihan = bunga; saham keluar', () => {
  const r = P.applyDay(model, [{ cid: '0000001', name: 'BUDI', rp: 1200000, sells: [{ stock: 'BBCA', vol: 100, value: 1300000 }], lr: 0, buys: [] }], { trxDate: D2, dueDate: D2 + 2, rate: 9.5 });
  const c = model.clients.get('0000001'); const rp = c.rows.filter((x) => x.G === 'RP');
  assert.deepStrictEqual(rp.map((x) => [x.P, x.man.N]), [['A', -1000000], ['B', -200000]]);
  assert.ok(c.rows.some((x) => x.G === 'TRX' && x.C === 'S' && x.H === 'BBCA' && x.J === -100));
  const bal = P.trancheBalances(c).bal; assert.strictEqual(bal.get('A'), 0); assert.strictEqual(bal.get('B'), 300000);
  const r2 = P.applyDay(P.parsePorto(roundTrip(P.buildWorkbook(model, XS, { asOf: D2, rate: 9.5, prices }))),
    [{ cid: '0000001', name: 'BUDI', rp: 999999, sells: [], lr: 0, buys: [] }], { trxDate: D2 + 1, dueDate: D2 + 3, rate: 9.5 });
  assert.match(r2.log.find((x) => x.kind === 'RP').text, /sisa 699999 = bunga/); void r;
});
t('porto lama: tarif per baris & koreksi manual di ujung formula dipertahankan; nilai manual tetap nilai', () => {
  const ws = { '!ref': 'A1:T8', A1: { t: 's', v: '0000009' }, D1: { t: 's', v: 'ANI' } };
  P.HEAD.forEach((hd, i) => { ws['ABCDEFGHIJKLMNOPQRST'[i] + 3] = { t: 's', v: hd }; });
  const put = (a, v, f) => { ws[a] = typeof v === 'number' ? { t: 'n', v } : { t: 's', v }; if (f) ws[a].f = f; };
  put('A4', '0000009'); put('C4', 'B'); put('D4', D1 - 30); put('E4', D1 - 28); put('H4', 'BBCA'); put('J4', 10); put('M4', 1000); put('P4', 'A');
  put('N4', 1000, P.F.N(4)); put('Q4', 5, P.F.Q(4, 6, 9).replace(/9%/g, '9.25%'));
  put('A5', '0000009'); put('C5', 'S'); put('D5', D1 - 2); put('E5', D1); put('H5', 'BBCA'); put('J5', -10); put('M5', -1100); put('P5', 'A');
  put('N5', -1000, P.F.N(5) + '+25000'); put('Q5', -7);
  put('F6', '0000009');
  const wbIn = { SheetNames: ['COMP', '0000009'], Sheets: { COMP: { '!ref': 'A1:A1' }, '0000009': ws } };
  const m = P.parsePorto(wbIn); const rows = m.clients.get('0000009').rows;
  assert.strictEqual(rows[0].rate, 9.25); assert.strictEqual(rows[1].adj.N, '+25000'); assert.strictEqual(rows[1].man.Q, -7);
  const out = roundTrip(P.buildWorkbook(m, XS, { asOf: D1, rate: 9.5, prices })).Sheets['0000009'];
  assert.match(out.Q4.f, /9\.25%\/360/); assert.match(out.N5.f, /\)\+25000$/); assert.strictEqual(out.Q5.v, -7); assert.ok(!out.Q5.f);
});
t('posting dari hasil validasi: hanya nasabah PEI yang lolos (masuk iFast) dan punya CID', () => {
  const vres = C.runValidasi(ctx, C.newState());
  const { postings, skipped } = P.postingsFromResults(vres, ctx, C.lolosRp, C.lolosLr);
  assert.ok(postings.length > 0);
  for (const p of postings) {
    const d = vres.get(p.sid); assert.ok(d.is_pei);
    if (p.lr) { assert.ok(C.lolosLr(d)); assert.strictEqual(p.lr, d.lr_final); assert.ok(p.buys.length); }
    if (p.rp) { assert.ok(C.lolosRp(d)); assert.strictEqual(p.rp, d.rp_final); }
    assert.strictEqual(p.cid, ctx.sidToCid.get(p.sid));
  }
  assert.ok(Array.isArray(skipped));
  const m = empty(); P.applyDay(m, postings, { trxDate: D1, dueDate: D1 + 2, rate: 9.5 });
  const wb = roundTrip(P.buildWorkbook(m, XS, { asOf: D1, rate: 9.5, prices }));
  assert.strictEqual(wb.SheetNames.length, 2 + m.clients.size);
});
t('bukan file porto → ditolak dengan pesan jelas', () => {
  assert.throws(() => P.parsePorto({ SheetNames: ['Sheet1'], Sheets: { Sheet1: { '!ref': 'A1:A1', A1: { t: 's', v: 'x' } } } }), /Bukan file porto/);
});
console.log(`\n${n} tes porto lulus`);
