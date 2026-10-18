(() => {
  'use strict';
  const C = window.PEICore, X = window.XLSX;
  // ── helper DOM (selalu textContent → aman dari XSS walau nama nasabah berisi HTML) ──
  const h = (tag, props, ...kids) => {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(props || {})) {
      if (k === 'class') el.className = v; else if (k.startsWith('on')) el.addEventListener(k.slice(2), v); else if (v === true) el.setAttribute(k, ''); else if (v !== false && v != null) el.setAttribute(k, v);
    }
    for (const kid of kids.flat()) if (kid != null && kid !== false) el.append(kid.nodeType ? kid : document.createTextNode(String(kid)));
    return el;
  };
  const $ = (id) => document.getElementById(id);
  const rp = (v) => (v == null ? '-' : C.fmtRp(v));
  const pct = (v) => (v == null ? '-' : C.fmtPct(v));
  const today = () => new Date().toISOString().slice(0, 10).replace(/-/g, '');

  const SLOTS = [
    { id: 'closing', label: '1. Closing Price', acc: '.xlsx', req: 1, kind: 'rows' },
    { id: 'risk', label: '2. Risk Parameter', acc: '.txt', req: 1, kind: 'text' },
    { id: 'op', label: '3. Outstanding Position', acc: '.txt', req: 1, kind: 'text' },
    { id: 'netting', label: '4. List of Invoice', acc: '.csv,.xls,.xlsx', req: 1, kind: 'rows' },
    { id: 'cl', label: '5. Credit Limit', acc: '.txt', req: 1, kind: 'text' },
    { id: 'sidclient', label: '6. SID Client (daftar nasabah PEI)', acc: '.xlsx', req: 1, kind: 'rows' },
    { id: 'pos', label: '7. Participant Stock Position', acc: '.txt', kind: 'text', hint: 'opsional — koreksi OP T+2 belum settle' },
    { id: 'lrsettle', label: '8. Data Transaksi LR', acc: '.txt', kind: 'text', hint: 'opsional — koreksi Loan Existing belum settle' },
  ];
  const S = { user: 'Lokal', role: 'PEI', files: {}, ctx: null, st: C.newState(), pei: null, mnc: null, tab: 'upload', msg: null, sent: 0, rpSaved: false, rpUnlocked: false, rpFile: null, hasilSaved: false, hasil: null, valRun: false };
  const canEdit = () => S.role === 'PEI';

  const readFile = (f, kind) => new Promise((res, rej) => {
    const r = new FileReader(); r.onerror = () => rej(new Error('Gagal membaca ' + f.name));
    if (kind === 'text') { r.onload = () => res(String(r.result)); r.readAsText(f); return; }
    r.onload = () => {
      try {
        const wb = X.read(r.result, { type: 'array', raw: /\.csv$/i.test(f.name) });
        if (kind === 'hasil') {
          const name = wb.SheetNames.find((n) => n === 'Loan Request (LR)') || wb.SheetNames.find((n) => n === 'Buy (Loan)');
          if (!name) throw new Error('Sheet "Loan Request (LR)" tidak ada di Hasil MNC. Sheet tersedia: ' + wb.SheetNames.join(', '));
          const m = new Map(); X.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' }).slice(1).forEach((row) => { const k = String(row[0]); if (!m.has(k)) m.set(k, []); m.get(k).push(row); });
          return res(m);
        }
        res(X.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: undefined }));
      } catch (e) { rej(e); }
    };
    r.readAsArrayBuffer(f);
  });

  async function setFile(slot, f) {
    try { S.files[slot.id] = { name: f.name, data: await readFile(f, slot.kind) }; return null; }
    catch (er) { delete S.files[slot.id]; return `${slot.label.replace(/^\d+\.\s*/, '')}: ${er.message}`; }
  }
  // Terima kumpulan berkas (dari pilih folder / pilih banyak berkas / drag-drop), cocokkan dengan 6 slot wajib dari nama
  async function loadMany(list) {
    const files = [...list].filter((f) => f && f.name);
    const { picked, dupes, ignored } = C.matchFolderFiles(files);
    const errs = [];
    for (const slot of SLOTS.filter((x) => x.req)) if (picked[slot.id]) { const e = await setFile(slot, picked[slot.id]); if (e) errs.push(e); }
    const miss = SLOTS.filter((x) => x.req && !S.files[x.id]).map((x) => x.label.replace(/^\d+\.\s*/, '').replace(/ \(.*\)$/, ''));
    const dup = Object.keys(dupes).map((id) => SLOTS.find((x) => x.id === id).label.replace(/^\d+\.\s*/, '').replace(/ \(.*\)$/, '') + ` (${dupes[id]} berkas, dipakai yang terbaru: ${picked[id].name})`);
    const found = Object.keys(picked).length;
    S.msg = errs.length ? { er: 1, t: 'Ada berkas yang gagal dibaca. ' + errs.join(' | ') }
      : !found ? { er: 1, t: `Tidak ada berkas wajib yang dikenali dari ${files.length} berkas. Pastikan nama berkas memuat: Closing Price, RiskParameter, OP_, List Invoice, CreditLimit, SID Client.` }
      : { t: `${found} dari 6 berkas wajib terdeteksi otomatis.` + (miss.length ? ` Belum ada: ${miss.join(', ')}.` : '') + (dup.length ? ` Catatan duplikat: ${dup.join('; ')}.` : '') };
    S.ignored = ignored; render();
  }
  // drag-drop: dukung folder (baca isi satu level) maupun berkas
  async function filesFromDrop(dt) {
    const items = [...(dt.items || [])].map((i) => (i.webkitGetAsEntry ? i.webkitGetAsEntry() : null));
    if (!items.length || items.some((e) => !e)) return [...dt.files];
    const out = [];
    const fileOf = (en) => new Promise((res) => en.file(res, () => res(null)));
    for (const en of items) {
      if (en.isFile) out.push(await fileOf(en));
      else if (en.isDirectory) {
        const rd = en.createReader(); let batch;
        do { batch = await new Promise((res) => rd.readEntries(res, () => res([]))); for (const c of batch) if (c.isFile) out.push(await fileOf(c)); } while (batch.length);
      }
    }
    return out.filter(Boolean);
  }

  // ── audit: dicatat lokal di sesi browser (dapat diunduh dari tab Audit) ──
  function audit() {}
  const recompute = () => { S.pei = C.runPei(S.ctx, S.st); S.mnc = C.runMnc(S.ctx, S.st); };
  function edit(kind, sid, raw, d) {
    if (!canEdit()) return;
    const digits = String(raw).replace(/[^\d]/g, '');
    const val = digits === '' ? null : Number(digits);
    C.setOverride(S.st, kind, sid, val, S.user, val == null ? 'kembali ke rekomendasi' : 'edit manual'); audit(); recompute(); render();
  }
  const download = (name, blob) => { const a = h('a', { href: URL.createObjectURL(blob), download: name }); document.body.append(a); a.click(); a.remove(); };
  const xlsxOut = (name, sheets) => { const wb = X.utils.book_new(); for (const [n, rows] of Object.entries(sheets)) X.utils.book_append_sheet(wb, X.utils.json_to_sheet(rows), n); X.writeFile(wb, name); };
  const csvOut = (name, rows) => download(name, new Blob(['\ufeff' + C.toCsv(rows)], { type: 'text/csv;charset=utf-8' }));

  // ── tabel generik ──
  function table(cols, rows) {
    return h('div', { class: 'tw' }, h('table', {}, h('thead', {}, h('tr', {}, cols.map((c) => h('th', {}, c.t)))),
      h('tbody', {}, rows.map((r) => h('tr', {}, cols.map((c) => { const v = c.f(r); return h('td', { class: (c.n ? 'n ' : '') + (c.cls ? c.cls(r) : '') }, v); }))))));
  }
  const cell = (d, kind, field, label) => canEdit()
    ? h('span', {}, h('input', { class: 'cell' + (d[kind + '_edited'] ? ' ed' : ''), type: 'text', inputmode: 'numeric', 'aria-label': `${label} ${d.sid}`, value: Math.round(d[field]).toString(), onchange: (e) => edit(kind, d.sid, e.target.value, d) }),
      d[kind + '_edited'] ? h('button', { class: 'ghost', title: 'Kembalikan ke rekomendasi', onclick: () => edit(kind, d.sid, '', d) }, '↺') : null)
    : rp(d[field]);

  // ── tab Upload ──
  function viewUpload() {
    const slot = (s) => h('div', { class: 'up' + (S.files[s.id] ? ' done' : '') },
      h('b', {}, s.label.replace(/ \(opsional\)$/, '')), s.hint ? h('small', {}, s.hint.replace(/^opsional — /, '')) : null,
      h('input', { type: 'file', accept: s.acc, 'aria-label': s.label, onchange: async (e) => {
        const f = e.target.files[0]; if (!f) { delete S.files[s.id]; return render(); }
        const er = await setFile(s, f); S.msg = er ? { er: 1, t: er } : null; render(); } }),
      S.files[s.id] ? h('div', { class: 'ok' }, '✓ ' + S.files[s.id].name) : null);
    const req = SLOTS.filter((s) => s.req), opt = SLOTS.filter((s) => !s.req);
    const have = req.filter((s) => S.files[s.id]).length, ready = have === req.length;
    const bar = h('i'); bar.style.width = Math.round(have / req.length * 100) + '%'; // via CSSOM (CSP melarang atribut style)
    const dir = h('input', { type: 'file', id: 'pickDir', 'aria-label': 'Pilih folder', webkitdirectory: true, multiple: true, hidden: true,
      onchange: (e) => { const l = [...e.target.files].filter((f) => (f.webkitRelativePath || '').split('/').length <= 2); e.target.value = ''; loadMany(l); } });
    const multi = h('input', { type: 'file', id: 'pickMulti', 'aria-label': 'Pilih beberapa berkas', multiple: true, hidden: true,
      onchange: (e) => { const l = [...e.target.files]; e.target.value = ''; loadMany(l); } });
    const zone = h('div', { class: 'drop', id: 'dropzone' },
      h('div', { class: 'big' }, '📁'), h('b', {}, 'Cara cepat: pilih satu folder berisi semua berkas harian'),
      h('p', { class: 'mut' }, 'Aplikasi mengenali 6 berkas wajib dari namanya (Closing Price, RiskParameter, OP_, List Invoice, CreditLimit, SID Client). Berkas lain di folder itu diabaikan.'),
      h('div', { class: 'row c' }, h('button', { onclick: () => dir.click() }, '📁 Pilih Folder'), h('button', { class: 'ghost', onclick: () => multi.click() }, 'Pilih beberapa berkas'), h('span', { class: 'mut' }, 'atau seret folder/berkas ke sini')),
      dir, multi);
    zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('over'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('over'));
    zone.addEventListener('drop', async (e) => { e.preventDefault(); zone.classList.remove('over'); loadMany(await filesFromDrop(e.dataTransfer)); });
    const checklist = h('div', { class: 'chk' }, req.map((s) => h('span', { class: 'ci ' + (S.files[s.id] ? 'ok' : 'mut') }, (S.files[s.id] ? '✓ ' : '○ ') + s.label.replace(/^\d+\.\s*/, '').replace(/ \(.*\)$/, ''), S.files[s.id] ? h('small', {}, ' — ' + S.files[s.id].name) : null)));
    const manual = h('details', { class: 'card man' }, h('summary', {}, 'Unggah satu per satu / koreksi manual (6 berkas wajib)'), h('div', { class: 'grid' }, req.map(slot)));
    if (!ready) manual.open = true;
    return h('div', {},
      h('div', { class: 'card' }, h('h2', {}, 'File input wajib'), zone, checklist),
      manual,
      h('div', { class: 'card' }, h('h2', {}, 'File opsional'), h('div', { class: 'grid' }, opt.map(slot))),
      h('div', { class: 'card run' }, h('button', { disabled: !ready, onclick: run }, '▶ Jalankan Proses'),
        h('div', { class: 'prog' }, bar),
        h('span', { class: 'mut' }, ready ? 'Semua berkas wajib sudah terisi. Siap diproses.' : `${have} dari ${req.length} berkas wajib terisi.`)),
      S.msg ? h('div', { class: 'msg ' + (S.msg.er ? 'er' : '') }, S.msg.t) : null);
  }
  function run() {
    try {
      const f = (id) => S.files[id] && S.files[id].data;
      S.ctx = C.buildContext({ closingRows: f('closing'), riskText: f('risk'), opText: f('op'), nettingRows: f('netting'), clText: f('cl'), sidClientRows: f('sidclient'), posText: f('pos'), lrSettleText: f('lrsettle'), hasilBuyRows: f('hasil') });
      S.st = C.newState(); S.sent = 0; S.rpSaved = S.rpUnlocked = S.hasilSaved = S.valRun = false; S.rpFile = null; S.hasil = null; recompute(); S.tab = 'rp';
      const vd = S.ctx.clValueDate, td = new Date().toISOString().slice(0, 10).replace(/-/g, '/');
      S.msg = { t: `✅ File bersama dimuat — ${S.ctx.op.size} SID di OP, ${S.ctx.cp.size} harga saham, ${S.ctx.netting.size} SID di Netting, ${S.ctx.cl.size} di Credit Limit, ${S.ctx.peiSids.size} nasabah PEI.` + (vd && vd !== td ? ` ⚠️ Value Date Credit Limit ${vd} ≠ hari ini ${td}.` : '') };
    } catch (e) { S.msg = { er: 1, t: 'Gagal memproses: ' + e.message }; S.tab = 'upload'; }
    render();
  }

  // ── ringkasan angka (seperti metrik di Streamlit) ──
  const stats = () => h('div', { class: 'stats' },
    [['Total Nasabah PEI Terdaftar', S.pei.size], ['Nasabah Transaksi Sell', [...S.pei.values()].filter((d) => d.total_rp_maks > 0).length], ['Nasabah Transaksi Buy', [...S.pei.values()].filter((d) => d.total_buy_val > 0).length]]
      .map(([t, v]) => h('div', { class: 'stat' }, h('small', {}, t), h('b', {}, String(v)))));
  const go = (id) => { S.tab = id; S.msg = null; render(); };
  const msgBox = () => (S.msg ? h('div', { class: 'msg ' + (S.msg.er ? 'er' : S.msg.wr ? 'wr' : '') }, S.msg.t) : null);
  const nextBar = (label, onclick, enabled, hint) => h('div', { class: 'card run' }, h('button', { disabled: !enabled, onclick }, label), h('span', { class: 'mut' }, hint));

  // ── LANGKAH 2: RP ──
  const rpRowsOf = () => [...S.pei.values()].filter((d) => d.total_rp_maks > 0);
  const lrRowsOf = () => [...S.pei.values()].filter((d) => d.total_buy_val > 0);
  const stockStr = (d) => d.rp_detail.filter((x) => x.ada_di_op).map((x) => `${x.stock} (${Math.trunc(x.lot_keluar)})`).join(', ');
  const rpFileRows = () => rpRowsOf().map((d) => ({ 'SID': d.sid, 'Nama': d.name, 'Saham (Lot Keluar)': stockStr(d), 'Loan Existing': d.loan_existing,
    'RP Min (ref, target 65%)': d.total_rp_min_ratio, 'RP Max (rekomendasi)': d.total_rp_maks, 'RP Final': d.rp_final }));
  function viewRp() {
    const rows = rpRowsOf();
    const cols = [{ t: 'SID', f: (d) => d.sid }, { t: 'Nama', f: (d) => d.name }, { t: 'Saham (Lot Keluar)', f: stockStr },
      { t: 'Loan Existing', n: 1, f: (d) => rp(d.loan_existing) }, { t: 'RP Min (ref, target 65%)', n: 1, f: (d) => rp(d.total_rp_min_ratio) },
      { t: 'RP Max (rekomendasi)', n: 1, f: (d) => rp(d.total_rp_maks) }, { t: 'RP Final', f: (d) => cell(d, 'rp', 'rp_final', 'RP Final') }, { t: 'Rasio after RP', n: 1, f: (d) => pct(d.rasio_rp) }];
    const save = () => { xlsxOut(`RP_Final_${today()}.xlsx`, { 'RP Final': rpFileRows() }); S.rpSaved = true; render(); };
    return h('div', {}, stats(), msgBox(),
      h('div', { class: 'msg' }, '💡 Menampilkan hanya nasabah PEI yang benar-benar memiliki transaksi Sell (saham yang match dengan Outstanding Position). RP Final default terisi RP Max (rekomendasi). Ubah manual di kolom RP Final bila perlu, lalu simpan ke Excel.'),
      rows.length ? h('div', {}, h('p', { class: 'mut' }, `📌 ${rows.length} nasabah PEI dengan transaksi Sell diproses (dari total ${S.pei.size} nasabah PEI ter-record).`), table(cols, rows),
        h('div', { class: 'row' }, h('button', { onclick: save }, '💾 Simpan RP Final (.xlsx)'), S.rpSaved ? h('span', { class: 'ok' }, '✓ RP Final sudah diunduh') : null))
        : h('p', { class: 'mut' }, 'Tidak ada nasabah PEI dengan transaksi Sell hari ini. Langkah RP dapat dilewati.'),
      nextBar('Lanjut ke Langkah 3: LR →', () => { S.rpUnlocked = true; go('lr'); }, S.rpSaved || !rows.length,
        S.rpSaved || !rows.length ? 'Siap. Di langkah 3 unggah file RP Final yang baru diunduh.' : 'Unduh RP Final dulu, lalu lanjut.'));
  }

  // ── LANGKAH 3: LR ──
  async function onRpFinalUpload(e) {
    const f = e.target.files[0]; if (!f) return;
    try {
      const rows = await readFile(f, 'rows');
      if (!rows.length || !('SID' in rows[0]) || !('RP Final' in rows[0])) throw new Error('Kolom "SID" dan "RP Final" tidak ditemukan. Pakai file RP_Final hasil unduhan dari Langkah 2.');
      const map = new Map(); let bad = 0;
      for (const r of rows) { const sid = String(r.SID == null ? '' : r.SID).trim(); if (!sid) continue; const v = C.pyFloat(r['RP Final']); if (Number.isNaN(v)) { bad++; map.set(sid, 0); } else map.set(sid, v); }
      let changed = 0;
      for (const [sid, d] of S.pei) { const nv = map.has(sid) ? map.get(sid) : 0; if (Math.round(d.rp_final) !== Math.round(nv)) { C.setOverride(S.st, 'rp', sid, nv, S.user, 'dari file RP Final'); changed++; } }
      audit(); recompute(); S.rpFile = { name: f.name, n: map.size, notPei: [...map.keys()].filter((k) => !S.pei.has(k)).length };
      S.msg = { t: `✅ RP Final dimuat — ${map.size} SID. Nasabah yang tidak ada di file dianggap RP Final = 0.` + (changed ? ` ${changed} nilai RP berbeda dari Langkah 2 dan memakai isi file.` : '') + (bad ? ` ⚠️ ${bad} baris RP Final tidak terbaca (dianggap 0).` : '') + (S.rpFile.notPei ? ` ${S.rpFile.notPei} SID di file bukan nasabah PEI (diabaikan).` : ''), wr: !!bad };
    } catch (er) { S.rpFile = null; S.msg = { er: 1, t: '❌ Gagal membaca file RP Final: ' + er.message }; }
    render();
  }
  const lrFileRows = () => lrRowsOf().map((d) => ({ 'SID': d.sid, 'Nama': d.name, 'Saham (Lot Beli)': Object.entries(d.buy_stocks).map(([s, b]) => `${s} (${Math.trunc(b.lot)})`).join(', '), 'Nilai Beli': d.total_buy_val,
    'RP Final (dari upload)': d.rp_final, 'Loan After RP': d.loan_after_rp, 'Avail Limit (Awal)': d.avail_limit, 'Δ Avail Limit (RP)': d.avail_efektif - d.avail_limit, 'Avail Efektif': d.avail_efektif,
    'LR @63%': d.max_lr_63, 'Max LR Final (rekomendasi)': d.max_lr_final, 'Binding': d.lr_binding, 'Rasio LR (jika full)': pct(d.rasio_lr), 'LR Final': d.lr_final }));
  const lr3Fail = (d) => d.has_lr && d.checks.some((c) => c.label === C.LR3 && !c.passed);
  function viewLr() {
    const up = h('div', { class: 'card' }, h('h2', {}, 'Upload RP Final'),
      h('div', { class: 'up' + (S.rpFile ? ' done' : '') }, h('b', {}, 'RP Final (.xlsx hasil unduhan Langkah 2)'),
        h('input', { type: 'file', accept: '.xlsx,.xls', 'aria-label': 'Upload RP Final', onchange: onRpFinalUpload }), S.rpFile ? h('div', { class: 'ok' }, '✓ ' + S.rpFile.name) : null),
      h('p', { class: 'mut' }, 'RP Final memengaruhi Loan After RP dan Available Limit efektif, sehingga LR dihitung dari nilai di file ini, bukan RP Max.'));
    if (!S.rpFile) return h('div', {}, stats(), msgBox(), up, h('div', { class: 'msg wr' }, '⬆️ Upload file RP Final untuk menghitung LR.'));
    const rows = lrRowsOf(); const fail = rows.filter(lr3Fail).length;
    const cols = [{ t: 'SID', f: (d) => d.sid }, { t: 'Nama', f: (d) => d.name }, { t: 'Saham (Lot Beli)', f: (d) => Object.entries(d.buy_stocks).map(([s, b]) => `${s} (${Math.trunc(b.lot)})`).join(', ') },
      { t: 'Nilai Beli', n: 1, f: (d) => rp(d.total_buy_val) }, { t: 'RP Final (dari upload)', n: 1, f: (d) => rp(d.rp_final) }, { t: 'Loan After RP', n: 1, f: (d) => rp(d.loan_after_rp) },
      { t: 'Avail Limit (Awal)', n: 1, f: (d) => rp(d.avail_limit) }, { t: 'Δ Avail Limit (RP)', n: 1, f: (d) => rp(d.avail_efektif - d.avail_limit) }, { t: 'Avail Efektif', n: 1, f: (d) => rp(d.avail_efektif) },
      { t: 'LR @63%', n: 1, f: (d) => rp(d.max_lr_63) }, { t: 'Max LR Final (rekomendasi)', n: 1, f: (d) => rp(d.max_lr_final) }, { t: 'Binding', f: (d) => d.lr_binding },
      { t: 'Rasio LR (jika full)', n: 1, f: (d) => pct(d.rasio_lr) }, { t: 'LR Final', f: (d) => cell(d, 'lr', 'lr_final', 'LR Final') },
      { t: 'Cek LR-3', f: (d) => (lr3Fail(d) ? '⚠️ Nilai Beli > LR Final' : d.adjusted ? '✂️ Di-adjust' : '✅ Muat'), cls: (d) => (lr3Fail(d) ? 'wr' : 'ok') }];
    const finish = () => { xlsxOut('Hasil_TRX_PEI.xlsx', C.hasilTrxPei(S.pei)); S.hasilSaved = true; render(); };
    return h('div', {}, stats(), msgBox(), up,
      h('div', { class: 'msg' }, '💡 Menampilkan hanya nasabah PEI dengan transaksi Buy margin. LR Final default = Max LR Final (rekomendasi). Ubah manual atau gunakan Auto-Adjust bila Nilai Beli melebihi kapasitas, lalu finalkan.'),
      rows.length ? h('div', {}, h('p', { class: 'mut' }, `📌 ${rows.length} nasabah PEI dengan transaksi Buy diproses (dari total ${S.pei.size} nasabah PEI ter-record).`), table(cols, rows),
        h('div', { class: 'row' }, h('button', { disabled: !fail || !canEdit(), onclick: () => { C.applyAutoAdjust(S.st, new Map([...S.pei].map(([k, v]) => [k, Object.assign({}, v, { is_pei: true })])), S.user); audit(); recompute(); render(); } }, `⚡ Auto-Adjust LR (${fail} nasabah melebihi kapasitas)`),
          S.st.adjusted.size ? h('button', { class: 'ghost', onclick: () => { S.st.adjusted.clear(); recompute(); render(); } }, 'Batalkan Auto-Adjust') : null,
          h('button', { class: 'ghost', onclick: () => xlsxOut(`LR_Final_${today()}.xlsx`, { 'LR Final': lrFileRows() }) }, '💾 Simpan LR Final (.xlsx)')))
        : h('p', { class: 'mut' }, 'Tidak ada nasabah PEI dengan transaksi Buy hari ini.'),
      h('div', { class: 'card run' }, h('button', { onclick: finish }, '✔ Final & Unduh Hasil_TRX_PEI.xlsx'),
        S.hasilSaved ? h('span', { class: 'ok' }, '✓ Hasil_TRX_PEI sudah dibuat') : h('span', { class: 'mut' }, 'Mengunci RP Final & LR Final saat ini ke file Hasil TRX.'),
        h('button', { disabled: !S.hasilSaved, onclick: () => go('val') }, 'Lanjut ke Langkah 4: Validasi →')));
  }

  // ── LANGKAH 4: Validasi + Export iFast ──
  async function onHasilUpload(e) {
    const f = e.target.files[0]; S.valRun = false;
    if (!f) { S.hasil = null; return render(); }
    try { S.hasil = { name: f.name, data: await readFile(f, 'hasil') }; S.msg = null; }
    catch (er) { S.hasil = null; S.msg = { er: 1, t: '❌ Gagal membaca Hasil TRX: ' + er.message }; }
    render();
  }
  function runVal() {
    // sheet LR dari MNC punya ≥14 kolom (saham=1, avail=4, volume=13). File Hasil_TRX_PEI buatan sendiri tidak punya → LR-1 pakai Available Qty Risk Parameter.
    let wide = false; for (const rows of S.hasil.data.values()) for (const r of rows) if (r.length >= 14) wide = true;
    S.ctx.hasilBuyRows = wide ? S.hasil.data : null; recompute(); S.valRun = true;
    S.msg = { t: wide ? '✅ Validasi selesai. LR-1 memakai Available/Volume dari sheet LR di file yang diunggah.' : '✅ Validasi selesai. File belum memuat kolom Available/Volume dari MNC, jadi LR-1 memakai Available Qty dari Risk Parameter.' }; render();
  }
  const mark = (ok, skipped) => (skipped ? '⏭' : ok ? '✅' : '❌');
  function valResults() {
    const g = C.globalResult(S.mnc); const rows = [...S.mnc.values()].filter((d) => d.is_pei && (d.has_rp || d.has_lr));
    const failing = (d) => d.checks.filter((c) => !c.passed).map((c) => c.label.split('.')[0] + ': ' + c.detail).join(' | ');
    const needAdj = rows.filter((d) => d.has_lr && d.checks.some((c) => c.label === C.LR3 && !c.passed)).length;
    const R = C.ifastRepayment(S.mnc), L = C.ifastLoan(S.mnc);
    const dropped = [...S.mnc.values()].filter((d) => d.is_pei && ((d.has_lr && !C.lolosLr(d) && d.lr_final > 0) || (d.has_rp && !d.rp_skipped && !C.lolosRp(d)))).length;
    const btn = (label, fn, n) => h('button', { disabled: !n, onclick: fn }, label);
    const block = (title, data, main, file) => h('div', { class: 'up' }, h('b', {}, `${title} — ${data[main].length} nasabah`),
      h('div', { class: 'row' }, btn('⬇️ Excel (.xlsx)', () => xlsxOut(`${file} ${today()}.xlsx`, data), data[main].length),
        btn('CSV (sheet utama)', () => csvOut(`${file} ${today()}.csv`, data[main]), data[main].length), btn('CSV (Stock)', () => csvOut(`${file} Stock ${today()}.csv`, data.Stock), data.Stock.length)));
    return h('div', {},
      h('div', { class: 'card' }, h('h2', {}, 'Hasil validasi'), h('div', { class: 'msg ' + (g.passed ? 'okb' : 'er') }, (g.passed ? '✅ Limit Participant LOLOS — ' : '❌ Limit Participant GAGAL — ') + g.detail),
        h('div', { class: 'row' }, h('button', { disabled: !needAdj || !canEdit(), onclick: () => { C.applyAutoAdjust(S.st, S.mnc, S.user); audit(); recompute(); render(); } }, `⚡ Auto-Adjust LR (${needAdj} nasabah gagal LR-3)`),
          h('span', { class: 'mut' }, 'Nasabah yang gagal check tidak masuk file iFast sampai lolos.')),
        table([{ t: 'SID', f: (d) => d.sid }, { t: 'Nama', f: (d) => d.name },
          { t: 'RP Final', n: 1, f: (d) => (d.has_rp && !d.rp_skipped ? rp(d.rp_final) : '-') }, { t: 'RP', f: (d) => (d.has_rp ? mark(C.lolosRp(d), d.rp_skipped) : '-') },
          { t: 'LR Final', n: 1, f: (d) => (d.has_lr ? rp(d.lr_final) : '-') }, { t: 'LR', f: (d) => (d.has_lr ? mark(C.lolosLr(d)) : '-') },
          { t: 'Check yang gagal', f: (d) => failing(d), cls: () => 'er' }], rows)),
      h('div', { class: 'card' }, h('h2', {}, 'Export ke sistem (iFast)'),
        dropped ? h('div', { class: 'msg wr' }, `⚠️ ${dropped} nasabah punya transaksi tetapi TIDAK masuk file iFast karena gagal validasi. Lihat tabel di atas atau tab Diagnostik.`) : null,
        h('div', { class: 'grid' }, block('Repayment iFast', R, 'Repayment Direct', 'Repayment Proceed'), block('Loan Request iFast', L, 'Loan', 'Loan Request')),
        h('p', { class: 'mut' }, 'Semua file memakai nilai RP Final / LR Final terkini.')));
  }
  function viewVal() {
    const up = h('div', { class: 'card' }, h('h2', {}, 'Upload Hasil TRX'),
      h('div', { class: 'up' + (S.hasil ? ' done' : '') }, h('b', {}, 'Hasil_TRX_PEI (.xlsx) dari Langkah 3'),
        h('small', {}, 'Harus memuat sheet "Loan Request (LR)" dan sheet Repayment.'),
        h('input', { type: 'file', accept: '.xlsx', 'aria-label': 'Upload Hasil TRX PEI', onchange: onHasilUpload }), S.hasil ? h('div', { class: 'ok' }, '✓ ' + S.hasil.name) : null),
      h('div', { class: 'row' }, h('button', { disabled: !S.hasil, onclick: runVal }, '▶ Jalankan Validasi MNC'), S.hasil ? null : h('span', { class: 'mut' }, 'Upload file Hasil TRX dulu.')));
    return h('div', {}, msgBox(), up, S.valRun ? valResults() : null);
  }

  // ── tab Diagnostik ──
  let diagOut = null;
  function viewDiag() {
    const inp = h('input', { type: 'text', placeholder: 'SID, CID, atau nama nasabah (mis. Firbert)', 'aria-label': 'Cari nasabah' });
    const go = () => { diagOut = C.diagnoseSid(inp.value, S.ctx, S.mnc); render(); };
    inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') go(); });
    const ico = { ok: '✅', warn: '⚠️', drop: '⛔' };
    return h('div', {}, h('p', { class: 'mut' }, 'Telusuri di tahap mana seorang nasabah gugur: SID Client → Netting → Closing Price → OP → Credit Limit → RP/LR → validasi → iFast.'),
      h('div', { class: 'row' }, inp, h('button', { onclick: go }, 'Telusuri')),
      diagOut ? h('div', {}, h('h3', {}, diagOut.sid ? `${diagOut.sid} — ${diagOut.verdict}` : diagOut.verdict),
        table([{ t: '', f: (s) => ico[s.status] }, { t: 'Tahap', f: (s) => s.stage }, { t: 'Keterangan', f: (s) => s.msg, cls: (s) => (s.status === 'drop' ? 'er' : s.status === 'warn' ? 'wr' : '') }], diagOut.stages)) : null);
  }

  // ── tab Audit ──
  function viewAudit() {
    const rows = S.st.audit.map((a) => ({ Waktu: a.ts, User: a.user, Jenis: a.kind, SID: a.sid, Dari: a.from, Ke: a.to, Catatan: a.note }));
    return h('div', {}, h('div', { class: 'row' }, h('button', { disabled: !rows.length, onclick: () => csvOut(`Audit PEI ${today()}.csv`, rows) }, '⬇️ Unduh log (CSV)'),
      h('span', { class: 'mut' }, 'Log sesi ini (hilang saat halaman ditutup atau di-refresh). Unduh CSV bila perlu disimpan.')),
      rows.length ? table(Object.keys(rows[0]).map((k) => ({ t: k, f: (r) => (r[k] == null ? '-' : r[k]) })), rows) : h('p', { class: 'mut' }, 'Belum ada perubahan.'));
  }

  const STEPS = [['upload', '1. Upload', viewUpload, () => true, () => !!S.ctx], ['rp', '2. RP', viewRp, () => !!S.ctx, () => S.rpSaved],
    ['lr', '3. LR', viewLr, () => !!S.ctx && S.rpUnlocked, () => S.hasilSaved], ['val', '4. Validasi & Export', viewVal, () => !!S.ctx && S.hasilSaved, () => S.valRun],
    ['diag', 'Kenapa tidak muncul?', viewDiag, () => !!S.ctx, null], ['audit', 'Audit', viewAudit, () => !!S.ctx, null]];
  function render() {
    const keep = [...document.querySelectorAll('.tw')].map((e) => e.scrollTop);
    $('tabs').replaceChildren(...STEPS.map(([id, label, , open, done]) => h('button', { class: (S.tab === id ? 'on ' : '') + (done && done() ? 'dn' : ''), disabled: !open(), onclick: () => go(id) }, (done && done() ? '✓ ' : '') + label)));
    const t = STEPS.find((x) => x[0] === S.tab); const v = t[2](); $('main').replaceChildren(S.tab === 'upload' ? v : h('div', { class: 'view' }, v));
    document.querySelectorAll('.tw').forEach((e, i) => { if (keep[i]) e.scrollTop = keep[i]; });
  }
  $('who').textContent = '👤 Mode lokal (tanpa login)';
  try { $('today').textContent = '📅 ' + new Date().toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }); } catch (e) {}
  render();
})();
