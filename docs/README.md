# PEI Generator (versi HTML, TANPA login) — MNC Sekuritas

Port dari `pages/Daily RP and LR Calculation.py` (Streamlit). Aplikasi statis murni: semua perhitungan jalan di browser, tidak ada server, tidak ada login, **berkas nasabah tidak pernah keluar dari komputer**.

## Struktur
```
public/    index.html, app.js, core.js, style.css, vendor/xlsx
tests/     tes pembanding vs Python asli, tes perilaku, tes UI
```

## Cara pakai
**Paling cepat (lokal):** klik dua kali `public/index.html`. Tidak perlu instal apa pun.

**Dihosting (Cloudflare Pages):** push ke branch baru (mis. `html-version`, bukan `main`), Create → Pages → Connect to Git, Framework *None*, Build command kosong, Output directory `public`. Tidak ada secret/variabel yang perlu diisi.

> Tanpa login, siapa pun yang tahu alamatnya bisa membuka aplikasinya. Data tetap aman karena tidak diunggah, tetapi untuk hosting internal sebaiknya batasi akses (mis. Cloudflare Access) atau pakai lokal saja.

## Beda dengan versi login
- Tidak ada peran MNC read-only; semua pengguna bisa edit RP/LR Final dan Auto-Adjust.
- Tab Audit hanya log sesi browser (hilang saat refresh); unduh CSV bila perlu disimpan.
- Folder `functions/`, `tools/hash-password.mjs`, halaman login, dan tes auth dihapus.

## Menjalankan tes
```
npm install
python3 tests/gen_data.py && python3 tests/ref_python.py   # perlu repo Streamlit di /home/claude/repo
npm test
```

## Catatan perilaku (penting)
- **Satu sumber kebenaran**: RP/LR Final yang diedit dipakai oleh Validasi, file iFast, dan Hasil_TRX_PEI. Rerun tidak mereset editan.
- RP yang dipotong terlalu kecil bisa membuat check **RP-3 (rasio < 65%)** gagal sehingga nasabah **keluar dari file iFast** (perilaku validasi asli; terlihat di tab Validasi dan Diagnostik).
- Dua mode hitung dipertahankan persis seperti Python: TRX PEI (`pei`) dan Validasi (`mnc`). Beda hanya untuk nasabah yang menjual tetapi Loan Existing = 0. Perlu keputusan apakah disatukan.
- Tanpa file Hasil MNC, check LR-1 memakai Available Qty dari Risk Parameter.
