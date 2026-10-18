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

## Alur 4 langkah (berurutan, langkah berikutnya terkunci sampai langkah sebelumnya selesai)
1. **Upload** — 6 berkas wajib (bisa per folder) + berkas opsional (Participant Stock Position, Data Transaksi LR). Klik *Jalankan Proses*.
2. **RP** — tabel nasabah PEI dengan transaksi Sell. RP Final default = RP Max; ubah bila perlu. Unduh `RP_Final_YYYYMMDD.xlsx`.
3. **LR** — unggah file RP Final dari langkah 2. LR dihitung ulang dari RP Final itu (nasabah yang tidak ada di file dianggap RP Final = 0). LR Final bisa diedit atau di-*Auto-Adjust* bila Nilai Beli melebihi kapasitas. Klik *Final & Unduh* untuk membuat `Hasil_TRX_PEI.xlsx` (opsional: `LR_Final_YYYYMMDD.xlsx`).
4. **Validasi & Export** — unggah `Hasil_TRX_PEI.xlsx`, jalankan validasi MNC, Auto-Adjust bila perlu, lalu unduh file iFast (Repayment Proceed dan Loan Request).

Tab *Kenapa tidak muncul?* (penelusuran per nasabah) dan *Audit* bisa dibuka kapan saja setelah langkah 1.

Catatan LR-1: bila sheet LR pada file yang diunggah di langkah 4 memuat kolom Available/Volume dari MNC (≥14 kolom), itu yang dipakai. Jika tidak (mis. Hasil_TRX_PEI buatan sendiri), LR-1 memakai Available Qty dari Risk Parameter.

## Upload cepat per folder
Kumpulkan berkas harian dalam satu folder, lalu klik **Pilih Folder** (atau seret folder ke kotak). 6 berkas wajib dikenali dari **nama**; berkas lain di folder diabaikan:

| Berkas wajib | Nama harus memuat | Format |
|---|---|---|
| Closing Price | `closing price` | .xlsx |
| Risk Parameter | `RiskParameter` | .txt |
| Outstanding Position | diawali `OP_` | .txt |
| List of Invoice | `list invoice` / `netting` | .csv/.xls/.xlsx |
| Credit Limit | `CreditLimit` | .txt |
| SID Client | `SID Client` | .xlsx |

Jika ada lebih dari satu berkas yang cocok, dipakai yang terbaru (tanggal modifikasi). Berkas opsional (7–8) tetap diunggah manual. Pola nama ada di `SLOT_PATTERNS` pada `core.js`.

## Beda dengan versi login
- Tidak ada peran MNC read-only; semua pengguna bisa edit RP/LR Final dan Auto-Adjust.
- Hasil MNC (sheet LR) tidak lagi di langkah 1; diunggah di langkah 4 sebagai Hasil TRX.
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
