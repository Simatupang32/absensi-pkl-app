# Absensi PKL

Aplikasi absensi siswa PKL: login username/password, selfie kamera, validasi radius hotel,
jam masuk/pulang (telat), sakit/izin dengan upload surat, departemen, dan rekap Excel/PDF.

## Isi repo

| File / folder | Fungsi |
|---|---|
| `apps-script/Code.gs`, `apps-script/Index.html` | Kode utama, ditempel di Google Apps Script (Google Sheet → Extensions → Apps Script) |
| `index.html` | Aplikasi HP (PWA) yang membuka Web App Apps Script, dengan logo sendiri |
| `manifest.webmanifest`, `sw.js` | Agar bisa dipasang di layar utama HP |
| `icons/` | Logo dalam berbagai ukuran (Android, iPhone, favicon) |

## Pemasangan

1. **Apps Script** – salin isi `apps-script/Code.gs` dan `apps-script/Index.html` ke editor Apps Script,
   lalu **Deploy → Manage deployments → ✏️ → New version → Deploy**. Salin *Web app URL* (berakhiran `/exec`).
2. **URL aplikasi** – buka `index.html`, ganti nilai `APP_URL` dengan URL tersebut, lalu commit.
3. **GitHub Pages** – Settings → Pages → *Source: Deploy from a branch* → Branch `main`, folder `/ (root)` → Save.
   Tunggu ±1 menit, aplikasi tersedia di `https://simatupang32.github.io/Absensi-PKL/`.

## Pasang di HP

- **Android (Chrome):** buka link di atas → tekan **Pasang** (atau menu ⋮ → *Install app / Tambahkan ke layar utama*).
- **iPhone (Safari):** buka link → tombol **Bagikan** → **Tambah ke Layar Utama**.

Ikon OJT akan muncul di layar utama dan aplikasi terbuka layar penuh tanpa bilah alamat.
Jika kamera/lokasi tidak berjalan di dalam aplikasi, tombol **Buka Aplikasi** membuka Web App langsung di browser.
