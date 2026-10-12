# Dashboard OH · OM NRE (versi web)

Versi web dari dashboard Overhaul Control Center OM NRE, untuk dibuka siapa saja lewat GitHub Pages tanpa akun Claude.

## Isi folder

| File | Fungsi |
|---|---|
| `index.html` | Halaman dashboard (5 tab, filter, mode terang/gelap) |
| `dash.js` | Memuat `data.json` dan CSV laporan harian, lalu menjalankan dashboard |
| `config.js` | Pengaturan: alamat CSV Google Sheets dan interval refresh |
| `data.json` | Data dari 3 file Excel. **Dibuat oleh `update_data.py`, jangan diedit manual** |
| `update_data.py` | Membaca Excel → menulis `data.json` → commit & push ke GitHub |
| `update.bat` | Klik 2x untuk menjalankan `update_data.py` |
| `export.js` | Tombol **Unduh PDF** di setiap tab (A4 potret, tampilan terang) |

## Sumber data

| Tab | Sumber | Update |
|---|---|---|
| Jadwal OH NRE, Biaya, Jadwal Rapat | `Jadwal OH NRE*.xlsx` | Jalankan `update.bat` |
| Monitoring Progress Harian | Google Sheets, sheet "Laporan Harian" | Otomatis (dibaca langsung) |
| Monitoring OH Ditgas 2026 | `Renc&Real OH OM4*.xlsx`, `Jadwal OH OM 5*.xlsx`, `Jadwal OH NRE*.xlsx` | Jalankan `update.bat` |

File Excel dicari di folder induk folder ini (`claude database`). Kalau ada beberapa file yang cocok, yang paling baru diubah yang dipakai.

## Pemasangan pertama (sekali saja)

1. **Buat repo baru** di GitHub, misalnya `dashboard-oh-nre`, sebagai *Public*. Jangan centang "Add a README".
2. **Hubungkan folder ini ke repo.** Buka Command Prompt di folder ini, lalu jalankan:
   ```
   git init -b main
   git add .
   git commit -m "Dashboard OH awal"
   git remote add origin https://github.com/<akun-github>/dashboard-oh-nre.git
   git push -u origin main
   ```
   (Atau pakai GitHub Desktop: *Add existing repository* → pilih folder ini → *Publish*.)
3. **Aktifkan GitHub Pages:** di repo, buka *Settings → Pages → Build and deployment*, pilih *Deploy from a branch*, branch `main`, folder `/ (root)`, lalu *Save*. Setelah 1–2 menit dashboard bisa dibuka di `https://<akun-github>.github.io/dashboard-oh-nre/`.
4. **Sambungkan laporan harian:** di Google Sheets, buka *File → Bagikan → Publikasikan ke web*. Pilih sheet **Laporan Harian**, format **Nilai yang dipisahkan koma (.csv)**, klik *Publikasikan*, lalu salin link-nya ke `harianCsvUrl` di `config.js`. Commit dan push `config.js`.

## Memperbarui data Excel

- **Manual:** klik 2x `update.bat`. Script membaca ketiga Excel. Kalau ada perubahan, `data.json` ditimpa lalu otomatis di-commit dan di-push. Kalau tidak ada perubahan, tidak ada yang dikirim.
- **Otomatis (Task Scheduler Windows):**
  1. Buka *Task Scheduler → Create Basic Task*, beri nama "Update Dashboard OH".
  2. *Trigger:* Daily, lalu di *Properties → Triggers → Edit* centang *Repeat task every 30 minutes* selama 1 hari.
  3. *Action:* Start a program → `update.bat`, isi *Add arguments* dengan `auto` (supaya jendela tidak menunggu tombol), dan *Start in* dengan path folder ini.

## Unduh PDF

Setiap tab punya tombol **Unduh PDF** di kanan atas. PDF dibuat langsung di browser: ukuran A4 potret, selalu tampilan terang, berisi tab yang sedang dibuka sesuai filter yang dipilih (tahun, UP, divisi, status, pencarian). Tabel dicetak lengkap tanpa scroll, dan setiap halaman diberi nama tab, tanggal cetak, dan nomor halaman.

## Catatan

- `data.json` selalu ditimpa, jadi ukurannya tetap sekitar 300–400 KB. Riwayat versi lama tersimpan di Git sebagai selisih kecil, jadi tidak ada kuota yang habis.
- Repo dan Google Sheets yang di-*publish* bersifat publik: siapa pun yang tahu link-nya bisa melihat isi dashboard, termasuk angka biaya.
- Jika folder ini ada di dalam OneDrive dan muncul konflik sinkronisasi pada folder `.git`, pindahkan repo ke luar OneDrive dan isi `EXCEL_DIR` di `update_data.py` (atau variabel lingkungan `OH_EXCEL_DIR`) dengan path folder `claude database`.
- Butuh Python 3 dengan paket `openpyxl` (`pip install openpyxl`) dan Git yang sudah login ke akun GitHub.
