# SIKeu — halaman publik

Folder ini yang diunggah ke GitHub Pages. Isinya satu halaman pembungkus yang memuat aplikasi
Apps Script Anda, sehingga pengurus cukup membuka subdomain yayasan.

## Sekali setel

1. Buka `index.html`, cari baris:

   ```js
   var URL_APPSCRIPT = 'GANTI_DENGAN_URL_APPSCRIPT';
   ```

   Ganti dengan URL web app dari **Deploy → Manage deployments** di Apps Script
   (bentuknya `https://script.google.com/macros/s/AKfycb.../exec`).

2. Buka `CNAME`, ganti isinya dengan subdomain Anda, misalnya `keuangan.yimr.or.id`.
   Hapus berkas ini bila tidak memakai domain sendiri.

## Unggah ke GitHub

```bash
cd publik
git init
git add .
git commit -m "Halaman publik SIKeu"
git branch -M main
git remote add origin https://github.com/<akun>/sikeu-publik.git
git push -u origin main
```

Lalu di repositori: **Settings → Pages → Source: Deploy from a branch → main / (root) → Save**.
Tunggu 1–2 menit, alamatnya aktif di `https://<akun>.github.io/sikeu-publik/`.

## Menyambungkan subdomain

Di panel DNS domain Anda, tambahkan satu catatan:

| Tipe | Nama | Nilai |
|---|---|---|
| `CNAME` | `keuangan` | `<akun>.github.io` |

Kembali ke **Settings → Pages → Custom domain**, isi `keuangan.yimr.or.id`, **Save**, lalu centang
**Enforce HTTPS** setelah sertifikatnya terbit (5–30 menit).

## Catatan

- Pengurus tetap perlu masuk dengan akun Google yang diizinkan pada deployment Apps Script.
- Pada iPhone, bila halaman tampak kosong: *Settings → Safari → Prevent Cross-Site Tracking* dimatikan.
- Memperbarui aplikasi lewat **Manage deployments → Edit → New version** tidak mengubah URL,
  jadi folder ini tidak perlu disentuh lagi.
