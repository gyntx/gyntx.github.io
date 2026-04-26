# Rekap Bulanan Hotel

Web app untuk generate rekap bulanan hotel otomatis dari file laporan harian `.xlsx`.

## Fitur
- Upload file `.xlsx` laporan harian
- Generate 3 sheet rekap otomatis:
  - **MBR OCC RECAPITULATION** - Rekap okupansi harian
  - **REVENUE PER TIPE KAMAR** - Revenue per tipe kamar (Manggar, Queen, Mahligai)
  - **SEGMENTASI** - Market segmentasi per hari (QTY, Revenue, ARR)
- Semua proses di browser — data tidak dikirim ke server

## Cara Deploy ke GitHub Pages

1. Buat repository baru di GitHub (misal: `rekap-hotel`)
2. Upload semua file di folder ini ke repository
3. Buka **Settings** → **Pages**
4. Pilih **Source**: `Deploy from a branch`
5. Pilih branch `main`, folder `/ (root)`
6. Klik **Save**
7. Website aktif di `https://username.github.io/rekap-hotel`

## Struktur File
```
├── index.html
├── css/
│   └── style.css
├── js/
│   └── app.js
└── README.md
```
