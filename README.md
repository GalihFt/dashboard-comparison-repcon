# Pre vs Post Survey Monitor

Streamlit app untuk memantau perbandingan PRE SURVEY dan POST SURVEY.

## Struktur

- `app.py` - aplikasi Streamlit.
- `data/summary.parquet` - data summary siap pakai.
- `data/detail.parquet` - data detail siap pakai.
- `requirements.txt` - dependency Streamlit Cloud.

App langsung membaca file parquet di folder `data/`, jadi tidak perlu memproses Excel besar saat halaman dibuka.

Pencarian detail menerima satu atau beberapa `NO_EOR`. Pisahkan beberapa nilai dengan
baris baru, koma, atau titik koma.

## Jalankan Lokal

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy Streamlit Cloud

1. Upload/push isi folder ini ke GitHub.
2. Buat app baru di Streamlit Cloud.
3. Pilih main file `app.py`.
4. Deploy.
