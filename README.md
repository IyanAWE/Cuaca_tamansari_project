# 🌧️ Real-Time Cuaca Tamansari, Bandung

Aplikasi Streamlit untuk menampilkan cuaca real-time di kawasan Tamansari, Bandung. Data diambil dari [OpenWeather](https://openweathermap.org/) dan disimpan otomatis ke Google Sheets tiap 10 menit. Dirancang untuk monitoring kondisi lingkungan berbasis data.

---

## 🚀 Fitur Utama

- 🔄 Auto-refresh data setiap 10 menit (tanpa reload manual)
- 🌡️ Tampilkan data suhu, kelembapan, kecepatan angin, dan deskripsi cuaca
- 📈 Visualisasi tren suhu dalam bentuk grafik
- 🗂️ Simpan histori data ke Google Sheets (append, bukan overwrite)
- ☁️ Integrasi API OpenWeather (versi One Call 3.0)
- ⏱️ Timestamp akurat dengan timezone Asia/Jakarta

---

## 🛠️ Teknologi

- [Streamlit](https://streamlit.io/) – Web UI framework
- [OpenWeather API](https://openweathermap.org/api/one-call-3) – Sumber data cuaca
- [Google Sheets API](https://developers.google.com/sheets/api) – Penyimpanan data historis
- `gspread`, `oauth2client`, `pytz`, dll.

---

## 📦 Instalasi

1. **Clone repositori**

```bash
git clone https://github.com/username/proyek-cuaca-tamansari.git
cd proyek-cuaca-tamansari
```

2. **Buat environment & install dependensi**

```bash
pip install -r requirements.txt
```

3. **Buat file `secrets.toml` untuk Streamlit**

Letakkan di `.streamlit/secrets.toml`:

```toml
OPENWEATHER_API_KEY = "your_openweather_api_key"
GOOGLE_CREDS = "base64_encoded_google_service_account_json"
```

> 🔐 Encode file JSON Google Service Account pakai:
> `base64 -w 0 your-credentials.json`

4. **Jalankan app**

```bash
streamlit run cuaca_tamansari_streamlit.py
```

---

## 📊 Contoh Tampilan

![Tampilan Cuaca](https://user-images.githubusercontent.com/your-ss.png)

---

## 📝 Lisensi

MIT License. Bebas digunakan untuk keperluan edukasi, riset, dan proyek lingkungan.

---

## 🙌 Kontribusi

Kalau Anda punya ide untuk integrasi dengan BMKG, prediksi cuaca jangka panjang, atau export CSV, feel free buat pull request atau open issue 🚀
