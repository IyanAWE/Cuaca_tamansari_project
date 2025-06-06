import streamlit as st
import requests
import pandas as pd
import json
import base64
from datetime import datetime
from streamlit_autorefresh import st_autorefresh
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from gspread_dataframe import set_with_dataframe
from pytz import timezone

# === DECODE BASE64 CREDENTIALS ===
def fix_padding(b64_string):
    return b64_string + "=" * (-len(b64_string) % 4)

b64_creds = st.secrets["GOOGLE_CREDS"]
fixed = fix_padding(b64_creds)
GOOGLE_CREDS = json.loads(base64.b64decode(fixed).decode("utf-8"))

# === KONFIGURASI ===
LAT = -6.90389
LON = 107.61861
SPREADSHEET_NAME = "Data Streamlit Cuaca Bandung"
OPENWEATHER_SPREADSHEET_ID = "1OuuELHENRXPTUzS6kvtAZaNYh1XNbblc3alX_VIN8qE"  # ID Spreadsheet OpenWeather
BMKG_SPREADSHEET_ID = "1Eac7sce0H0pkg3PQslBhjPcAc_5nMw-AFFZCgKUabNQ"
API_KEY = st.secrets["OPENWEATHER_API_KEY"]
wib = timezone("Asia/Jakarta")

st.set_page_config(page_title="Cuaca Tamansari", page_icon="🌧️", layout="wide")
st.title("🌧️ Dashboard Cuaca Tamansari: OpenWeather vs BMKG (OCR)")

refresh_trigger = st_autorefresh(interval=900000, key="data_refresh")

if 'data_history' not in st.session_state:
    st.session_state['data_history'] = []

def weather_emoji(desc):
    desc = desc.lower()
    if "thunderstorm" in desc:
        return "⛈️ " + desc
    elif "drizzle" in desc:
        return "🌦️ " + desc
    elif "rain" in desc:
        return "🌧️ " + desc
    elif "snow" in desc:
        return "❄️ " + desc
    elif any(x in desc for x in ["mist", "smoke", "haze", "fog"]):
        return "🌫️ " + desc
    elif any(x in desc for x in ["sand", "dust", "ash", "squall", "tornado"]):
        return "🌪️ " + desc
    elif "clear" in desc:
        return "☀️ " + desc
    elif "few clouds" in desc:
        return "🌤️ " + desc
    elif "scattered clouds" in desc:
        return "🌥️ " + desc
    elif "broken clouds" in desc or "overcast clouds" in desc:
        return "☁️ " + desc
    else:
        return "❓ " + desc
        
def get_google_sheet_data(spreadsheet_id):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(GOOGLE_CREDS, scope)
    client = gspread.authorize(creds)
    
    # Buka spreadsheet berdasarkan ID dan ambil sheet pertama
    sheet = client.open_by_key(spreadsheet_id).sheet1
    data = sheet.get_all_records()  # Ambil semua data sebagai list
    
    # Convert list ke DataFrame Pandas
    df = pd.DataFrame(data)

    # Hapus kolom 'Wind_kmh' agar tidak ditampilkan
    df = df.drop(columns=['Windspeed_(kmh)'], errors='ignore')
    
    return df


def simpan_ke_google_sheets(df):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(GOOGLE_CREDS, scope)
    client = gspread.authorize(creds)
    sheet = client.open(SPREADSHEET_NAME).sheet1

    existing = sheet.get_all_values()
    if not existing:
        sheet.append_row(list(df.columns))

    last_row = df.tail(1).values.tolist()[0]
    sheet.append_row(last_row)

def ambil_data_bmkg_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(GOOGLE_CREDS, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(BMKG_SPREADSHEET_ID).sheet1
    data = sheet.get_all_records()
    return pd.DataFrame(data)

def fetch_weather():
    url = f"https://api.openweathermap.org/data/3.0/onecall?lat={LAT}&lon={LON}&appid={API_KEY}&units=metric&lang=en"
    try:
        response = requests.get(url)
        data = response.json()
        current = data['current']
        temp = current['temp']
        humidity = current['humidity']
        desc = current['weather'][0]['description']
        icon = current['weather'][0]['icon']
        icon_url = f"https://openweathermap.org/img/wn/{icon}@2x.png"
        wind = current.get('wind_speed', 0)
        timestamp = datetime.now(wib).strftime('%Y-%m-%d %H:%M:%S')

        # Menambahkan data OpenWeather terbaru ke session state
        st.session_state['data_history'].append({
            'Time': timestamp,
            'Temperature': temp,
            'Humidity': humidity,
            'Weather': desc,
            'Wind_kmh': wind
        })

        df = pd.DataFrame(st.session_state['data_history'])
        simpan_ke_google_sheets(df)

        return temp, desc, humidity, wind, icon_url, timestamp, data
    except Exception as e:
        st.error(f"Gagal ambil data OpenWeather: {e}")
        return None, None, None, None, None, None, None

do_refresh = st.button("🔁 Refresh Now") or refresh_trigger > 0

if do_refresh or 'full_data' not in st.session_state:
    temp, desc, humidity, wind, icon_url, timestamp, full_data = fetch_weather()
    st.session_state['full_data'] = full_data
    st.session_state['temp'] = temp
    st.session_state['desc'] = desc
    st.session_state['humidity'] = humidity
    st.session_state['wind'] = wind
    st.session_state['icon_url'] = icon_url
    st.session_state['timestamp'] = timestamp
else:
    temp = st.session_state.get('temp')
    desc = st.session_state.get('desc')
    humidity = st.session_state.get('humidity')
    wind = st.session_state.get('wind')
    icon_url = st.session_state.get('icon_url')
    timestamp = st.session_state.get('timestamp')
    full_data = st.session_state.get('full_data')

col1, col2 = st.columns(2)

with col1:
    st.header("📡 OpenWeather (Live API)")
    if temp:
        st.image(icon_url, width=100)
        st.metric("🌡️ Temperature", f"{temp} °C")
        st.metric("💧 Humidity", f"{humidity}%")
        st.metric("🌬️ Wind Speed", f"{wind} km/h")
        st.markdown(f"### {weather_emoji(desc)}")
        st.caption(f"Last updated: {timestamp}")
    else:
        st.info("Belum ada data OpenWeather.")

with col2:
    st.header("🛰️ BMKG via OCR (Google Sheets)")
    try:
        df_bmkg = ambil_data_bmkg_sheet()
        if not df_bmkg.empty:
            latest = df_bmkg.tail(1).squeeze()
            st.metric("🌡️ Temperature", f"{latest['Temperature']} °C")
            st.metric("💧 Humidity", f"{latest['Humidity']}%")
            st.metric("🌬️ Wind Speed", f"{latest['Wind_kmh']} km/h")
            st.markdown(f"### ☁️ {latest['Weather']}")
            st.caption(f"Last updated: {latest['Time']}")
        else:
            st.info("Belum ada data dari BMKG.")
    except Exception as e:
        st.warning(f"⚠️ Gagal baca BMKG Real-Time: {e}")

# Tampilkan data dalam bentuk tabel
st.subheader("📊 Tabel Data Historis BMKG dan OpenWeather")
st.write("**Data Historis OpenWeatherMap**")
st.caption("Dikarenakan library streamlit yang digunakan untuk menampilkan tabel tidak bisa digunakan untuk membaca koma dari spreadsheet, maka:")
st.caption("- Cara membaca data suhu adalah menambahkan koma/titik setelah 2 angka pertama. Contoh: 2566 = 25.66, 245 = 24.5")
st.caption("- Kolom wind speed tidak diperlihatkan karena inkonsistensi pada pola numerik data, tidak seperti pada data suhu yang cenderung seragam.")
df_openweather = get_google_sheet_data(OPENWEATHER_SPREADSHEET_ID)
st.dataframe(df_openweather)  # Menampilkan data dari spreadsheet OpenWeather

st.write("**Data Historis BMKG**")
df_bmkg = ambil_data_bmkg_sheet()
st.dataframe(df_bmkg)  # Menampilkan data BMKG yang sudah difilter

# Grafik suhu historis
try:
    df_open = pd.DataFrame(st.session_state['data_history'])
    df_open["Time"] = pd.to_datetime(df_open["Time"], errors='coerce')
    df_open["Temperature"] = pd.to_numeric(df_open["Temperature"], errors='coerce')
    df_open = df_open.dropna(subset=["Time", "Temperature"])

    df_bmkg["Time"] = pd.to_datetime(df_bmkg["Time"], errors='coerce')
    df_bmkg["Temperature"] = pd.to_numeric(df_bmkg["Temperature"], errors='coerce')
    df_bmkg = df_bmkg.dropna(subset=["Time", "Temperature"])

    st.subheader("📈 Grafik Suhu Historis (12 data terakhir)")

    st.write("**OpenWeather**")
    df_open["TimeLabel"] = df_open["Time"].dt.strftime("%H:%M")
    st.line_chart(df_open.set_index("TimeLabel")[["Temperature"]].tail(12))

    st.write("**BMKG (OCR)**")
    df_bmkg["TimeLabel"] = df_bmkg["Time"].dt.strftime("%H:%M")
    st.line_chart(df_bmkg.set_index("TimeLabel")[["Temperature"]].tail(12))

except Exception as e:
    st.warning(f"⚠️ Gagal tampilkan grafik suhu: {e}")

st.caption("🔁 Auto-refresh tiap 30 menit || Kiri: OpenWeather API • Kanan: BMKG OCR + Grafik")
