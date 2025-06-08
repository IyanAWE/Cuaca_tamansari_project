from playwright.sync_api import sync_playwright
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2 import service_account
from datetime import datetime
import os
import pytesseract
import cv2
import numpy as np
import re
import gspread
from gspread_dataframe import set_with_dataframe
import pandas as pd
import shutil
from PIL import Image

# === KONFIGURASI ===
SERVICE_ACCOUNT_FILE = "D:/DATA WEB THING/plated-magpie-458707-q3-6425b9bf77f5.json"
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
SPREADSHEET_ID = "1Eac7sce0H0pkg3PQslBhjPcAc_5nMw-AFFZCgKUabNQ"
SHEET_NAME = "Sheet1"
DRIVE_FOLDER_ID = "1eY2UZbd8QUA0p2EvmHUwjPpHFcgfePis"
CROP_HEIGHT = 850  # tinggi pixel bagian atas BMKG yang kita ambil

SCOPES = [
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/spreadsheets'
]

pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

# === AUTENTIKASI ===
def authenticate():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=SCOPES
    )
    return creds

# === CROP GAMBAR ATAS SAJA ===
def crop_image(image_path, output_path, height=CROP_HEIGHT):
    img = Image.open(image_path)
    cropped = img.crop((0, 0, img.width, height))
    cropped.save(output_path)

# === UPLOAD FILE KE GOOGLE DRIVE ===
def upload_screenshot(file_path, creds):
    service = build('drive', 'v3', credentials=creds)
    filename = os.path.basename(file_path)
    file_metadata = {'name': filename, 'parents': [DRIVE_FOLDER_ID]}
    media = MediaFileUpload(file_path, mimetype='image/png')
    uploaded_file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()
    print(f"✅ File '{filename}' berhasil diupload ke Google Drive (ID: {uploaded_file['id']})")

# === EKSTRAK DATA CUACA DARI GAMBAR ===
def extract_metrics(image_path):
    img = cv2.imread(image_path)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    text = pytesseract.image_to_string(gray)

    def extract(pattern):
        match = re.search(pattern, text, re.IGNORECASE)
        return match.group(1).strip() if match else None

    temp = extract(r"(\d+)\s*[°oO]?\s*[Cc]")
    humidity = extract(r"Kelembapan[:\s]*([0-9]+)%")
    wind = extract(r"Angin[:\s]*([0-9]+)\s*km/jam")

    weather_keywords = ["Cerah", "Berawan", "Hujan", "Mendung", "Kabut"]
    weather = None
    for word in weather_keywords:
        if re.search(word, text, re.IGNORECASE):
            weather = word.capitalize()
            break

    return temp, humidity, wind, weather

# === SIMPAN KE GOOGLE SHEETS ===
def simpan_ke_sheets(creds, data):
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    worksheet = sh.worksheet(SHEET_NAME)

    df = pd.DataFrame([data])
    values = worksheet.get_all_values()

    if not values or len(values) == 0:
        set_with_dataframe(worksheet, df)
    else:
        worksheet.append_row(df.iloc[0].tolist())

# === MAIN ===
timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
filename = f"bmkg_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
save_path = os.path.join(".", filename)
crop_path = f"cropped_{filename}"

# Step 1: Ambil screenshot BMKG
with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page(viewport={"width": 1280, "height": 1000})
    page.goto("https://www.bmkg.go.id/cuaca/prakiraan-cuaca/32.73.09.1003")
    page.wait_for_timeout(3000)
    page.screenshot(path=save_path)
    browser.close()

# Step 2: Crop bagian atas saja
crop_image(save_path, crop_path)

# Step 3: Autentikasi
creds = authenticate()

# Step 4: OCR dari hasil crop
temp, humidity, wind, weather = extract_metrics(crop_path)
print(f"✅ OCR Data: Temp={temp}, Humidity={humidity}, Wind={wind}, Weather={weather}")

# Step 5: Simpan ke Google Sheets
data = {
    "Time": timestamp,
    "Temperature": temp,
    "Humidity": humidity,
    "Weather": weather,
    "Wind_kmh": wind
}
simpan_ke_sheets(creds, data)

# Step 6: Upload hasil crop ke Google Drive
upload_screenshot(crop_path, creds)

# Step 7: Arsipkan hasil crop
arsip_folder = "D:/DATA WEB THING/arsip_screenshot"
os.makedirs(arsip_folder, exist_ok=True)
new_path = os.path.join(arsip_folder, os.path.basename(crop_path))
shutil.copy2(crop_path, new_path)

# Cleanup file sementara
os.remove(save_path)
os.remove(crop_path)

print(f"📁 File dipindahkan ke: {new_path}")
