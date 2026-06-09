# 🚀 SETUP GUIDE — Material Status Dashboard

ทำตามขั้นตอนนี้ครั้งเดียว ใช้เวลาประมาณ 30 นาที

---

## ขั้นตอนที่ 1 — สร้าง Google Cloud Service Account

1. ไปที่ https://console.cloud.google.com
2. สร้าง Project ใหม่ (ชื่ออะไรก็ได้)
3. เปิด **APIs & Services → Library**
   - เปิดใช้ **Google Sheets API**
   - เปิดใช้ **Google Drive API**
4. ไปที่ **APIs & Services → Credentials**
   - คลิก **Create Credentials → Service Account**
   - ตั้งชื่อ เช่น `material-status-app`
   - คลิก **Done**
5. คลิกที่ Service Account ที่สร้าง → **Keys → Add Key → JSON**
   - ดาวน์โหลด file มา (ชื่อจะเป็น `xxxx.json`)
   - **rename เป็น `credentials.json`** แล้ววางไว้ใน folder `streamlit_app/`

---

## ขั้นตอนที่ 2 — เตรียมข้อมูลใน Google Sheets

เปิด Terminal / Command Prompt ในโฟลเดอร์ `streamlit_app`:

```bash
pip install gspread google-auth pandas openpyxl
python prepare_gsheet.py --excel "../Material_Status_Summary_Improved.xlsx" --creds credentials.json
```

Script จะสร้าง Google Sheet ชื่อ **"Material_Status_Dashboard"** และ copy ข้อมูลเข้าไป

> ⚠️ **สำคัญ:** Copy ค่า `client_email` จากไฟล์ `credentials.json` แล้ว share Google Sheet นั้นให้ email นั้นมีสิทธิ์ **Editor**

---

## ขั้นตอนที่ 3 — Deploy บน Streamlit Cloud (ฟรี)

### 3.1 Push โค้ดขึ้น GitHub
1. สร้าง GitHub repo ใหม่ (private ได้)
2. Push ไฟล์ทั้งหมดใน `streamlit_app/` ขึ้นไป
   - **อย่า** push ไฟล์ `credentials.json` (เพิ่มใน `.gitignore`)

```
streamlit_app/
├── app.py
├── requirements.txt
└── (อย่า push credentials.json)
```

### 3.2 สร้าง App บน Streamlit Cloud
1. ไปที่ https://share.streamlit.io → **New app**
2. เลือก GitHub repo และไฟล์ `app.py`
3. คลิก **Advanced settings → Secrets** แล้ววางข้อมูลนี้:

```toml
EDIT_PASSWORD = "ใส่ password ที่ต้องการ"

[gcp_service_account]
type = "service_account"
project_id = "your-project-id"
private_key_id = "..."
private_key = "-----BEGIN RSA PRIVATE KEY-----\n...\n-----END RSA PRIVATE KEY-----\n"
client_email = "material-status-app@xxx.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
```

> ค่าทั้งหมดอยู่ในไฟล์ `credentials.json` ที่ดาวน์โหลดมา

4. คลิก **Deploy** → รอ 2-3 นาที

---

## ขั้นตอนที่ 4 — แชร์ลิงค์

หลัง deploy เสร็จ จะได้ลิงค์แบบนี้:
```
https://your-app-name.streamlit.app
```

ส่งให้ทีมงานได้เลย — ใครที่มีลิงค์ดูได้ทันที, แก้ไขต้องใส่ password

---

## การอัพเดทข้อมูลในอนาคต

- ข้อมูลทั้งหมดอยู่ใน **Google Sheets "Material_Status_Dashboard"**
- แก้ไขได้ทั้งจาก App (กรอก password) หรือเปิด Google Sheets โดยตรง
- App refresh ข้อมูลใหม่ **ทุก 30 วินาที** อัตโนมัติ

---

## โครงสร้างไฟล์

```
streamlit_app/
├── app.py               ← โค้ดหลัก Streamlit
├── requirements.txt     ← Python packages
├── prepare_gsheet.py    ← Script เตรียม Google Sheets (รัน 1 ครั้ง)
├── secrets_template.toml← Template สำหรับ Streamlit secrets
└── SETUP_GUIDE.md       ← คู่มือนี้
```

---

## ปัญหาที่พบบ่อย

| ปัญหา | วิธีแก้ |
|-------|---------|
| `SpreadsheetNotFound` | ตรวจสอบชื่อ Sheet และ share ให้ service account email |
| `PERMISSION_DENIED` | เปิด Sheets API และ Drive API ใน Google Cloud |
| ข้อมูลไม่ update | กด Refresh Now หรือรอ 30 วินาที |
| Password ไม่ทำงาน | ตรวจสอบค่า `EDIT_PASSWORD` ใน Streamlit Secrets |
