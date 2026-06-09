# Payment Tracker Dashboard - Streamlit MVP

Dashboard ini dibuat untuk monitoring invoice payment tracker dari Excel.

## Fitur MVP
- Upload Excel tracker
- KPI Total Invoice, Total Amount, Overdue, Due <= 7 Days, Due <= 30 Days
- Filter Company, Brand, Currency, Vendor, Due Status
- Chart Outstanding by Vendor
- Chart Invoice Due Status
- Chart Monthly Due Amount
- Detail invoice table
- Download filtered data

## Cara Deploy Online via Streamlit Cloud

### 1. Buat akun GitHub
Buka github.com, login atau sign up.

### 2. Buat repository baru
Nama contoh:
payment-tracker-dashboard

### 3. Upload file ini ke repository
Upload:
- app.py
- requirements.txt

JANGAN upload file Excel asli kalau datanya confidential.

### 4. Buka Streamlit Cloud
Buka:
https://streamlit.io/cloud

Login pakai GitHub.

### 5. Create new app
Pilih:
- Repository: payment-tracker-dashboard
- Branch: main
- Main file path: app.py

Klik Deploy.

### 6. Setelah dashboard online
Buka URL Streamlit yang muncul.

Untuk MVP tercepat:
- pakai mode Upload Excel
- upload file Excel dari SharePoint
- dashboard langsung muncul

## Catatan SharePoint Refresh

Mode SharePoint Link akan berhasil hanya jika link dapat di-download langsung tanpa login tambahan.

Kalau link masih membutuhkan login Microsoft kantor, Streamlit Cloud tidak bisa membaca langsung tanpa setup authentication.

Solusi phase berikutnya:
1. Microsoft Graph API
2. Azure App Registration
3. Secret credentials disimpan di Streamlit Secrets
4. Dashboard bisa refresh langsung dari SharePoint secara secure

## Kolom yang dibaca dashboard
Dashboard otomatis menyesuaikan kolom, tetapi idealnya file Excel punya kolom:
- Posting Date
- Document Date
- Due Date
- Due Month
- Week
- External Document No.
- Vendor No.
- Vendor Name
- Company Code
- Brand Code
- Currency Code
- Original Amount