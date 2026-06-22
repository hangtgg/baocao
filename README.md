# Web App Theo Doi Ho Tro Khach Hang

Web app noi bo de:

- Nhap luot ho tro khach hang hang ngay
- Xem, sua, xoa danh sach luot ho tro
- Loc du lieu theo khoang ngay, dich vu, trang thai, nguoi ho tro
- Xuat Excel tong hop theo bo loc hien tai

## Stack

- Backend: Python + Flask
- Database: PostgreSQL tren Render, SQLite khi chay local neu chua khai bao `DATABASE_URL`
- Frontend: HTML, CSS, JavaScript thuan
- Export Excel: `openpyxl`
- Deploy production: `gunicorn`

## Tinh nang hien co

- Form nhap luot ho tro hang ngay
- Danh sach luot ho tro co the chinh sua va xoa
- Truong `Nguoi gui ho tro` co the de trong va go tim nhanh
- Truong `Noi dung ho tro` co goi y theo dich vu
- Xuat file Excel gom:
  - `Chi tiet ho tro`
  - `Tong hop dich vu`
  - `Thong ke nguoi gui`
  - `Noi dung goi y`

## Cai dat va chay local

### 1. Tao moi truong ao

```powershell
python -m venv .venv
.venv\Scripts\activate
```

### 2. Cai thu vien

```powershell
pip install -r requirements.txt
```

### 3. Chay app

```powershell
python app.py
```

Mac dinh app chay tai:

```text
http://127.0.0.1:5000
```

Neu muon mo trong mang LAN:

```powershell
$env:APP_HOST="0.0.0.0"
python app.py
```

## Bien moi truong

- `APP_HOST`: host de Flask bind, mac dinh `0.0.0.0` khi chay file truc tiep
- `APP_PORT`: cong chay app, mac dinh `5000`
- `APP_DEBUG`: bat/tat debug, mac dinh `1` khi chay local
- `DATABASE_URL`: neu khai bao, app se dung PostgreSQL thay cho SQLite
- `APP_DATA_DIR`: thu muc luu SQLite, mac dinh `./data`
- `DB_PATH`: duong dan day du toi file SQLite. Neu dat bien nay thi uu tien hon `APP_DATA_DIR`

Vi du:

```powershell
$env:APP_DATA_DIR="C:\tmp\support-data"
python app.py
```

## Du lieu

Mac dinh khi chay local, app luu du lieu trong SQLite:

```text
data/support_tracker.db
```

Lan chay dau tien app tu dong:

- Tao bang du lieu
- Tao danh sach dich vu mac dinh
- Tao du lieu mau de test nhanh

Neu muon xoa sach du lieu local:

```powershell
Remove-Item data\support_tracker.db
python app.py
```

## Deploy len Render voi Postgres

App da duoc sua de:

- Uu tien dung Postgres neu co `DATABASE_URL`
- Tu dong fallback ve SQLite khi chay local ma chua khai bao `DATABASE_URL`
- Tren Render, neu thieu `DATABASE_URL` app se bao loi ngay de tranh ghi nham vao bo nho tam

Luu y:

- Du lieu SQLite cu trong `data/support_tracker.db` khong tu dong chuyen sang Postgres
- Lan deploy dau tien len Render se dung database Postgres moi, vi vay neu can giu lich su cu thi can import/migrate rieng

### Cach deploy de khong mat du lieu

1. Day code len GitHub.
2. Trong Render, tao truoc mot Postgres database.
3. Mo database vua tao va copy `Internal Database URL`.
4. Tao web service tu repo nay, hoac dung `Blueprint` voi file [render.yaml](./render.yaml).
5. Khi Render hoi bien moi truong `DATABASE_URL`, dan gia tri `Internal Database URL` vao.
6. Deploy.

### Cau hinh Render de nghi

- Web Service
  - Build command: `pip install -r requirements.txt`
  - Start command: `gunicorn app:app --workers 1 --threads 4`
  - Health check: `/`
- Environment variables
  - `DATABASE_URL`: URL ket noi Postgres cua Render
  - `APP_DEBUG=0`
  - `WEB_CONCURRENCY=1`

### Render Blueprint

File [render.yaml](./render.yaml) da duoc doi sang mo hinh Postgres:

- Web service plan `free`
- Khong dung persistent disk nua
- `DATABASE_URL` duoc de `sync: false` de ban nhap gia tri khi tao service

### Tai sao khong con bi mat du lieu

- Truoc day, app dung SQLite tren filesystem cuc bo cua Render.
- Render web service co filesystem tam, nen sau restart hoac redeploy du lieu co the mat.
- Khi chuyen sang Postgres, du lieu duoc luu trong database quan ly rieng cua Render, khong phu thuoc vao filesystem cua web service.

## Nguon tham khao chinh cho Render

- Render Flask quickstart: `gunicorn app:app`
- Render web services co the nhan `PORT`
- Render Postgres dung rieng voi web service, phu hop hon SQLite khi can giu du lieu qua deploy

Tai lieu chinh thuc:

- https://render.com/docs/deploy-flask
- https://render.com/docs/environment-variables
- https://render.com/docs/blueprint-spec

## Cau truc du an

```text
BAOCAO/
|-- app.py
|-- render.yaml
|-- requirements.txt
|-- README.md
|-- static/
|   |-- app.js
|   `-- styles.css
|-- templates/
|   `-- index.html
`-- data/
    `-- support_tracker.db
```

## Ghi chu

- App hien toi uu cho may tinh
- Neu dung Postgres tren Render, ban khong can persistent disk cho du lieu nghiep vu nua
- Van nen giu `WEB_CONCURRENCY=1` cho app nho de de theo doi va tranh phat sinh ghi dong thoi khong can thiet
