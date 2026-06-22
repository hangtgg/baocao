# Web App Theo Doi Ho Tro Khach Hang

Web app noi bo de:

- Nhap luot ho tro khach hang hang ngay
- Xem, sua, xoa danh sach luot ho tro
- Loc du lieu theo khoang ngay, dich vu, trang thai, nguoi ho tro
- Xuat Excel tong hop theo bo loc hien tai

## Stack

- Backend: Python + Flask
- Database: SQLite
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
- `APP_DATA_DIR`: thu muc luu SQLite, mac dinh `./data`
- `DB_PATH`: duong dan day du toi file SQLite. Neu dat bien nay thi uu tien hon `APP_DATA_DIR`

Vi du:

```powershell
$env:APP_DATA_DIR="C:\tmp\support-data"
python app.py
```

## Du lieu

App luu du lieu trong SQLite:

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

## Deploy len Render

Repo da kem san file [render.yaml](./render.yaml) de deploy bang Render Blueprint.

### Cach nhanh nhat

1. Day code len GitHub.
2. Dang nhap Render.
3. Chon `New` -> `Blueprint`.
4. Chon repo nay.
5. Render se doc `render.yaml` va tao web service.

### Cau hinh Render da duoc them san

- Build command: `pip install -r requirements.txt`
- Start command: `gunicorn app:app --workers 1 --threads 4`
- Health check: `/`
- Persistent disk mount tai `/var/data`
- SQLite duoc luu trong `/var/data/support_tracker.db`

### Luu y quan trong voi SQLite tren Render

- Render mac dinh dung filesystem tam. Neu khong gan persistent disk, du lieu se mat sau moi lan redeploy hoac restart.
- Persistent disk chi dung duoc voi web service tra phi.
- File `render.yaml` da chon `plan: starter` va gan san disk de phu hop voi SQLite.

Neu ban muon dung goi free, nen chuyen sang Postgres thay vi SQLite.

## Nguon tham khao chinh cho Render

- Render Flask quickstart: `gunicorn app:app`
- Render persistent disk: filesystem mac dinh la ephemeral, chi data trong mount path moi duoc giu lai

Tai lieu chinh thuc:

- https://render.com/docs/deploy-flask
- https://render.com/docs/disks
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
- Neu deploy Render va dung SQLite, nen giu `WEB_CONCURRENCY=1` de tranh tang nguy co khoa file SQLite khi ghi dong thoi
