# e-Praktika

e-Praktika Django asosidagi veb-tizim bo‘lib, talabalar amaliyoti uchun kerakli hujjatlarni bitta joyda yuklash, boshqarish va Word formatda yaratish uchun mo‘ljallangan.

## Loyiha maqsadi

Loyiha maqsadi Excel orqali talabalar ma’lumotlarini qabul qilish va har bir talaba uchun quyidagi hujjatlarni tayyorlash:

- `shartnoma.docx`
- `kundalik.docx`
- `yollanma.docx`

Natijada agar 15 ta talaba bo‘lsa, ZIP ichida 15 ta alohida papka yaratiladi va har bir papka ichida shu talaba uchun 3 ta hujjat bo‘ladi. Agar bir xil korxonada bir nechta talaba amaliyot o‘tayotgan bo‘lsa, shartnomadagi `Talabalar soni` maydoni shu korxona bo‘yicha umumiy sonni ko‘rsatadi.

## Asosiy imkoniyatlar

- `.xlsx` yoki `.docx` orqali talabalar ro‘yxatini yuklash
- bepul `Namuna Excel` yuklab olish
- namuna Excel ichida 15 ta default talaba ma’lumoti bilan tayyor jadval
- 3-kurs uchun `Ishlab chiqarish amaliyoti`
- 4-kurs uchun `Bitiruv oldi amaliyoti`
- har bir talaba uchun alohida hujjat generatsiyasi
- barcha hujjatlarni ZIP ko‘rinishida yuklab olish
- login, register, profil va akkaunt sozlamalari
- parolni yangilash
- REST API orqali talabalar CRUD

## Texnologiyalar

- `Python 3.12+`
- `Django 5.x`
- `Django REST Framework`
- `python-docx`
- `openpyxl`
- `HTML`, `CSS`, `JavaScript`
- `Jazzmin` admin panel

## Ishlash tartibi

1. Foydalanuvchi tizimga kiradi.
2. `Namuna Excel` faylini yuklab oladi yoki o‘z faylini yuklaydi.
3. Kurs va amaliyot sanalarini tanlaydi.
4. Tizim talabalarni bazaga saqlaydi.
5. Har bir talaba uchun `shartnoma`, `kundalik`, `yo‘llanma` tayyorlanadi.
6. ZIP ichida har bir talaba uchun alohida papka yaratiladi.

## ZIP tuzilmasi

```text
Hujjatlar/
├── 01_Aliyev_Bekzod_Anvar_o_g_li/
│   ├── shartnoma.docx
│   ├── kundalik.docx
│   └── yollanma.docx
├── 02_Karimova_Maftuna_Jamshid_qizi/
│   ├── shartnoma.docx
│   ├── kundalik.docx
│   └── yollanma.docx
```

## Asosiy sahifalar

- `/` - bosh sahifa
- `/excel/login/` - kirish
- `/excel/register/` - ro‘yxatdan o‘tish
- `/excel/` - dashboard
- `/excel/upload/` - Excel yoki Word yuklash
- `/excel/profile/` - profil
- `/excel/settings/` - akkaunt sozlamalari
- `/admin/` - admin panel

## Default foydalanuvchilar

`python manage.py create_admin` buyrug‘i quyidagi foydalanuvchilarni yaratadi:

- `admin / Admin12345`
- `operator / Operator12345`

## Loyihani ishga tushirish

```bash
git clone https://github.com/xavfli/amaliyotdocx.git
cd amaliyotdocx
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python manage.py migrate
python manage.py create_admin
python manage.py runserver
```

SQLite baza vaqtinchalik joyga yozilishi uchun kerak bo‘lsa:

```powershell
$env:SQLITE_DB_PATH = Join-Path $env:TEMP 'amaliyotdocx.sqlite3'
python manage.py runserver
```

## API endpointlar

```http
GET     /excel/api/students/
POST    /excel/api/students/
GET     /excel/api/students/<id>/
PUT     /excel/api/students/<id>/
DELETE  /excel/api/students/<id>/
```

Qo‘shimcha API:

```http
POST    /excel/api/token/
GET     /excel/api2/students/
POST    /excel/api2/students/
GET     /excel/api2/students/<id>/
PUT     /excel/api2/students/<id>/
DELETE  /excel/api2/students/<id>/
```

## Template fayllar

Loyiha hozir quyidagi haqiqiy Word blankalar bilan ishlaydi:

- `app_excel/templates/app_excel/shartnoma_template.docx`
- `app_excel/templates/app_excel/kundalik_template.docx`
- `app_excel/templates/app_excel/yollanma_template.docx`

Nusxalari quyidagi joyda ham mavjud:

- `app_shartnoma/contract_templates/shartnoma_template.docx`
- `app_shartnoma/contract_templates/kundalik_template.docx`
- `app_shartnoma/contract_templates/yollanma_template.docx`

Asl yuklangan manbalar:

- `app_excel/document_sources/Bitiruv oldi Shartnomasi.doc`
- `app_excel/document_sources/Bitiruv oldi amaliyot Kundalik.doc`
- `app_excel/document_sources/Bitiruv oldi Yo'llanma.doc`

## Muhim eslatmalar

- `Namuna Excel` ichida `Elektron pochta`, `Amaliyot boshlanish sanasi`, `Amaliyot tugash sanasi` ustunlari yo‘q
- sanalar yuklash formasida alohida tanlanadi
- hujjatlar `python-docx` orqali to‘ldiriladi

## Muallif

Boburbek (`xavfli`)


