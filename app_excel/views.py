from docx import Document
from docxtpl import DocxTemplate
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
import openpyxl
from django.http import HttpResponse, FileResponse
from rest_framework.decorators import api_view
from .models import Student, Profile
import io, os, zipfile
from django.conf import settings
from datetime import datetime
from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth import update_session_auth_hash
from django.contrib.auth.models import User
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator
from .serializers import StudentSerializer
from rest_framework import viewsets
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status
from rest_framework.authtoken.models import Token


def ensure_profile(user):
    return Profile.objects.get_or_create(user=user)[0]


def get_template_sources():
    source_dir = os.path.join(settings.BASE_DIR, "app_excel", "document_sources")
    return [
        {
            "key": "shartnoma",
            "label": "Shartnoma",
            "filename": "Bitiruv oldi Shartnomasi.doc",
            "kind": "DOC manba",
            "exists": os.path.exists(os.path.join(source_dir, "Bitiruv oldi Shartnomasi.doc")),
        },
        {
            "key": "kundalik",
            "label": "Kundalik",
            "filename": "Bitiruv oldi amaliyot Kundalik.doc",
            "kind": "DOC manba",
            "exists": os.path.exists(os.path.join(source_dir, "Bitiruv oldi amaliyot Kundalik.doc")),
        },
        {
            "key": "yollanma",
            "label": "Yo‘llanma",
            "filename": "Bitiruv oldi Yo'llanma.doc",
            "kind": "DOC manba",
            "exists": os.path.exists(os.path.join(source_dir, "Bitiruv oldi Yo'llanma.doc")),
        },
    ]


def get_runtime_template_paths():
    template_dir = os.path.join(settings.BASE_DIR, "app_excel", "templates", "app_excel")
    return {
        "shartnoma": os.path.join(template_dir, "shartnoma_template.docx"),
        "kundalik": os.path.join(template_dir, "kundalik_template.docx"),
        "yollanma": os.path.join(template_dir, "yollanma_template.docx"),
    }


def validate_runtime_templates():
    missing = [
        name for name, path in get_runtime_template_paths().items()
        if not os.path.exists(path)
    ]
    return missing


def get_available_runtime_template_path(template_key):
    path = get_runtime_template_paths()[template_key]
    return path if os.path.exists(path) else None


def get_practice_type(course):
    return "Ishlab chiqarish amaliyoti" if int(course) == 3 else "Bitiruv oldi amaliyoti"


REQUIRED_STUDENT_FIELDS = [
    ("full_name", "Talabaning F.I.Sh."),
    ("faculty", "Fakultet"),
    ("direction", "Yo'nalish"),
    ("group", "Guruh"),
    ("company", "Korxona nomi"),
    ("company_address", "Korxona manzili"),
    ("company_director", "Korxona rahbari F.I.Sh."),
    ("company_phone", "Korxona telefoni"),
    ("practice_supervisor", "Universitet amaliyot rahbari"),
    ("faculty_dean", "Fakultet dekani"),
    ("department_head", "Kafedra mudiri"),
]


def get_missing_student_fields(student):
    missing = []
    for field_name, label in REQUIRED_STUDENT_FIELDS:
        value = getattr(student, field_name, "")
        if not str(value or "").strip():
            missing.append(label)
    return missing


def validate_students_for_documents(students, course, start_date_str, end_date_str):
    issues = []
    if not course:
        issues.append("Kurs tanlanmagan.")
    if not start_date_str or not end_date_str:
        issues.append("Boshlanish va tugash sanalari to'liq kiritilmagan.")

    for index, student in enumerate(students, start=1):
        missing = get_missing_student_fields(student)
        if missing:
            issues.append(f"{index}. {student.full_name}: {', '.join(missing)}")
    return issues


def sanitize_filename(value):
    safe = "".join(char if char.isalnum() or char in (" ", "-", "_") else "_" for char in (value or "").strip())
    return "_".join(safe.split()) or "talaba"


def normalize_excel_header(value):
    return " ".join(str(value or "").strip().lower().split())


def get_excel_value(row_map, *aliases):
    for alias in aliases:
        normalized_alias = normalize_excel_header(alias)
        if normalized_alias in row_map:
            return str(row_map.get(normalized_alias) or "").strip()
    return ""


def normalize_excel_date(value):
    if value in (None, ""):
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if hasattr(value, "strftime"):
        try:
            return value.strftime("%Y-%m-%d")
        except Exception:
            pass
    raw = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(raw, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return raw


def get_company_student_count(company_name):
    return Student.objects.filter(company=company_name).count()


def build_student_context(student, course, practice_type, start_date, end_date):
    return {
        'FULL_NAME': student.full_name,
        'DIRECTION': student.direction or student.faculty,
        'GROUP': student.group,
        'COMPANY': student.company,
        'ADDRESS': student.company_address,
        'DIRECTOR': student.company_director,
        'PHONE': student.company_phone,
        'SUPERVISOR': student.practice_supervisor,
        'FACULTY': student.faculty,
        'COURSE': course,
        'PRACTICE_TYPE': practice_type,
        'STUDENT_COUNT': get_company_student_count(student.company),
        'START_DATE': start_date,
        'END_DATE': end_date,
    }


def save_document_response(document, filename):
    buffer = io.BytesIO()
    document.save(buffer)
    buffer.seek(0)
    response = HttpResponse(
        buffer.read(),
        content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    return response


def parse_date_parts(date_str):
    date_obj = datetime.strptime(date_str, "%Y-%m-%d")
    oylar = [
        "yanvar", "fevral", "mart", "aprel", "may", "iyun",
        "iyul", "avgust", "sentabr", "oktabr", "noyabr", "dekabr"
    ]
    return {
        "day": f"{date_obj.day:02}",
        "month": oylar[date_obj.month - 1],
        "year": str(date_obj.year),
        "short": date_obj.strftime("%d.%m.%Y"),
    }


def remove_cell_shading(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    for child in list(tc_pr):
        if child.tag == qn('w:shd'):
            tc_pr.remove(child)


def clear_run_highlight(paragraph):
    for run in paragraph.runs:
        run.font.highlight_color = None
        run.font.color.rgb = RGBColor(0, 0, 0)


def style_cell(cell, horizontal=WD_ALIGN_PARAGRAPH.CENTER, vertical=WD_ALIGN_VERTICAL.CENTER, font_size_pt=None):
    remove_cell_shading(cell)
    cell.vertical_alignment = vertical
    for paragraph in cell.paragraphs:
        paragraph.alignment = horizontal
        clear_run_highlight(paragraph)
        if font_size_pt is not None:
            for run in paragraph.runs:
                run.font.size = Pt(font_size_pt)


def center_paragraph(paragraph, font_size_pt=None):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    clear_run_highlight(paragraph)
    if font_size_pt is not None:
        for run in paragraph.runs:
            run.font.size = Pt(font_size_pt)


def normalize_document_formatting(doc):
    for paragraph in doc.paragraphs:
        clear_run_highlight(paragraph)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                remove_cell_shading(cell)
                for paragraph in cell.paragraphs:
                    clear_run_highlight(paragraph)


def clear_paragraph_content(paragraph):
    for run in list(paragraph.runs):
        paragraph._p.remove(run._element)


def add_styled_run(paragraph, text, *, underline=False, italic=False, font_size_pt=None):
    run = paragraph.add_run(text)
    run.underline = underline
    run.italic = italic
    run.font.color.rgb = RGBColor(0, 0, 0)
    if font_size_pt is not None:
        run.font.size = Pt(font_size_pt)
    return run


def set_paragraph_segments(paragraph, segments, *, font_size_pt=None):
    alignment = paragraph.alignment
    clear_paragraph_content(paragraph)
    for segment in segments:
        if len(segment) == 2:
            text, underline = segment
            italic = False
        else:
            text, underline, italic = segment
        add_styled_run(paragraph, text, underline=underline, italic=italic, font_size_pt=font_size_pt)
    paragraph.alignment = alignment


def format_document_date(date_str):
    date_obj = datetime.strptime(date_str, "%Y-%m-%d")
    oylar = [
        "yanvar", "fevral", "mart", "aprel", "may", "iyun",
        "iyul", "avgust", "sentabr", "oktabr", "noyabr", "dekabr"
    ]
    return f"{date_obj.year}-yil “{date_obj.day}” {oylar[date_obj.month - 1]}"


def build_fallback_document(title, student, course, practice_type, start_date_str, end_date_str):
    doc = Document()
    doc.add_heading(title, level=0)
    doc.add_paragraph(f"Talaba: {student.full_name}")
    doc.add_paragraph(f"Yo'nalish/Fakultet: {student.direction or student.faculty}")
    doc.add_paragraph(f"Guruh: {student.group}")
    doc.add_paragraph(f"Kurs: {course}")
    doc.add_paragraph(f"Amaliyot turi: {practice_type}")
    doc.add_paragraph(f"Korxona: {student.company}")
    doc.add_paragraph(f"Korxona manzili: {student.company_address}")
    doc.add_paragraph(f"Korxona rahbari: {student.company_director}")
    doc.add_paragraph(f"Universitet rahbari: {student.practice_supervisor}")
    doc.add_paragraph(f"Telefon: {student.company_phone}")
    doc.add_paragraph(f"Boshlanish sanasi: {start_date_str}")
    doc.add_paragraph(f"Tugash sanasi: {end_date_str}")

    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    header = table.rows[0].cells
    header[0].text = "Maydon"
    header[1].text = "Qiymat"
    rows = [
        ("F.I.Sh.", student.full_name),
        ("Yo'nalish", student.direction or student.faculty),
        ("Guruh", student.group),
        ("Korxona", student.company),
        ("Manzil", student.company_address),
        ("Rahbar", student.company_director),
        ("Supervisor", student.practice_supervisor),
        ("Sana", f"{start_date_str} - {end_date_str}"),
    ]
    for key, value in rows:
        cells = table.add_row().cells
        cells[0].text = str(key)
        cells[1].text = str(value)

    return doc


def generate_shartnoma_document(student, course, practice_type, start_date_str, end_date_str):
    template_path = get_available_runtime_template_path("shartnoma")
    if not template_path:
        return build_fallback_document("Shartnoma", student, course, practice_type, start_date_str, end_date_str)

    doc = Document(template_path)
    normalize_document_formatting(doc)
    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)
    company_count = get_company_student_count(student.company)

    doc.paragraphs[0].text = f"TALABALARNING {practice_type.upper()}NI TASHKIL ETISH BO‘YICHA SHARTNOMA"
    doc.paragraphs[3].text = f"Toshkent sh.                                                                 {start_parts['year']}-yil «{start_parts['day']}» {start_parts['month']}"
    doc.paragraphs[5].text = (
        "Biz, quyida imzo chekuvchilar, Muhammad Al-Xorazmiy nomidagi Toshkent axborot texnologiyalari universiteti "
        f"(keying o‘rinlarda – Universitet)dan tayinlangan shaxs K.Tashev - o‘quv ishlari bo‘yicha prorektori bir tomondan, {student.company}"
    )
    doc.paragraphs[7].text = (
        f"(keying o‘rinlarda – Korxona)dan tayinlangan shaxs {student.company_director}"
    )

    main_table = doc.tables[0]
    main_table.rows[2].cells[1].text = student.direction or student.faculty
    main_table.rows[2].cells[2].text = str(course)
    main_table.rows[2].cells[3].text = practice_type
    main_table.rows[2].cells[4].text = str(company_count)
    main_table.rows[2].cells[5].text = start_parts["short"]
    main_table.rows[2].cells[6].text = end_parts["short"]
    style_cell(main_table.rows[2].cells[0], horizontal=WD_ALIGN_PARAGRAPH.CENTER, font_size_pt=11)
    style_cell(main_table.rows[2].cells[2], horizontal=WD_ALIGN_PARAGRAPH.CENTER, font_size_pt=11)
    style_cell(main_table.rows[2].cells[3], horizontal=WD_ALIGN_PARAGRAPH.CENTER, font_size_pt=11)
    style_cell(main_table.rows[2].cells[4], horizontal=WD_ALIGN_PARAGRAPH.CENTER, font_size_pt=11)
    style_cell(main_table.rows[2].cells[5], horizontal=WD_ALIGN_PARAGRAPH.CENTER, font_size_pt=10)
    style_cell(main_table.rows[2].cells[6], horizontal=WD_ALIGN_PARAGRAPH.CENTER, font_size_pt=10)
    main_table.rows[2].cells[1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    for paragraph in main_table.rows[2].cells[1].paragraphs:
        clear_run_highlight(paragraph)

    info_table = doc.tables[1]
    company_block = f"Korxona\n{student.company}\n{student.company_address}\nTel.: {student.company_phone}"
    info_table.rows[0].cells[1].text = company_block
    info_table.rows[0].cells[2].text = company_block
    info_table.rows[4].cells[2].text = (
        "Korxonadan ajratilgan\n"
        f"amaliyot rahbari ________ {student.company_director}"
    )
    style_cell(info_table.rows[0].cells[1], horizontal=WD_ALIGN_PARAGRAPH.CENTER, font_size_pt=10)
    style_cell(info_table.rows[0].cells[2], horizontal=WD_ALIGN_PARAGRAPH.CENTER, font_size_pt=10)
    style_cell(info_table.rows[4].cells[2], horizontal=WD_ALIGN_PARAGRAPH.CENTER, font_size_pt=10)
    return doc


def generate_kundalik_document(student, course, practice_type, start_date_str, end_date_str):
    template_path = get_available_runtime_template_path("kundalik")
    if not template_path:
        return build_fallback_document("Kundalik", student, course, practice_type, start_date_str, end_date_str)

    doc = Document(template_path)
    normalize_document_formatting(doc)
    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)
    academic_year = f"{start_parts['year']}/{int(start_parts['year']) + 1} o‘quv yili"

    doc.paragraphs[5].text = (
        f"{student.direction or student.faculty} ta’lim yo‘nalishi {course} - bosqich talabasi {student.full_name} ning "
        f"{academic_year}dagi {practice_type.lower()}"
    )
    doc.paragraphs[10].text = (
        f"1.1. Amaliyot joyi va muddati {student.company}, {student.company_address} muddati: "
        f"{start_parts['day']} {start_parts['month']} {start_parts['year']} dan "
        f"{end_parts['day']} {end_parts['month']} {end_parts['year']} gacha"
    )
    doc.paragraphs[12].text = f"Universitetdan {student.practice_supervisor}"
    doc.paragraphs[14].text = f"Korxonadan {student.company_director}"
    doc.paragraphs[16].text = (
        f"1.3. Talaba {student.full_name} ga “{student.faculty}” kafedrasidan berilgan individual topshiriqlar "
        f"{practice_type.lower()} bo'yicha kundalik yuritish va hisobot tayyorlash."
    )
    doc.paragraphs[17].text = (
        f"1.4. Amaliyotga keldi: {start_parts['year']}-yil «{start_parts['day']}» {start_parts['month']}, "
        f"ketdi: {end_parts['year']}-yil «{end_parts['day']}» {end_parts['month']}"
    )
    if len(doc.paragraphs) > 44:
        doc.paragraphs[42].text = f"Universitetdan amaliyot rahbari        ____________________    {student.practice_supervisor}"
        doc.paragraphs[43].text = f"Kafedra mudiri                         ____________________    {student.department_head}"
    if course == 3:
        doc.paragraphs[25].text = "2.6. Amaliyot muddati 3-bosqichda o'quv reja asosida belgilanadi."
    for idx in (12, 14):
        if idx < len(doc.paragraphs):
            center_paragraph(doc.paragraphs[idx], font_size_pt=11)
    for idx in (13, 15):
        if idx < len(doc.paragraphs):
            center_paragraph(doc.paragraphs[idx], font_size_pt=9)
    return doc


def generate_yollanma_document(student, course, practice_type, start_date_str, end_date_str):
    template_path = get_available_runtime_template_path("yollanma")
    if not template_path:
        return build_fallback_document("Yo'llanma", student, course, practice_type, start_date_str, end_date_str)

    doc = Document(template_path)
    normalize_document_formatting(doc)
    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)

    doc.paragraphs[4].text = (
        "Muhammad al-Xorazmiy nomidagi Toshkent axborot texnologiyalari universiteti O‘zbekiston Respublikasi "
        "Oliy taʼlim, fan va innovatsiyalar vazirligi tomonidan tasdiqlangan “Oliy ta’lim muassasalari "
        f"talabalarining malaka amaliyotini o‘tash tartibi to‘g‘risidagi nizom” va {student.company}"
    )
    doc.paragraphs[6].text = (
        f"{student.company} hamda Muhammad al-Xorazmiy nomidagi TATU o‘rtasidagi shartnoma asosida"
    )
    doc.paragraphs[7].text = f"{student.direction or student.faculty} ta’lim yo‘nalishida tahsil olayotgan talaba"
    doc.paragraphs[9].text = f"{student.full_name} {practice_type.lower()}ni o‘tash uchun"
    doc.paragraphs[12].text = f"{student.company} ga yubormoqda."
    doc.paragraphs[15].text = (
        f"Amaliyot muddati:     “{start_parts['day']}” {start_parts['month']} {start_parts['year']}-yildan,       "
        f"“{end_parts['day']}” {end_parts['month']} {end_parts['year']} yilgacha"
    )
    doc.paragraphs[37].text = student.full_name
    doc.paragraphs[39].text = f"{student.company} ga  amaliyotga keldi."
    for idx in (5, 8, 10, 13, 37):
        if idx < len(doc.paragraphs):
            center_paragraph(doc.paragraphs[idx], font_size_pt=11)
    for idx in (6, 9, 38):
        if idx < len(doc.paragraphs):
            center_paragraph(doc.paragraphs[idx], font_size_pt=9)
    return doc


@login_required(login_url='login')
def dashboard_view(request):
    profile = ensure_profile(request.user)
    students = Student.objects.all()
    student_page = Paginator(students.order_by("-id"), 3).get_page(request.GET.get("students_page"))
    context = {
        "profile": profile,
        "student_count": students.count(),
        "enterprise_count": students.values("company").distinct().count(),
        "document_count": students.count() * 3,
        "recent_students": student_page.object_list,
        "student_page": student_page,
    }
    return render(request, "app_excel/dashboard.html", context)





class StudentListCreateAPIView(APIView):
    def get(self, request):
        students = Student.objects.all()
        serializer = StudentSerializer(students, many=True)
        return Response(serializer.data)

    def post(self, request):
        serializer = StudentSerializer(data=request.data)
        if serializer.is_valid():
            serializer.save()
            return Response(serializer.data, status=status.HTTP_201_CREATED)
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)


class StudentRetrieveUpdateDestroyAPIView(APIView):
    def get_object(self, pk):
        try:
            return Student.objects.get(pk=pk)
        except Student.DoesNotExist:
            return None

    def get(self, request, pk):
        student = self.get_object(pk)
        if not student:
            return Response(status=status.HTTP_404_NOT_FOUND)
        serializer = StudentSerializer(student)
        return Response(serializer.data)

    def put(self, request, pk):
        student = self.get_object(pk)
        if not student:
            return Response(status=status.HTTP_404_NOT_FOUND)
        serializer = StudentSerializer(student, data=request.data)
        if serializer.is_valid():
            serializer.save()
            return Response(serializer.data)
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

    def delete(self, request, pk):
        student = self.get_object(pk)
        if not student:
            return Response(status=status.HTTP_404_NOT_FOUND)
        student.delete()
        return Response(status=status.HTTP_204_NO_CONTENT)

class StudentViewSet(viewsets.ModelViewSet):
    queryset = Student.objects.all()
    serializer_class = StudentSerializer


class StudentListCreateAPIView(APIView):
    def get(self, request):
        students = Student.objects.all()
        serializer = StudentSerializer(students, many=True)
        return Response(serializer.data)

    def post(self, request):
        serializer = StudentSerializer(data=request.data)
        if serializer.is_valid():
            serializer.save()
            return Response(serializer.data, status=status.HTTP_201_CREATED)
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)



class StudentRetrieveUpdateDestroyAPIView(APIView):
    def get_object(self, pk):
        try:
            return Student.objects.get(pk=pk)
        except Student.DoesNotExist:
            return None

    def get(self, request, pk):
        student = self.get_object(pk)
        if not student:
            return Response(status=status.HTTP_404_NOT_FOUND)
        serializer = StudentSerializer(student)
        return Response(serializer.data)

    def put(self, request, pk):
        student = self.get_object(pk)
        if not student:
            return Response(status=status.HTTP_404_NOT_FOUND)
        serializer = StudentSerializer(student, data=request.data)
        if serializer.is_valid():
            serializer.save()
            return Response(serializer.data)
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

    def delete(self, request, pk):
        student = self.get_object(pk)
        if not student:
            return Response(status=status.HTTP_404_NOT_FOUND)
        student.delete()
        return Response(status=status.HTTP_204_NO_CONTENT)



def format_uzbek_date(date_str):
    try:
        date_obj = datetime.strptime(date_str, "%Y-%m-%d")
        oylar = [
            "yanvar", "fevral", "mart", "aprel", "may", "iyun",
            "iyul", "avgust", "sentabr", "oktabr", "noyabr", "dekabr"
        ]
        return f"{date_obj.year}-yil «{date_obj.day}» {oylar[date_obj.month - 1]}"
    except Exception:
        return ""


@login_required(login_url='login')
def download_sample_excel(request):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Talabalar"

    headers = [
        "Talabaning F.I.Sh.",
        "Fakultet",
        "Yo'nalish",
        "Guruh",
        "Kurs",
        "Amaliyot turi",
        "Boshlanish sanasi",
        "Tugash sanasi",
        "Korxona nomi",
        "Korxona manzili",
        "Korxona rahbari F.I.Sh.",
        "Universitet mas'ul rahbari",
        "Telefon raqami",
        "Fakultet dekani",
        "Kafedra mudiri",
        "Korxonadagi lavozimi",
        "Izoh",
    ]

    sample_rows = [
        ["Aliyev Bekzod Anvar o'g'li", "Axborot texnologiyalari", "Dasturiy injiniring", "SE-401", 4, "Bitiruv oldi amaliyoti", "2025-02-17", "2025-04-26", "TechSoft MCHJ", "Toshkent sh., Yunusobod tumani", "Karimov Sardor Rustamovich", "Rasulov Dilshod Qodirovich", "+998901112233", "O.B. Ro'zibayev", "N.O. Raximov", "Amaliyot rahbari", "Barcha ustunlar majburiy"],
        ["Karimova Maftuna Jamshid qizi", "Axborot texnologiyalari", "Kompyuter injiniring", "KI-402", 4, "Bitiruv oldi amaliyoti", "2025-02-17", "2025-04-26", "Digital Systems", "Toshkent sh., Chilonzor tumani", "Ergashev Oybek Bahodirovich", "Rasulov Dilshod Qodirovich", "+998901112234", "O.B. Ro'zibayev", "N.O. Raximov", "Amaliyot rahbari", "Bo'sh qoldirmang"],
        ["Toshpo'latov Azizbek Shavkat o'g'li", "Kompyuter texnologiyalari", "Axborot xavfsizligi", "AT-403", 4, "Bitiruv oldi amaliyoti", "2025-02-17", "2025-04-26", "SecureNet Group", "Toshkent sh., Shayxontohur tumani", "Mamatqulov Anvar Tohirovich", "Sobirova Nargiza Akmalovna", "+998901112235", "O.B. Ro'zibayev", "N.O. Raximov", "Amaliyot rahbari", "Telefon formatini saqlang"],
    ]

    header_fill = openpyxl.styles.PatternFill("solid", fgColor="4338CA")
    header_font = openpyxl.styles.Font(bold=True, color="FFFFFF")
    header_alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_side = openpyxl.styles.Side(style="thin", color="000000")
    medium_side = openpyxl.styles.Side(style="medium", color="000000")
    data_alignment = openpyxl.styles.Alignment(vertical="center", horizontal="left")
    data_fill = openpyxl.styles.PatternFill("solid", fgColor="F8FAFC")

    for index, header in enumerate(headers, start=1):
        cell = sheet.cell(row=1, column=index, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = openpyxl.styles.Border(
            left=medium_side,
            right=medium_side,
            top=medium_side,
            bottom=medium_side,
        )
        sheet.column_dimensions[openpyxl.utils.get_column_letter(index)].width = max(len(header) + 4, 20)

    for row_index, row_data in enumerate(sample_rows, start=2):
        for column_index, value in enumerate(row_data, start=1):
            cell = sheet.cell(row=row_index, column=column_index, value=value)
            cell.alignment = data_alignment
            cell.fill = data_fill
            cell.border = openpyxl.styles.Border(
                left=thin_side,
                right=thin_side,
                top=thin_side,
                bottom=thin_side,
            )

    for row in range(2, len(sample_rows) + 2):
        sheet.row_dimensions[row].height = 22

    guide_sheet = workbook.create_sheet("Yoriqnoma")
    guide_rows = [
        ["Ustun", "Izoh", "Majburiy"],
        ["Talabaning F.I.Sh.", "Talabaning to'liq F.I.Sh.", "Ha"],
        ["Fakultet", "Masalan: Axborot texnologiyalari", "Ha"],
        ["Yo'nalish", "Masalan: Dasturiy injiniring", "Ha"],
        ["Guruh", "Masalan: SE-401", "Ha"],
        ["Kurs", "3 yoki 4", "Ha"],
        ["Amaliyot turi", "Masalan: Bitiruv oldi amaliyoti", "Ha"],
        ["Boshlanish sanasi", "YYYY-MM-DD formatda", "Ha"],
        ["Tugash sanasi", "YYYY-MM-DD formatda", "Ha"],
        ["Korxona nomi", "Hujjatlarda chiqadigan tashkilot nomi", "Ha"],
        ["Korxona manzili", "To'liq manzil", "Ha"],
        ["Korxona rahbari F.I.Sh.", "Korxona rahbari yoki korxonadagi amaliyot rahbari", "Ha"],
        ["Universitet mas'ul rahbari", "Universitetdagi amaliyot rahbari", "Ha"],
        ["Telefon raqami", "Masalan: +998901112233", "Ha"],
        ["Fakultet dekani", "Yo'llanma uchun kerak", "Ha"],
        ["Kafedra mudiri", "Kundalik va yo'llanma uchun kerak", "Ha"],
        ["Korxonadagi lavozimi", "Hozircha ixtiyoriy eslatma ustuni", "Yo'q"],
        ["Izoh", "Foydalanuvchi uchun eslatma", "Yo'q"],
    ]

    for row_index, row_data in enumerate(guide_rows, start=1):
        for column_index, value in enumerate(row_data, start=1):
            cell = guide_sheet.cell(row=row_index, column=column_index, value=value)
            cell.alignment = openpyxl.styles.Alignment(vertical="center", horizontal="left", wrap_text=True)
            cell.border = openpyxl.styles.Border(
                left=thin_side,
                right=thin_side,
                top=thin_side,
                bottom=thin_side,
            )
            if row_index == 1:
                cell.font = header_font
                cell.fill = header_fill
            else:
                cell.fill = data_fill

    guide_sheet.column_dimensions["A"].width = 28
    guide_sheet.column_dimensions["B"].width = 48
    guide_sheet.column_dimensions["C"].width = 14
    guide_sheet.freeze_panes = "A2"

    sheet.row_dimensions[1].height = 28
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(len(headers))}{len(sample_rows) + 1}"

    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)

    response = HttpResponse(
        buffer.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = 'attachment; filename="namuna_talabalar.xlsx"'
    return response


@login_required(login_url='login')
def download_template_source(request, filename):
    allowed_files = {item["key"]: item["filename"] for item in get_template_sources()}
    target_name = allowed_files.get(filename)
    if not target_name:
        return HttpResponse("Shablon topilmadi.", status=404)

    file_path = os.path.join(settings.BASE_DIR, "app_excel", "document_sources", target_name)
    if not os.path.exists(file_path):
        return HttpResponse("Shablon fayli topilmadi.", status=404)

    return FileResponse(open(file_path, "rb"), as_attachment=True, filename=target_name)


@login_required(login_url='login')
def upload_excel(request):
    profile = ensure_profile(request.user)
    if request.method == 'POST' and request.FILES.get('file'):
        uploaded_file = request.FILES['file']
        filename = uploaded_file.name.lower()
        course = None
        start_date = ""
        end_date = ""
        practice_type = ""

        Student.objects.all().delete()

        try:
            if filename.endswith('.xlsx'):
                wb = openpyxl.load_workbook(uploaded_file)
                sheet = wb.active
                header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), None)
                if not header_row:
                    raise ValueError("Excel faylda sarlavha qatori topilmadi.")

                normalized_headers = [normalize_excel_header(value) for value in header_row]
                for row in sheet.iter_rows(min_row=2, values_only=True):
                    row_map = {
                        normalized_headers[index]: row[index]
                        for index in range(min(len(normalized_headers), len(row)))
                        if normalized_headers[index]
                    }
                    full_name = get_excel_value(
                        row_map,
                        "Talabaning F.I.Sh.",
                        "F.I.Sh.",
                        "Talaba F.I.Sh.",
                    )
                    if not full_name:
                        continue
                    if course is None:
                        raw_course = get_excel_value(row_map, "Kurs")
                        course = int(raw_course) if raw_course else 4
                    if not start_date:
                        start_date = normalize_excel_date(row_map.get(normalize_excel_header("Boshlanish sanasi")))
                    if not end_date:
                        end_date = normalize_excel_date(row_map.get(normalize_excel_header("Tugash sanasi")))
                    if not practice_type:
                        practice_type = get_excel_value(row_map, "Amaliyot turi")
                    Student.objects.create(
                        full_name=full_name,
                        direction=get_excel_value(row_map, "Yo'nalish"),
                        faculty=get_excel_value(row_map, "Fakultet"),
                        group=get_excel_value(row_map, "Guruh"),
                        company=get_excel_value(row_map, "Korxona nomi", "Amaliyot o'tash joyi"),
                        company_address=get_excel_value(row_map, "Korxona manzili"),
                        company_director=get_excel_value(
                            row_map,
                            "Korxona rahbari F.I.Sh.",
                            "Korxonadagi amaliyot rahbari",
                        ),
                        practice_supervisor=get_excel_value(
                            row_map,
                            "Universitet mas'ul rahbari",
                            "Universitetdagi amaliyot rahbari",
                        ),
                        company_phone=get_excel_value(row_map, "Telefon raqami", "Korxona telefoni"),
                        faculty_dean=get_excel_value(row_map, "Fakultet dekani"),
                        department_head=get_excel_value(row_map, "Kafedra mudiri"),
                    )

                if course is None:
                    raise ValueError("Excel faylda kamida bitta talaba qatori bo'lishi kerak.")
                if not start_date or not end_date:
                    raise ValueError("Excel faylda 'Boshlanish sanasi' va 'Tugash sanasi' ustunlarini to'ldiring.")
                if not practice_type:
                    practice_type = get_practice_type(course)

            elif filename.endswith('.docx'):
                course = int(request.POST.get("course") or request.session.get("course") or 4)
                start_date = request.POST.get("start_date", "") or request.session.get("start_date", "")
                end_date = request.POST.get("end_date", "") or request.session.get("end_date", "")
                practice_type = get_practice_type(course)
                doc = Document(uploaded_file)
                table = doc.tables[0]
                for row in table.rows[1:]:
                    cells = row.cells
                    Student.objects.create(
                        full_name=cells[1].text.strip(),
                        direction="",
                        group=cells[2].text.strip(),
                        company=cells[3].text.strip(),
                        company_address=cells[4].text.strip(),
                        company_director=cells[5].text.strip(),
                        company_phone=cells[6].text.strip(),
                        practice_supervisor=cells[7].text.strip(),
                        faculty=cells[8].text.strip(),
                        faculty_dean="",
                        department_head="",
                    )
            else:
                return render(request, 'app_excel/upload.html', {
                    'error': 'Faqat .xlsx yoki .docx fayl yuklash mumkin.',
                    'profile': profile,
                    'template_sources': get_template_sources(),
                })

            request.session['uploaded'] = True
            request.session['course'] = course
            request.session['start_date'] = start_date
            request.session['end_date'] = end_date
            request.session['practice_type'] = practice_type or get_practice_type(course)
            return redirect('upload_excel')

        except Exception as e:
            return render(request, 'app_excel/upload.html', {
                'error': str(e),
                'profile': profile,
                'template_sources': get_template_sources(),
            })

    uploaded = request.session.pop('uploaded', False)
    return render(request, 'app_excel/upload.html', {
        'uploaded': uploaded,
        'profile': profile,
        'template_sources': get_template_sources(),
    })


@login_required(login_url='login')
def export_to_word(request):
    missing_templates = validate_runtime_templates()
    if missing_templates:
        missing_labels = ", ".join(missing_templates)
        return HttpResponse(f"Faol DOCX shablonlar topilmadi: {missing_labels}. Avval tizim shablonlarini joylang.", status=400)

    students = Student.objects.all()
    if not students.exists():
        return HttpResponse("Hali hech qanday talaba ma'lumotlari mavjud emas.")

    course = int(request.GET.get("course", 4))
    practice_type = get_practice_type(course)
    first_student = students.first()

    start_date_str = request.session.get("start_date", "2026-02-02")
    end_date_str = request.session.get("end_date", "2026-04-11")
    doc = generate_shartnoma_document(first_student, course, practice_type, start_date_str, end_date_str)
    return save_document_response(doc, f"{sanitize_filename(first_student.company)}_shartnoma.docx")




@login_required(login_url='login')
def export_all_documents_zip(request):
    missing_templates = validate_runtime_templates()
    if missing_templates:
        missing_labels = ", ".join(missing_templates)
        return HttpResponse(f"ZIP yaratish uchun faol DOCX shablonlar topilmadi: {missing_labels}.", status=400)

    students = Student.objects.all()
    if not students.exists():
        return HttpResponse("Talabalar ma'lumotlari topilmadi.")

    # Foydalanuvchi kiritgan session ma'lumotlari
    course = int(request.session.get("course", 4))
    start_date_str = request.session.get("start_date", "2025-02-17")
    end_date_str = request.session.get("end_date", "2025-04-26")

    # O'zbekcha oylar
    OY_NOMLARI = {
        "January": "yanvar", "February": "fevral", "March": "mart",
        "April": "aprel", "May": "may", "June": "iyun", "July": "iyul",
        "August": "avgust", "September": "sentabr", "October": "oktabr",
        "November": "noyabr", "December": "dekabr"
    }


    start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
    end_date = datetime.strptime(end_date_str, "%Y-%m-%d")
    practice_type = get_practice_type(course)


    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for index, student in enumerate(students, start=1):
            folder = f"Hujjatlar/{index:02}_{sanitize_filename(student.full_name)}"
            doc1 = generate_shartnoma_document(student, course, practice_type, start_date_str, end_date_str)
            io1 = io.BytesIO()
            doc1.save(io1)
            io1.seek(0)
            zip_file.writestr(f"{folder}/shartnoma.docx", io1.read())

            doc2 = generate_kundalik_document(student, course, practice_type, start_date_str, end_date_str)
            io2 = io.BytesIO()
            doc2.save(io2)
            io2.seek(0)
            zip_file.writestr(f"{folder}/kundalik.docx", io2.read())

            doc3 = generate_yollanma_document(student, course, practice_type, start_date_str, end_date_str)
            io3 = io.BytesIO()
            doc3.save(io3)
            io3.seek(0)
            zip_file.writestr(f"{folder}/yollanma.docx", io3.read())

    zip_buffer.seek(0)
    response = HttpResponse(zip_buffer, content_type='application/zip')
    response['Content-Disposition'] = 'attachment; filename=Hujjatlar.zip'
    return response



@login_required(login_url='login')
def generate_contract_for_company(request, company_name):
    missing_templates = validate_runtime_templates()
    if missing_templates:
        missing_labels = ", ".join(missing_templates)
        return HttpResponse(f"Faol DOCX shablonlar topilmadi: {missing_labels}.", status=400)

    students = Student.objects.filter(company=company_name)
    if not students.exists():
        return HttpResponse(" Bu korxona bo‘yicha talabalar topilmadi.")

    course = int(request.GET.get("course", 4))
    practice_type = get_practice_type(course)

    first_student = students.first()
    student_count = students.count()

    start_date_str = request.session.get("start_date", "2026-02-02")
    end_date_str = request.session.get("end_date", "2026-04-11")
    doc = generate_shartnoma_document(first_student, course, practice_type, start_date_str, end_date_str)
    return save_document_response(doc, f"{sanitize_filename(company_name)}_shartnoma.docx")



def login_view(request):
    if request.method == "POST":
        username = request.POST.get("username")
        password = request.POST.get("password")
        remember = request.POST.get("remember")

        user = authenticate(request, username=username, password=password)
        if user:
            login(request, user)
            ensure_profile(user)

            if remember:
                request.session.set_expiry(31536000)
            else:
                request.session.set_expiry(0)

            return redirect("dashboard")
        else:
            messages.error(request, " Login yoki parol noto‘g‘ri.")

    return render(request, "app_excel/login.html")


def register_view(request):
    if request.method == "POST":
        full_name = request.POST.get("full_name", "").strip()
        username = request.POST.get("username")
        email = request.POST.get("email", "").strip()
        password1 = request.POST.get("password1")
        password2 = request.POST.get("password2")

        if not full_name:
            messages.error(request, "F.I.Sh. kiritilishi kerak.")
        elif not email:
            messages.error(request, "Email manzilini kiriting.")
        elif password1 != password2:
            messages.error(request, "Parollar mos emas!")
        elif User.objects.filter(username=username).exists():
            messages.error(request, "Bu login allaqachon mavjud!")
        elif User.objects.filter(email=email).exists():
            messages.error(request, "Bu email allaqachon mavjud!")
        elif len(password1 or "") < 8:
            messages.error(request, "Parol kamida 8 ta belgidan iborat bo‘lishi kerak.")
        else:
            name_parts = full_name.split(maxsplit=1)
            first_name = name_parts[0]
            last_name = name_parts[1] if len(name_parts) > 1 else ""
            user = User.objects.create_user(
                username=username,
                email=email,
                password=password1,
                first_name=first_name,
                last_name=last_name,
            )
            profile = ensure_profile(user)
            profile.save()

            messages.success(request, "Muvaffaqiyatli ro‘yxatdan o‘tildi.")
            login(request, user)
            return redirect("dashboard")

    return render(request, "app_excel/register.html")


def logout_view(request):
    logout(request)
    return redirect("login")


def generate_documents_zip():
    missing_templates = validate_runtime_templates()
    if missing_templates:
        raise FileNotFoundError(f"Faol DOCX shablonlar topilmadi: {', '.join(missing_templates)}")

    students = Student.objects.all()
    course = int(getattr(settings, "DEFAULT_PRACTICE_COURSE", 4))
    practice_type = get_practice_type(course)
    start_date_str = "2026-02-02"
    end_date_str = "2026-04-11"

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for index, student in enumerate(students, start=1):
            folder = f"Hujjatlar/{index:02}_{sanitize_filename(student.full_name)}"
            generated_docs = [
                ("shartnoma", generate_shartnoma_document(student, course, practice_type, start_date_str, end_date_str)),
                ("kundalik", generate_kundalik_document(student, course, practice_type, start_date_str, end_date_str)),
                ("yollanma", generate_yollanma_document(student, course, practice_type, start_date_str, end_date_str)),
            ]
            for name, generated_doc in generated_docs:
                io_doc = io.BytesIO()
                generated_doc.save(io_doc)
                io_doc.seek(0)
                zip_file.writestr(f"{folder}/{name}.docx", io_doc.read())

    zip_buffer.seek(0)
    return zip_buffer


@login_required(login_url='login')
def export_to_word(request):
    students = Student.objects.all()
    if not students.exists():
        return HttpResponse("Hali hech qanday talaba ma'lumotlari mavjud emas.")

    course = int(request.GET.get("course", 4))
    practice_type = get_practice_type(course)
    first_student = students.first()
    start_date_str = request.session.get("start_date", "2026-02-02")
    end_date_str = request.session.get("end_date", "2026-04-11")

    doc = generate_shartnoma_document(first_student, course, practice_type, start_date_str, end_date_str)
    return save_document_response(doc, f"{sanitize_filename(first_student.company)}_shartnoma.docx")


@login_required(login_url='login')
def export_all_documents_zip(request):
    students = Student.objects.all()
    if not students.exists():
        return HttpResponse("Talabalar ma'lumotlari topilmadi.", status=400)

    course = int(request.session.get("course", 4))
    start_date_str = request.session.get("start_date", "2025-02-17")
    end_date_str = request.session.get("end_date", "2025-04-26")
    validation_issues = validate_students_for_documents(students, course, start_date_str, end_date_str)
    if validation_issues:
        issues_text = " ; ".join(validation_issues[:10])
        return HttpResponse(f"Ma'lumotlarni to'ldiring. ZIP yuklanmadi: {issues_text}", status=400)
    practice_type = get_practice_type(course)

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for index, student in enumerate(students, start=1):
            folder = f"Hujjatlar/{index:02}_{sanitize_filename(student.full_name)}"
            generated_docs = [
                ("shartnoma", generate_shartnoma_document(student, course, practice_type, start_date_str, end_date_str)),
                ("kundalik", generate_kundalik_document(student, course, practice_type, start_date_str, end_date_str)),
                ("yollanma", generate_yollanma_document(student, course, practice_type, start_date_str, end_date_str)),
            ]
            for filename, generated_doc in generated_docs:
                temp_buffer = io.BytesIO()
                generated_doc.save(temp_buffer)
                temp_buffer.seek(0)
                zip_file.writestr(f"{folder}/{filename}.docx", temp_buffer.read())

    zip_buffer.seek(0)
    response = HttpResponse(zip_buffer.read(), content_type='application/zip')
    response['Content-Disposition'] = 'attachment; filename="Hujjatlar.zip"'
    return response


@login_required(login_url='login')
def generate_contract_for_company(request, company_name):
    students = Student.objects.filter(company=company_name)
    if not students.exists():
        return HttpResponse("Bu korxona bo'yicha talabalar topilmadi.", status=404)

    course = int(request.GET.get("course", 4))
    practice_type = get_practice_type(course)
    first_student = students.first()
    start_date_str = request.session.get("start_date", "2026-02-02")
    end_date_str = request.session.get("end_date", "2026-04-11")

    doc = generate_shartnoma_document(first_student, course, practice_type, start_date_str, end_date_str)
    return save_document_response(doc, f"{sanitize_filename(company_name)}_shartnoma.docx")


def generate_documents_zip():
    students = Student.objects.all()
    course = int(getattr(settings, "DEFAULT_PRACTICE_COURSE", 4))
    practice_type = get_practice_type(course)
    start_date_str = "2026-02-02"
    end_date_str = "2026-04-11"

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for index, student in enumerate(students, start=1):
            folder = f"Hujjatlar/{index:02}_{sanitize_filename(student.full_name)}"
            generated_docs = [
                ("shartnoma", generate_shartnoma_document(student, course, practice_type, start_date_str, end_date_str)),
                ("kundalik", generate_kundalik_document(student, course, practice_type, start_date_str, end_date_str)),
                ("yollanma", generate_yollanma_document(student, course, practice_type, start_date_str, end_date_str)),
            ]
            for filename, generated_doc in generated_docs:
                temp_buffer = io.BytesIO()
                generated_doc.save(temp_buffer)
                temp_buffer.seek(0)
                zip_file.writestr(f"{folder}/{filename}.docx", temp_buffer.read())

    zip_buffer.seek(0)
    return zip_buffer


def generate_shartnoma_document(student, course, practice_type, start_date_str, end_date_str):
    template_path = get_available_runtime_template_path("shartnoma")
    if not template_path:
        return build_fallback_document("Shartnoma", student, course, practice_type, start_date_str, end_date_str)

    doc = Document(template_path)
    normalize_document_formatting(doc)
    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)
    company_count = get_company_student_count(student.company)

    doc.paragraphs[0].text = f"TALABALARNING {practice_type.upper()}NI TASHKIL ETISH BO'YICHA SHARTNOMA"
    doc.paragraphs[3].text = f"Toshkent sh.                                                                 {start_parts['year']}-yil «{start_parts['day']}» {start_parts['month']}"
    doc.paragraphs[5].text = (
        "Biz, quyida imzo chekuvchilar, Muhammad Al-Xorazmiy nomidagi Toshkent axborot texnologiyalari universiteti "
        "(keying o'rinlarda – Universitet)dan tayinlangan shaxs Dj.Sultanov - o'quv ishlari bo'yicha prorektor bir tomondan,"
    )
    doc.paragraphs[6].text = student.company
    doc.paragraphs[8].text = "(keying o'rinlarda – Korxona) dan tayinlangan shaxs"
    doc.paragraphs[9].text = student.company_director

    main_table = doc.tables[0]
    main_table.rows[2].cells[1].text = student.direction or student.faculty
    main_table.rows[2].cells[2].text = str(course)
    main_table.rows[2].cells[3].text = practice_type
    main_table.rows[2].cells[4].text = str(company_count)
    main_table.rows[2].cells[5].text = start_parts["short"]
    main_table.rows[2].cells[6].text = end_parts["short"]

    info_table = doc.tables[1]
    company_block = f"Korxona\n{student.company}\n{student.company_address}\nTel.: {student.company_phone}"
    info_table.rows[0].cells[1].text = company_block
    info_table.rows[4].cells[1].text = (
        "Korxonadan ajratilgan\n"
        f"amaliyot rahbari ________   {student.company_director}"
    )
    return doc


def generate_yollanma_document(student, course, practice_type, start_date_str, end_date_str):
    template_path = get_available_runtime_template_path("yollanma")
    if not template_path:
        return build_fallback_document("Yo'llanma", student, course, practice_type, start_date_str, end_date_str)

    doc = Document(template_path)
    normalize_document_formatting(doc)
    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)

    doc.paragraphs[4].text = (
        "Muxammad al-Xorazmiy nomidagi Toshkent axborot texnologiyalari universiteti O'zbekiston Respublikasi oliy va "
        "o'rta maxsus ta'lim vazirligi tomonidan tasdiqlangan “Oliy ta'lim muassasalari talabalarining malaka "
        f"amaliyotini o'tash tartibi to'g'risidagi nizom” va {student.company} hamda Muxammad al-Xorazmiy nomidagi "
        f"TATU o'rtasidagi shartnoma asosida {student.direction or student.faculty} ta'lim yo'nalishida tahsil olayotgan talaba"
    )
    doc.paragraphs[5].text = f"{student.full_name} {practice_type.lower()}ni o'tash uchun    {student.company}  ga yubormoqda."
    doc.paragraphs[7].text = f"Amaliyot muddati:     {start_parts['short']}-yildan,       {end_parts['short']} yilgacha"
    doc.paragraphs[11].text = (
        f"TATU dan ketdi  “{start_parts['day']}” {start_parts['month']} {start_parts['year']}-yil                              "
        f"Korxonaga keldi “{start_parts['day']}” {start_parts['month']} {start_parts['year']}-yil"
    )
    doc.paragraphs[12].text = f"Fakultet dekani  {student.faculty_dean}                                               Korxona rahbari  {student.company_director}"
    doc.paragraphs[17].text = (
        f"Korxonadan ketdi “{end_parts['day']}”  {end_parts['month']} {end_parts['year']}-yil                               "
        f"TATU ga keldi  “{end_parts['day']}” {end_parts['month']} {end_parts['year']}-yil"
    )
    doc.paragraphs[19].text = f"Korxona rahbari  {student.company_director}                                            Fakultet dekani  {student.faculty_dean}"
    doc.paragraphs[27].text = f"{student.full_name}    {student.company}   ga  amaliyotga keldi."
    doc.paragraphs[28].text = "Texnika xavfsizligi bo'yicha sinovni   “5”    bahoga topshirdi va amaliyotni o'tashga ruxsat berildi."
    doc.paragraphs[30].text = "Komissiya raisi  ____________    Amaliyot komissiyasi"
    doc.paragraphs[35].text = f"Talaba  {student.full_name}   ishga qo'yildi Amaliyotchi"
    doc.paragraphs[39].text = f"Korxonadan tayinlagan amaliyot rahbari     ____________    {student.company_director}"
    return doc


@login_required(login_url='login')
def profile_view(request):
    profile = ensure_profile(request.user)
    return render(request, 'app_excel/profile.html', {
        'user': request.user,
        'profile': profile
    })


@login_required(login_url='login')
def account_settings_view(request):
    profile = ensure_profile(request.user)
    if request.method == "POST":
        current_password = request.POST.get("current_password", "")
        new_password = request.POST.get("new_password", "")
        confirm_password = request.POST.get("confirm_password", "")

        if not request.user.check_password(current_password):
            messages.error(request, "Joriy parol noto‘g‘ri.")
        elif len(new_password) < 8:
            messages.error(request, "Yangi parol kamida 8 ta belgidan iborat bo‘lishi kerak.")
        elif new_password != confirm_password:
            messages.error(request, "Yangi parollar bir-biriga mos emas.")
        elif current_password == new_password:
            messages.error(request, "Yangi parol joriy paroldan farq qilishi kerak.")
        else:
            request.user.set_password(new_password)
            request.user.save()
            update_session_auth_hash(request, request.user)
            messages.success(request, "Parol muvaffaqiyatli yangilandi.")
            return redirect("account_settings")

    return render(request, 'app_excel/account_settings.html', {
        'user': request.user,
        'profile': profile,
    })



@login_required(login_url='login')
def account_settings_view(request):
    profile = ensure_profile(request.user)
    if request.method == "POST":
        action = request.POST.get("action", "password")

        if action == "profile":
            username = request.POST.get("username", "").strip()
            first_name = request.POST.get("first_name", "").strip()
            last_name = request.POST.get("last_name", "").strip()
            email = request.POST.get("email", "").strip()

            if not username:
                messages.error(request, "Foydalanuvchi nomi bo'sh bo'lishi mumkin emas.")
            elif User.objects.exclude(pk=request.user.pk).filter(username=username).exists():
                messages.error(request, "Bu foydalanuvchi nomi allaqachon band.")
            elif email and User.objects.exclude(pk=request.user.pk).filter(email__iexact=email).exists():
                messages.error(request, "Bu email allaqachon boshqa akkauntga ulangan.")
            else:
                request.user.username = username
                request.user.first_name = first_name
                request.user.last_name = last_name
                request.user.email = email
                request.user.save(update_fields=["username", "first_name", "last_name", "email"])
                messages.success(request, "Profil ma'lumotlari muvaffaqiyatli yangilandi.")
                return redirect("account_settings")
        else:
            current_password = request.POST.get("current_password", "")
            new_password = request.POST.get("new_password", "")
            confirm_password = request.POST.get("confirm_password", "")

            if not request.user.check_password(current_password):
                messages.error(request, "Joriy parol noto'g'ri.")
            elif len(new_password) < 8:
                messages.error(request, "Yangi parol kamida 8 ta belgidan iborat bo'lishi kerak.")
            elif new_password != confirm_password:
                messages.error(request, "Yangi parollar bir-biriga mos emas.")
            elif current_password == new_password:
                messages.error(request, "Yangi parol joriy paroldan farq qilishi kerak.")
            else:
                request.user.set_password(new_password)
                request.user.save()
                update_session_auth_hash(request, request.user)
                messages.success(request, "Parol muvaffaqiyatli yangilandi.")
                return redirect("account_settings")

    return render(request, 'app_excel/account_settings.html', {
        'user': request.user,
        'profile': profile,
    })


@api_view(['POST'])
def custom_login_api(request):
    username = request.data.get("username")
    password = request.data.get("password")

    user = authenticate(username=username, password=password)
    if user:
        token, created = Token.objects.get_or_create(user=user)
        return Response({"token": token.key, "username": user.username})
    return Response({"error": "Login yoki parol noto‘g‘ri"}, status=400)



def handler403(request, exception=None):
    return render(request, "app_excel/errors/403.html", status=403)

def handler404(request, exception=None):
    return render(request, "app_excel/errors/404.html", status=404)

def handler500(request):
    return render(request, "app_excel/errors/500.html", status=500)

def csrf_failure(request, reason=""):
    return render(request, 'app_excel/403_csrf.html', status=403)


def generate_kundalik_from_fixed_template(student, course, practice_type, start_date_str, end_date_str):
    doc = build_manual_template_document("kundalik")
    if doc is None:
        return None

    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)
    academic_year = f"{start_parts['year']}/{int(start_parts['year']) + 1}"
    direction_label = student.direction or student.faculty

    replacements = {
        5: f"{direction_label} ta'lim yo'nalishi {course} - bosqich talabasi {student.full_name} ning {academic_year} o'quv yilidagi {practice_type.lower()}",
        10: f"1.1. Amaliyot joyi va muddati {student.company}, {student.company_address} muddati: {start_parts['short']} dan   {end_parts['short']} gacha",
        12: "Universitetdan",
        14: "Korxonadan",
        16: f"1.3. Talaba {student.full_name} ga \"{student.faculty}\" kafedrasidan berilgan individual topshiriqlar {practice_type.lower()} bo'yicha kundalik yuritish va hisobot tayyorlash.",
        17: f"1.4. Amaliyotga keldi: {format_document_date(start_date_str)},  ketdi: {format_document_date(end_date_str)}",
        56: "Korxonadan rahbar",
        59: f"M.O'. \t\t{end_parts['year']}-yil       \"_____\"  ____________________",
        63: "Universitetdan amaliyot rahbari",
        65: "Kafedra mudiri",
        67: f"{end_parts['year']}-yil \"____\" ___________",
    }
    for index, text in replacements.items():
        if index < len(doc.paragraphs):
            doc.paragraphs[index].text = text

    if course == 3 and 25 < len(doc.paragraphs):
        doc.paragraphs[25].text = "2.6. Amaliyot muddati 3-bosqichda o'quv reja asosida belgilanadi."

    if 39 < len(doc.paragraphs):
        set_paragraph_segments(
            doc.paragraphs[39],
            [
                ("Universitetdan ", False),
                ("____________ ", False),
                (student.practice_supervisor, True),
                (" ", False),
                (format_document_date(start_date_str), True),
            ],
        )
    if 40 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[40], [("Imzo", False, True), ("\t\t         ", False), ("F.I.Sh.", False, True)])
    if 42 < len(doc.paragraphs):
        set_paragraph_segments(
            doc.paragraphs[42],
            [
                ("Korxonadan ", False),
                ("____________ ", False),
                (student.company_director, True),
                (" ", False),
                (format_document_date(end_date_str), True),
            ],
        )
    if 43 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[43], [("Imzo", False, True), ("\t\t         ", False), ("F.I.Sh.", False, True)])
    if 56 < len(doc.paragraphs):
        set_paragraph_segments(
            doc.paragraphs[56],
            [
                ("Korxonadan rahbar          ", False),
                ("________________ ", False),
                (student.company_director, True),
            ],
        )
    if 57 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[57], [("imzo", False, True), ("                                                              ", False), ("F.I.Sh.", False, True)])
    if 63 < len(doc.paragraphs):
        set_paragraph_segments(
            doc.paragraphs[63],
            [
                ("Universitetdan amaliyot rahbari      ", False),
                ("_____________ ", False),
                (student.practice_supervisor, True),
            ],
        )
    if 64 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[64], [("imzo", False, True), ("                                                              ", False), ("F.I.Sh.", False, True)])
    if 65 < len(doc.paragraphs):
        set_paragraph_segments(
            doc.paragraphs[65],
            [
                ("Kafedra mudiri                                ", False),
                ("_____________ ", False),
                (student.department_head, True),
            ],
        )
    if 66 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[66], [("imzo", False, True), ("                                                              ", False), ("F.I.Sh.", False, True)])

    normalize_document_formatting(doc)
    return doc



def generate_yollanma_document(student, course, practice_type, start_date_str, end_date_str):
    doc = build_manual_template_document("yollanma")
    if doc is None:
        return build_fallback_document("Yo'llanma", student, course, practice_type, start_date_str, end_date_str)

    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)
    direction_label = student.direction or student.faculty

    if 4 < len(doc.paragraphs):
        doc.paragraphs[4].text = (
            "Muxammad al-Xorazmiy nomidagi Toshkent axborot texnologiyalari universiteti O'zbekiston Respublikasi oliy va o'rta maxsus ta'lim vazirligi tomonidan tasdiqlangan "
            "\"Oliy ta'lim muassasalari talabalarining malaka amaliyotini o'tash tartibi to'g'risidagi nizom\" va"
        )
    if 5 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[5], [(student.company, True)])
        center_paragraph(doc.paragraphs[5], font_size_pt=11)
    if 6 < len(doc.paragraphs):
        doc.paragraphs[6].text = "hamda Muxammad al-Xorazmiy nomidagi TATU o'rtasidagi shartnoma asosida"
    if 7 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[7], [(direction_label, True)])
        center_paragraph(doc.paragraphs[7], font_size_pt=11)
    if 8 < len(doc.paragraphs):
        doc.paragraphs[8].text = ""
    if 9 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[9], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[9], font_size_pt=11)
    if 10 < len(doc.paragraphs):
        doc.paragraphs[10].text = ""
    if 12 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[12], [(student.company, True)])
        center_paragraph(doc.paragraphs[12], font_size_pt=11)
    if 13 < len(doc.paragraphs):
        doc.paragraphs[13].text = ""
    if 15 < len(doc.paragraphs):
        doc.paragraphs[15].text = f"Amaliyot muddati:     {start_parts['short']}-yildan,       {end_parts['short']} yilgacha"
    if 19 < len(doc.paragraphs):
        doc.paragraphs[19].text = f"TATU dan ketdi: \"{start_parts['day']}\" {start_parts['month']} {start_parts['year']}-yil                              Korxonaga keldi: \"{start_parts['day']}\" {start_parts['month']} {start_parts['year']}-yil"
    if 20 < len(doc.paragraphs):
        doc.paragraphs[20].text = f"Fakultet dekani: _____________________ {student.faculty_dean}                                 Korxona rahbari: _____________________ {student.company_director}"
    if 21 < len(doc.paragraphs):
        doc.paragraphs[21].text = ""
    if 26 < len(doc.paragraphs):
        doc.paragraphs[26].text = f"Korxonadan ketdi: \"{end_parts['day']}\" {end_parts['month']} {end_parts['year']}-yil                               TATU ga keldi: \"{end_parts['day']}\" {end_parts['month']} {end_parts['year']}-yil"
    if 28 < len(doc.paragraphs):
        doc.paragraphs[28].text = f"Korxona rahbari: _____________________ {student.company_director}                                 Fakultet dekani: _____________________ {student.faculty_dean}"
    if 29 < len(doc.paragraphs):
        doc.paragraphs[29].text = ""
    if 37 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[37], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[37], font_size_pt=11)
    if 38 < len(doc.paragraphs):
        doc.paragraphs[38].text = ""
    if 39 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[39], [(student.company, True)])
        center_paragraph(doc.paragraphs[39], font_size_pt=11)
    if 40 < len(doc.paragraphs):
        doc.paragraphs[40].text = ""
    if 41 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[41], [("Texnika xavfsizligi bo'yicha sinovni \"", False), ("5", True), ("\" bahoga topshirdi va amaliyotni o'tashga ruxsat berildi.", False)])
    if 44 < len(doc.paragraphs):
        doc.paragraphs[44].text = f"Komissiya raisi       ____________    {student.practice_supervisor}"
    if 45 < len(doc.paragraphs):
        doc.paragraphs[45].text = ""
    if 49 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[49], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[49], font_size_pt=11)
    if 50 < len(doc.paragraphs):
        doc.paragraphs[50].text = ""
    if 52 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[52], [("Amaliyotchi", True)])
        center_paragraph(doc.paragraphs[52], font_size_pt=11)
    if 53 < len(doc.paragraphs):
        doc.paragraphs[53].text = ""
    if 59 < len(doc.paragraphs):
        doc.paragraphs[59].text = f"Korxonadan tayinlagan amaliyot rahbari     ____________    {student.company_director}"

    normalize_document_formatting(doc)
    return doc


def generate_yollanma_document(student, course, practice_type, start_date_str, end_date_str):
    doc = build_manual_template_document("yollanma")
    if doc is None:
        return build_fallback_document("Yo'llanma", student, course, practice_type, start_date_str, end_date_str)

    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)
    direction_label = student.direction or student.faculty

    if 4 < len(doc.paragraphs):
        doc.paragraphs[4].text = (
            "Muxammad al-Xorazmiy nomidagi Toshkent axborot texnologiyalari universiteti O'zbekiston Respublikasi oliy va o'rta maxsus ta'lim vazirligi tomonidan tasdiqlangan "
            "\"Oliy ta'lim muassasalari talabalarining malaka amaliyotini o'tash tartibi to'g'risidagi nizom\" va"
        )
    if 5 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[5], [(student.company, True)])
        center_paragraph(doc.paragraphs[5], font_size_pt=11)
    if 6 < len(doc.paragraphs):
        doc.paragraphs[6].text = "hamda Muxammad al-Xorazmiy nomidagi TATU o'rtasidagi shartnoma asosida"
    if 7 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[7], [(direction_label, True)])
        center_paragraph(doc.paragraphs[7], font_size_pt=11)
    if 8 < len(doc.paragraphs):
        doc.paragraphs[8].text = ""
    if 9 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[9], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[9], font_size_pt=11)
    if 10 < len(doc.paragraphs):
        doc.paragraphs[10].text = ""
    if 12 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[12], [(student.company, True)])
        center_paragraph(doc.paragraphs[12], font_size_pt=11)
    if 13 < len(doc.paragraphs):
        doc.paragraphs[13].text = ""
    if 15 < len(doc.paragraphs):
        doc.paragraphs[15].text = f"Amaliyot muddati:     {start_parts['short']}-yildan,       {end_parts['short']} yilgacha"
    if 19 < len(doc.paragraphs):
        doc.paragraphs[19].text = f"TATU dan ketdi: \"{start_parts['day']}\" {start_parts['month']} {start_parts['year']}-yil                              Korxonaga keldi: \"{start_parts['day']}\" {start_parts['month']} {start_parts['year']}-yil"
    if 20 < len(doc.paragraphs):
        doc.paragraphs[20].text = f"Fakultet dekani: _____________________ {student.faculty_dean}                                 Korxona rahbari: _____________________ {student.company_director}"
    if 21 < len(doc.paragraphs):
        doc.paragraphs[21].text = ""
    if 26 < len(doc.paragraphs):
        doc.paragraphs[26].text = f"Korxonadan ketdi: \"{end_parts['day']}\" {end_parts['month']} {end_parts['year']}-yil                               TATU ga keldi: \"{end_parts['day']}\" {end_parts['month']} {end_parts['year']}-yil"
    if 28 < len(doc.paragraphs):
        doc.paragraphs[28].text = f"Korxona rahbari: _____________________ {student.company_director}                                 Fakultet dekani: _____________________ {student.faculty_dean}"
    if 29 < len(doc.paragraphs):
        doc.paragraphs[29].text = ""
    if 37 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[37], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[37], font_size_pt=11)
    if 38 < len(doc.paragraphs):
        doc.paragraphs[38].text = ""
    if 39 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[39], [(student.company, True)])
        center_paragraph(doc.paragraphs[39], font_size_pt=11)
    if 40 < len(doc.paragraphs):
        doc.paragraphs[40].text = ""
    if 41 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[41], [("Texnika xavfsizligi bo'yicha sinovni \"", False), ("5", True), ("\" bahoga topshirdi va amaliyotni o'tashga ruxsat berildi.", False)])
    if 44 < len(doc.paragraphs):
        doc.paragraphs[44].text = f"Komissiya raisi       ____________    {student.practice_supervisor}"
    if 45 < len(doc.paragraphs):
        doc.paragraphs[45].text = ""
    if 49 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[49], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[49], font_size_pt=11)
    if 50 < len(doc.paragraphs):
        doc.paragraphs[50].text = ""
    if 52 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[52], [("Amaliyotchi", True)])
        center_paragraph(doc.paragraphs[52], font_size_pt=11)
    if 53 < len(doc.paragraphs):
        doc.paragraphs[53].text = ""
    if 59 < len(doc.paragraphs):
        doc.paragraphs[59].text = f"Korxonadan tayinlagan amaliyot rahbari     ____________    {student.company_director}"

    normalize_document_formatting(doc)
    return doc

def generate_yollanma_document(student, course, practice_type, start_date_str, end_date_str):
    doc = build_manual_template_document("yollanma")
    if doc is None:
        return build_fallback_document("Yo'llanma", student, course, practice_type, start_date_str, end_date_str)

    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)
    direction_label = student.direction or student.faculty

    if 4 < len(doc.paragraphs):
        doc.paragraphs[4].text = (
            "Muxammad al-Xorazmiy nomidagi Toshkent axborot texnologiyalari universiteti O'zbekiston Respublikasi oliy va o'rta maxsus ta'lim vazirligi tomonidan tasdiqlangan "
            "\"Oliy ta'lim muassasalari talabalarining malaka amaliyotini o'tash tartibi to'g'risidagi nizom\" va"
        )
    if 5 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[5], [(student.company, True)])
        center_paragraph(doc.paragraphs[5], font_size_pt=11)
    if 6 < len(doc.paragraphs):
        doc.paragraphs[6].text = "hamda Muxammad al-Xorazmiy nomidagi TATU o'rtasidagi shartnoma asosida"
    if 7 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[7], [(direction_label, True)])
        center_paragraph(doc.paragraphs[7], font_size_pt=11)
    if 8 < len(doc.paragraphs):
        doc.paragraphs[8].text = ""
    if 9 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[9], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[9], font_size_pt=11)
    if 10 < len(doc.paragraphs):
        doc.paragraphs[10].text = ""
    if 12 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[12], [(student.company, True)])
        center_paragraph(doc.paragraphs[12], font_size_pt=11)
    if 13 < len(doc.paragraphs):
        doc.paragraphs[13].text = ""
    if 15 < len(doc.paragraphs):
        doc.paragraphs[15].text = f"Amaliyot muddati:     {start_parts['short']}-yildan,       {end_parts['short']} yilgacha"
    if 19 < len(doc.paragraphs):
        doc.paragraphs[19].text = f"TATU dan ketdi: \"{start_parts['day']}\" {start_parts['month']} {start_parts['year']}-yil                              Korxonaga keldi: \"{start_parts['day']}\" {start_parts['month']} {start_parts['year']}-yil"
    if 20 < len(doc.paragraphs):
        doc.paragraphs[20].text = f"Fakultet dekani: _____________________ {student.faculty_dean}                                 Korxona rahbari: _____________________ {student.company_director}"
    if 21 < len(doc.paragraphs):
        doc.paragraphs[21].text = ""
    if 26 < len(doc.paragraphs):
        doc.paragraphs[26].text = f"Korxonadan ketdi: \"{end_parts['day']}\" {end_parts['month']} {end_parts['year']}-yil                               TATU ga keldi: \"{end_parts['day']}\" {end_parts['month']} {end_parts['year']}-yil"
    if 28 < len(doc.paragraphs):
        doc.paragraphs[28].text = f"Korxona rahbari: _____________________ {student.company_director}                                 Fakultet dekani: _____________________ {student.faculty_dean}"
    if 29 < len(doc.paragraphs):
        doc.paragraphs[29].text = ""
    if 37 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[37], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[37], font_size_pt=11)
    if 38 < len(doc.paragraphs):
        doc.paragraphs[38].text = ""
    if 39 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[39], [(student.company, True)])
        center_paragraph(doc.paragraphs[39], font_size_pt=11)
    if 40 < len(doc.paragraphs):
        doc.paragraphs[40].text = ""
    if 41 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[41], [("Texnika xavfsizligi bo'yicha sinovni \"", False), ("5", True), ("\" bahoga topshirdi va amaliyotni o'tashga ruxsat berildi.", False)])
    if 44 < len(doc.paragraphs):
        doc.paragraphs[44].text = f"Komissiya raisi       ____________    {student.practice_supervisor}"
    if 45 < len(doc.paragraphs):
        doc.paragraphs[45].text = ""
    if 49 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[49], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[49], font_size_pt=11)
    if 50 < len(doc.paragraphs):
        doc.paragraphs[50].text = ""
    if 52 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[52], [("Amaliyotchi", True)])
        center_paragraph(doc.paragraphs[52], font_size_pt=11)
    if 53 < len(doc.paragraphs):
        doc.paragraphs[53].text = ""
    if 59 < len(doc.paragraphs):
        doc.paragraphs[59].text = f"Korxonadan tayinlagan amaliyot rahbari     ____________    {student.company_director}"

    normalize_document_formatting(doc)
    return doc


def generate_kundalik_document(student, course, practice_type, start_date_str, end_date_str):
    doc = build_manual_template_document("kundalik")
    if doc is None:
        return build_fallback_document("Kundalik", student, course, practice_type, start_date_str, end_date_str)

    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)
    academic_year = f"{start_parts['year']}/{int(start_parts['year']) + 1}"
    direction_label = student.direction or student.faculty

    if 5 < len(doc.paragraphs):
        doc.paragraphs[5].text = (
            f"{direction_label} ta'lim yo'nalishi {course} - bosqich talabasi "
            f"{student.full_name} ning {academic_year} o'quv yilidagi {practice_type.lower()}"
        )
    if 10 < len(doc.paragraphs):
        doc.paragraphs[10].text = (
            f"1.1. Amaliyot joyi va muddati {student.company}, {student.company_address} "
            f"muddati: {start_parts['short']} dan {end_parts['short']} gacha"
        )
    if 12 < len(doc.paragraphs):
        set_paragraph_segments(
            doc.paragraphs[12],
            [("Universitetdan ", False), (student.practice_supervisor, True)],
        )
    if 13 < len(doc.paragraphs):
        doc.paragraphs[13].text = ""
    if 14 < len(doc.paragraphs):
        set_paragraph_segments(
            doc.paragraphs[14],
            [("Korxonadan ", False), (student.company_director, True)],
        )
    if 15 < len(doc.paragraphs):
        doc.paragraphs[15].text = ""
    if 16 < len(doc.paragraphs):
        doc.paragraphs[16].text = (
            f"1.3. Talaba {student.full_name} ga \"{student.faculty}\" kafedrasidan berilgan "
            f"individual topshiriqlar {practice_type.lower()} bo'yicha kundalik yuritish va hisobot tayyorlash."
        )
    if 17 < len(doc.paragraphs):
        doc.paragraphs[17].text = (
            f"1.4. Amaliyotga keldi: {format_document_date(start_date_str)}, "
            f"ketdi: {format_document_date(end_date_str)}"
        )
    if course == 3 and 25 < len(doc.paragraphs):
        doc.paragraphs[25].text = "2.6. Amaliyot muddati 3-bosqichda o'quv reja asosida belgilanadi."
    if 39 < len(doc.paragraphs):
        set_paragraph_segments(
            doc.paragraphs[39],
            [
                ("Universitetdan ", False),
                ("__________ ", False),
                (student.practice_supervisor, True),
                (" ", False),
                (format_document_date(start_date_str), True),
            ],
        )
    if 40 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[40], [("Imzo", False, True), ("\t\t", False), ("F.I.Sh.", False, True)])
    if 42 < len(doc.paragraphs):
        set_paragraph_segments(
            doc.paragraphs[42],
            [
                ("Korxonadan ", False),
                ("__________ ", False),
                (student.company_director, True),
                (" ", False),
                (format_document_date(end_date_str), True),
            ],
        )
    if 43 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[43], [("Imzo", False, True), ("\t\t", False), ("F.I.Sh.", False, True)])
    if 56 < len(doc.paragraphs):
        set_paragraph_segments(
            doc.paragraphs[56],
            [("Korxonadan rahbar          ", False), ("____________ ", False), (student.company_director, True)],
        )
    if 57 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[57], [("imzo", False, True), ("                                                              ", False), ("F.I.Sh.", False, True)])
    if 63 < len(doc.paragraphs):
        set_paragraph_segments(
            doc.paragraphs[63],
            [("Universitetdan amaliyot rahbari      ", False), ("____________ ", False), (student.practice_supervisor, True)],
        )
    if 64 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[64], [("imzo", False, True), ("                                                              ", False), ("F.I.Sh.", False, True)])
    if 65 < len(doc.paragraphs):
        set_paragraph_segments(
            doc.paragraphs[65],
            [("Kafedra mudiri                                ", False), ("____________ ", False), (student.department_head, True)],
        )
    if 66 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[66], [("imzo", False, True), ("                                                              ", False), ("F.I.Sh.", False, True)])

    normalize_document_formatting(doc)
    return doc


def generate_yollanma_document(student, course, practice_type, start_date_str, end_date_str):
    doc = build_manual_template_document("yollanma")
    if doc is None:
        return build_fallback_document("Yo'llanma", student, course, practice_type, start_date_str, end_date_str)

    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)
    direction_label = student.direction or student.faculty

    if 4 < len(doc.paragraphs):
        doc.paragraphs[4].text = (
            "Muxammad al-Xorazmiy nomidagi Toshkent axborot texnologiyalari universiteti O'zbekiston Respublikasi oliy va o'rta maxsus ta'lim vazirligi tomonidan tasdiqlangan "
            "\"Oliy ta'lim muassasalari talabalarining malaka amaliyotini o'tash tartibi to'g'risidagi nizom\" va"
        )
    if 5 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[5], [(student.company, True)])
        center_paragraph(doc.paragraphs[5], font_size_pt=11)
    if 6 < len(doc.paragraphs):
        doc.paragraphs[6].text = "hamda Muxammad al-Xorazmiy nomidagi TATU o'rtasidagi shartnoma asosida"
    if 7 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[7], [(direction_label, True)])
        center_paragraph(doc.paragraphs[7], font_size_pt=11)
    if 8 < len(doc.paragraphs):
        doc.paragraphs[8].text = ""
    if 9 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[9], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[9], font_size_pt=11)
    if 10 < len(doc.paragraphs):
        doc.paragraphs[10].text = ""
    if 12 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[12], [(student.company, True)])
        center_paragraph(doc.paragraphs[12], font_size_pt=11)
    if 13 < len(doc.paragraphs):
        doc.paragraphs[13].text = ""
    if 15 < len(doc.paragraphs):
        doc.paragraphs[15].text = f"Amaliyot muddati:     {start_parts['short']}-yildan,       {end_parts['short']} yilgacha"
    if 19 < len(doc.paragraphs):
        doc.paragraphs[19].text = f"TATU dan ketdi: \"{start_parts['day']}\" {start_parts['month']} {start_parts['year']}-yil                              Korxonaga keldi: \"{start_parts['day']}\" {start_parts['month']} {start_parts['year']}-yil"
    if 20 < len(doc.paragraphs):
        doc.paragraphs[20].text = f"Fakultet dekani: _____________________ {student.faculty_dean}                                 Korxona rahbari: _____________________ {student.company_director}"
    if 21 < len(doc.paragraphs):
        doc.paragraphs[21].text = ""
    if 26 < len(doc.paragraphs):
        doc.paragraphs[26].text = f"Korxonadan ketdi: \"{end_parts['day']}\" {end_parts['month']} {end_parts['year']}-yil                               TATU ga keldi: \"{end_parts['day']}\" {end_parts['month']} {end_parts['year']}-yil"
    if 28 < len(doc.paragraphs):
        doc.paragraphs[28].text = f"Korxona rahbari: _____________________ {student.company_director}                                 Fakultet dekani: _____________________ {student.faculty_dean}"
    if 29 < len(doc.paragraphs):
        doc.paragraphs[29].text = ""
    if 37 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[37], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[37], font_size_pt=11)
    if 38 < len(doc.paragraphs):
        doc.paragraphs[38].text = ""
    if 39 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[39], [(student.company, True)])
        center_paragraph(doc.paragraphs[39], font_size_pt=11)
    if 40 < len(doc.paragraphs):
        doc.paragraphs[40].text = ""
    if 41 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[41], [("Texnika xavfsizligi bo'yicha sinovni \"", False), ("5", True), ("\" bahoga topshirdi va amaliyotni o'tashga ruxsat berildi.", False)])
    if 44 < len(doc.paragraphs):
        doc.paragraphs[44].text = f"Komissiya raisi       ____________    {student.practice_supervisor}"
    if 45 < len(doc.paragraphs):
        doc.paragraphs[45].text = ""
    if 49 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[49], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[49], font_size_pt=11)
    if 50 < len(doc.paragraphs):
        doc.paragraphs[50].text = ""
    if 52 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[52], [("Amaliyotchi", True)])
        center_paragraph(doc.paragraphs[52], font_size_pt=11)
    if 53 < len(doc.paragraphs):
        doc.paragraphs[53].text = ""
    if 59 < len(doc.paragraphs):
        doc.paragraphs[59].text = f"Korxonadan tayinlagan amaliyot rahbari     ____________    {student.company_director}"

    normalize_document_formatting(doc)
    return doc


def generate_yollanma_document(student, course, practice_type, start_date_str, end_date_str):
    doc = build_manual_template_document("yollanma")
    if doc is None:
        return build_fallback_document("Yo'llanma", student, course, practice_type, start_date_str, end_date_str)

    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)
    direction_label = student.direction or student.faculty

    if 4 < len(doc.paragraphs):
        doc.paragraphs[4].text = (
            "Muxammad al-Xorazmiy nomidagi Toshkent axborot texnologiyalari universiteti O'zbekiston Respublikasi oliy va o'rta maxsus ta'lim vazirligi tomonidan tasdiqlangan "
            "\"Oliy ta'lim muassasalari talabalarining malaka amaliyotini o'tash tartibi to'g'risidagi nizom\" va"
        )
    if 5 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[5], [(student.company, True)])
        center_paragraph(doc.paragraphs[5], font_size_pt=11)
    if 6 < len(doc.paragraphs):
        doc.paragraphs[6].text = "hamda Muxammad al-Xorazmiy nomidagi TATU o'rtasidagi shartnoma asosida"
    if 7 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[7], [(direction_label, True)])
        center_paragraph(doc.paragraphs[7], font_size_pt=11)
    if 8 < len(doc.paragraphs):
        doc.paragraphs[8].text = ""
    if 9 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[9], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[9], font_size_pt=11)
    if 10 < len(doc.paragraphs):
        doc.paragraphs[10].text = ""
    if 12 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[12], [(student.company, True)])
        center_paragraph(doc.paragraphs[12], font_size_pt=11)
    if 13 < len(doc.paragraphs):
        doc.paragraphs[13].text = ""
    if 15 < len(doc.paragraphs):
        doc.paragraphs[15].text = f"Amaliyot muddati:     {start_parts['short']}-yildan,       {end_parts['short']} yilgacha"
    if 19 < len(doc.paragraphs):
        doc.paragraphs[19].text = f"TATU dan ketdi: \"{start_parts['day']}\" {start_parts['month']} {start_parts['year']}-yil                              Korxonaga keldi: \"{start_parts['day']}\" {start_parts['month']} {start_parts['year']}-yil"
    if 20 < len(doc.paragraphs):
        doc.paragraphs[20].text = f"Fakultet dekani: _____________________ {student.faculty_dean}                                 Korxona rahbari: _____________________ {student.company_director}"
    if 21 < len(doc.paragraphs):
        doc.paragraphs[21].text = ""
    if 26 < len(doc.paragraphs):
        doc.paragraphs[26].text = f"Korxonadan ketdi: \"{end_parts['day']}\" {end_parts['month']} {end_parts['year']}-yil                               TATU ga keldi: \"{end_parts['day']}\" {end_parts['month']} {end_parts['year']}-yil"
    if 28 < len(doc.paragraphs):
        doc.paragraphs[28].text = f"Korxona rahbari: _____________________ {student.company_director}                                 Fakultet dekani: _____________________ {student.faculty_dean}"
    if 29 < len(doc.paragraphs):
        doc.paragraphs[29].text = ""
    if 37 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[37], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[37], font_size_pt=11)
    if 38 < len(doc.paragraphs):
        doc.paragraphs[38].text = ""
    if 39 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[39], [(student.company, True)])
        center_paragraph(doc.paragraphs[39], font_size_pt=11)
    if 40 < len(doc.paragraphs):
        doc.paragraphs[40].text = ""
    if 41 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[41], [("Texnika xavfsizligi bo'yicha sinovni \"", False), ("5", True), ("\" bahoga topshirdi va amaliyotni o'tashga ruxsat berildi.", False)])
    if 44 < len(doc.paragraphs):
        doc.paragraphs[44].text = f"Komissiya raisi       ____________    {student.practice_supervisor}"
    if 45 < len(doc.paragraphs):
        doc.paragraphs[45].text = ""
    if 49 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[49], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[49], font_size_pt=11)
    if 50 < len(doc.paragraphs):
        doc.paragraphs[50].text = ""
    if 52 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[52], [("Amaliyotchi", True)])
        center_paragraph(doc.paragraphs[52], font_size_pt=11)
    if 53 < len(doc.paragraphs):
        doc.paragraphs[53].text = ""
    if 59 < len(doc.paragraphs):
        doc.paragraphs[59].text = f"Korxonadan tayinlagan amaliyot rahbari     ____________    {student.company_director}"

    normalize_document_formatting(doc)
    return doc


def generate_shartnoma_document(student, course, practice_type, start_date_str, end_date_str):
    doc = build_manual_template_document("shartnoma")
    if doc is None:
        return build_fallback_document("Shartnoma", student, course, practice_type, start_date_str, end_date_str)

    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)
    company_count = get_company_student_count(student.company)

    if len(doc.paragraphs) > 9:
        doc.paragraphs[0].text = f"TALABALARNING {practice_type.upper()}NI TASHKIL ETISH BO'YICHA SHARTNOMA"
        doc.paragraphs[3].text = f"Toshkent sh.                                                                 {start_parts['year']}-yil \"{start_parts['day']}\" {start_parts['month']}"
        doc.paragraphs[6].text = student.company
        doc.paragraphs[9].text = student.company_director

    if len(doc.tables) >= 2:
        main_table = doc.tables[0]
        if len(main_table.rows) > 2 and len(main_table.rows[2].cells) >= 7:
            row = main_table.rows[2].cells
            row[1].text = student.direction or student.faculty
            row[2].text = str(course)
            row[3].text = practice_type
            row[4].text = str(company_count)
            row[5].text = start_parts["short"]
            row[6].text = end_parts["short"]
            for cell, size in zip(row, (10, 9, 10, 9, 10, 9, 9)):
                style_cell(cell, horizontal=WD_ALIGN_PARAGRAPH.CENTER, font_size_pt=size)

        info_table = doc.tables[1]
        if len(info_table.rows) > 4 and len(info_table.rows[0].cells) >= 3:
            contact_block = f"{student.company_address}\nTel.: {student.company_phone}"
            info_table.rows[0].cells[1].text = contact_block
            style_cell(info_table.rows[0].cells[1], horizontal=WD_ALIGN_PARAGRAPH.CENTER, font_size_pt=10)

            info_table.rows[4].cells[1].text = (
                f"Fakultet dekani: ____________ {student.faculty_dean}\n"
                "(imzo)                                (F.I.Sh.)"
            )
            info_table.rows[4].cells[2].text = (
                "Korxonadan ajratilgan\n"
                f"amaliyot rahbari: ____________ {student.company_director}\n"
                "(imzo)                      (F.I.Sh.)"
            )
            style_cell(info_table.rows[4].cells[1], horizontal=WD_ALIGN_PARAGRAPH.CENTER, font_size_pt=10)
            style_cell(info_table.rows[4].cells[2], horizontal=WD_ALIGN_PARAGRAPH.CENTER, font_size_pt=10)

    normalize_document_formatting(doc)
    return doc


def generate_yollanma_document(student, course, practice_type, start_date_str, end_date_str):
    doc = build_manual_template_document("yollanma")
    if doc is None:
        return build_fallback_document("Yo'llanma", student, course, practice_type, start_date_str, end_date_str)

    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)
    direction_label = student.direction or student.faculty

    if 4 < len(doc.paragraphs):
        doc.paragraphs[4].text = (
            "Muxammad al-Xorazmiy nomidagi Toshkent axborot texnologiyalari universiteti O'zbekiston Respublikasi oliy va o'rta maxsus ta'lim vazirligi tomonidan tasdiqlangan "
            "“Oliy ta'lim muassasalari talabalarining malaka amaliyotini o'tash tartibi to'g'risidagi nizom” va"
        )
    if 5 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[5], [(student.company, True)])
        center_paragraph(doc.paragraphs[5], font_size_pt=11)
    if 6 < len(doc.paragraphs):
        doc.paragraphs[6].text = "hamda Muxammad al-Xorazmiy nomidagi TATU o'rtasidagi shartnoma asosida"
    if 7 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[7], [(direction_label, True), (" ta'lim yo'nalishida tahsil olayotgan talaba", False)])
    if 8 < len(doc.paragraphs):
        doc.paragraphs[8].text = ""
    if 9 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[9], [(student.full_name, True), (f" {practice_type.lower()}ni o'tash uchun", False)])
    if 10 < len(doc.paragraphs):
        doc.paragraphs[10].text = ""
    if 12 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[12], [(student.company, True), (" ga yubormoqda.", False)])
    if 13 < len(doc.paragraphs):
        doc.paragraphs[13].text = ""
    if 15 < len(doc.paragraphs):
        doc.paragraphs[15].text = f"Amaliyot muddati:     {start_parts['short']}-yildan,       {end_parts['short']} yilgacha"
    if 19 < len(doc.paragraphs):
        doc.paragraphs[19].text = f"TATU dan ketdi: \"{start_parts['day']}\" {start_parts['month']} {start_parts['year']}-yil                              Korxonaga keldi: \"{start_parts['day']}\" {start_parts['month']} {start_parts['year']}-yil"
    if 20 < len(doc.paragraphs):
        doc.paragraphs[20].text = f"Fakultet dekani: _____________________ {student.faculty_dean}                                 Korxona rahbari: _____________________ {student.company_director}"
    if 21 < len(doc.paragraphs):
        doc.paragraphs[21].text = ""
    if 26 < len(doc.paragraphs):
        doc.paragraphs[26].text = f"Korxonadan ketdi: \"{end_parts['day']}\" {end_parts['month']} {end_parts['year']}-yil                               TATU ga keldi: \"{end_parts['day']}\" {end_parts['month']} {end_parts['year']}-yil"
    if 28 < len(doc.paragraphs):
        doc.paragraphs[28].text = f"Korxona rahbari: _____________________ {student.company_director}                                 Fakultet dekani: _____________________ {student.faculty_dean}"
    if 29 < len(doc.paragraphs):
        doc.paragraphs[29].text = ""
    if 37 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[37], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[37], font_size_pt=11)
    if 38 < len(doc.paragraphs):
        doc.paragraphs[38].text = ""
    if 39 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[39], [(student.company, True), (" ga amaliyotga keldi.", False)])
    if 40 < len(doc.paragraphs):
        doc.paragraphs[40].text = ""
    if 41 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[41], [("Texnika xavfsizligi bo'yicha sinovni “", False), ("5", True), ("” bahoga topshirdi va amaliyotni o'tashga ruxsat berildi.", False)])
    if 44 < len(doc.paragraphs):
        doc.paragraphs[44].text = f"Komissiya raisi       ____________    {student.practice_supervisor}"
    if 45 < len(doc.paragraphs):
        doc.paragraphs[45].text = ""
    if 49 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[49], [("Talaba ", False), (student.full_name, True)])
    if 50 < len(doc.paragraphs):
        doc.paragraphs[50].text = ""
    if 52 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[52], [("ishga qo'yildi ", False), ("Amaliyotchi", True)])
    if 53 < len(doc.paragraphs):
        doc.paragraphs[53].text = ""
    if 59 < len(doc.paragraphs):
        doc.paragraphs[59].text = f"Korxonadan tayinlagan amaliyot rahbari     ____________    {student.company_director}"

    normalize_document_formatting(doc)
    return doc


def generate_yollanma_document(student, course, practice_type, start_date_str, end_date_str):
    doc = build_manual_template_document("yollanma")
    if doc is None:
        return build_fallback_document("Yo'llanma", student, course, practice_type, start_date_str, end_date_str)

    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)
    direction_label = student.direction or student.faculty

    if 4 < len(doc.paragraphs):
        doc.paragraphs[4].text = (
            "Muxammad al-Xorazmiy nomidagi Toshkent axborot texnologiyalari universiteti O'zbekiston Respublikasi oliy va o'rta maxsus ta'lim vazirligi tomonidan tasdiqlangan "
            "\"Oliy ta'lim muassasalari talabalarining malaka amaliyotini o'tash tartibi to'g'risidagi nizom\" va"
        )
    if 5 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[5], [(student.company, True)])
        center_paragraph(doc.paragraphs[5], font_size_pt=11)
    if 6 < len(doc.paragraphs):
        doc.paragraphs[6].text = "hamda Muxammad al-Xorazmiy nomidagi TATU o'rtasidagi shartnoma asosida"
    if 7 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[7], [(direction_label, True)])
        center_paragraph(doc.paragraphs[7], font_size_pt=11)
    if 8 < len(doc.paragraphs):
        doc.paragraphs[8].text = ""
    if 9 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[9], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[9], font_size_pt=11)
    if 10 < len(doc.paragraphs):
        doc.paragraphs[10].text = ""
    if 12 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[12], [(student.company, True)])
        center_paragraph(doc.paragraphs[12], font_size_pt=11)
    if 13 < len(doc.paragraphs):
        doc.paragraphs[13].text = ""
    if 15 < len(doc.paragraphs):
        doc.paragraphs[15].text = f"Amaliyot muddati:     {start_parts['short']}-yildan,       {end_parts['short']} yilgacha"
    if 19 < len(doc.paragraphs):
        doc.paragraphs[19].text = f"TATU dan ketdi: \"{start_parts['day']}\" {start_parts['month']} {start_parts['year']}-yil                              Korxonaga keldi: \"{start_parts['day']}\" {start_parts['month']} {start_parts['year']}-yil"
    if 20 < len(doc.paragraphs):
        doc.paragraphs[20].text = f"Fakultet dekani: _____________________ {student.faculty_dean}                                 Korxona rahbari: _____________________ {student.company_director}"
    if 21 < len(doc.paragraphs):
        doc.paragraphs[21].text = ""
    if 26 < len(doc.paragraphs):
        doc.paragraphs[26].text = f"Korxonadan ketdi: \"{end_parts['day']}\" {end_parts['month']} {end_parts['year']}-yil                               TATU ga keldi: \"{end_parts['day']}\" {end_parts['month']} {end_parts['year']}-yil"
    if 28 < len(doc.paragraphs):
        doc.paragraphs[28].text = f"Korxona rahbari: _____________________ {student.company_director}                                 Fakultet dekani: _____________________ {student.faculty_dean}"
    if 29 < len(doc.paragraphs):
        doc.paragraphs[29].text = ""
    if 37 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[37], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[37], font_size_pt=11)
    if 38 < len(doc.paragraphs):
        doc.paragraphs[38].text = ""
    if 39 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[39], [(student.company, True)])
        center_paragraph(doc.paragraphs[39], font_size_pt=11)
    if 40 < len(doc.paragraphs):
        doc.paragraphs[40].text = ""
    if 41 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[41], [("Texnika xavfsizligi bo'yicha sinovni \"", False), ("5", True), ("\" bahoga topshirdi va amaliyotni o'tashga ruxsat berildi.", False)])
    if 44 < len(doc.paragraphs):
        doc.paragraphs[44].text = f"Komissiya raisi       ____________    {student.practice_supervisor}"
    if 45 < len(doc.paragraphs):
        doc.paragraphs[45].text = ""
    if 49 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[49], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[49], font_size_pt=11)
    if 50 < len(doc.paragraphs):
        doc.paragraphs[50].text = ""
    if 52 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[52], [("Amaliyotchi", True)])
        center_paragraph(doc.paragraphs[52], font_size_pt=11)
    if 53 < len(doc.paragraphs):
        doc.paragraphs[53].text = ""
    if 59 < len(doc.paragraphs):
        doc.paragraphs[59].text = f"Korxonadan tayinlagan amaliyot rahbari     ____________    {student.company_director}"

    normalize_document_formatting(doc)
    return doc


def generate_yollanma_document(student, course, practice_type, start_date_str, end_date_str):
    doc = build_manual_template_document("yollanma")
    if doc is None:
        return build_fallback_document("Yo'llanma", student, course, practice_type, start_date_str, end_date_str)

    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)
    direction_label = student.direction or student.faculty

    if 4 < len(doc.paragraphs):
        doc.paragraphs[4].text = (
            "Muxammad al-Xorazmiy nomidagi Toshkent axborot texnologiyalari universiteti O'zbekiston Respublikasi "
            "oliy va o'rta maxsus ta'lim vazirligi tomonidan tasdiqlangan “Oliy ta'lim muassasalari talabalarining "
            "malaka amaliyotini o'tash tartibi to'g'risidagi nizom” va"
        )
    if 5 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[5], [(student.company, True)])
        center_paragraph(doc.paragraphs[5], font_size_pt=11)
    if 6 < len(doc.paragraphs):
        doc.paragraphs[6].text = "hamda Muxammad al-Xorazmiy nomidagi TATU o'rtasidagi shartnoma asosida"
    if 7 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[7], [(direction_label, True), (" ta'lim yo'nalishida tahsil olayotgan talaba", False)])
    if 8 < len(doc.paragraphs):
        doc.paragraphs[8].text = ""
    if 9 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[9], [(student.full_name, True), (f" {practice_type.lower()}ni o'tash uchun", False)])
    if 10 < len(doc.paragraphs):
        doc.paragraphs[10].text = ""
    if 12 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[12], [(student.company, True), (" ga yubormoqda.", False)])
    if 13 < len(doc.paragraphs):
        doc.paragraphs[13].text = ""
    if 15 < len(doc.paragraphs):
        doc.paragraphs[15].text = f"Amaliyot muddati:     {start_parts['short']}-yildan,       {end_parts['short']} yilgacha"
    if 19 < len(doc.paragraphs):
        doc.paragraphs[19].text = f"TATU dan ketdi: \"{start_parts['day']}\" {start_parts['month']} {start_parts['year']}-yil                              Korxonaga keldi: \"{start_parts['day']}\" {start_parts['month']} {start_parts['year']}-yil"
    if 20 < len(doc.paragraphs):
        doc.paragraphs[20].text = f"Fakultet dekani: _____________________ {student.faculty_dean}                                 Korxona rahbari: _____________________ {student.company_director}"
    if 21 < len(doc.paragraphs):
        doc.paragraphs[21].text = ""
    if 26 < len(doc.paragraphs):
        doc.paragraphs[26].text = f"Korxonadan ketdi: \"{end_parts['day']}\" {end_parts['month']} {end_parts['year']}-yil                               TATU ga keldi: \"{end_parts['day']}\" {end_parts['month']} {end_parts['year']}-yil"
    if 28 < len(doc.paragraphs):
        doc.paragraphs[28].text = f"Korxona rahbari: _____________________ {student.company_director}                                 Fakultet dekani: _____________________ {student.faculty_dean}"
    if 29 < len(doc.paragraphs):
        doc.paragraphs[29].text = ""
    if 37 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[37], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[37], font_size_pt=11)
    if 38 < len(doc.paragraphs):
        doc.paragraphs[38].text = ""
    if 39 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[39], [(student.company, True), (" ga  amaliyotga keldi.", False)])
    if 40 < len(doc.paragraphs):
        doc.paragraphs[40].text = ""
    if 41 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[41], [("Texnika xavfsizligi bo'yicha sinovni “", False), ("5", True), ("” bahoga topshirdi va amaliyotni o'tashga ruxsat berildi.", False)])
    if 44 < len(doc.paragraphs):
        doc.paragraphs[44].text = f"Komissiya raisi       ____________    {student.practice_supervisor}"
    if 45 < len(doc.paragraphs):
        doc.paragraphs[45].text = ""
    if 49 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[49], [("Talaba ", False), (student.full_name, True)])
    if 50 < len(doc.paragraphs):
        doc.paragraphs[50].text = ""
    if 52 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[52], [("ishga qo'yildi ", False), ("Amaliyotchi", True)])
    if 53 < len(doc.paragraphs):
        doc.paragraphs[53].text = ""
    if 59 < len(doc.paragraphs):
        doc.paragraphs[59].text = f"Korxonadan tayinlagan amaliyot rahbari     ____________    {student.company_director}"

    normalize_document_formatting(doc)
    return doc


def format_full_date_label(date_str):
    parts = parse_date_parts(date_str)
    return f"{parts['day']} {parts['month']} {parts['year']}-yil"


def render_docx_template(template_key, context):
    template_path = get_available_runtime_template_path(template_key)
    if not template_path:
        return None

    template = DocxTemplate(template_path)
    if not template.get_undeclared_template_variables():
        return None
    template.render(context)
    buffer = io.BytesIO()
    template.save(buffer)
    buffer.seek(0)
    return Document(buffer)


def build_common_template_context(student, course, practice_type, start_date_str, end_date_str):
    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)
    direction_label = student.direction or student.faculty

    return {
        "F_I_SH": student.full_name,
        "fakulteti": direction_label,
        "Fakulteti": direction_label,
        "kurs": str(course),
        "amaliyot_turi": practice_type,
        "amaliyot_otish_joyi": student.company,
        "Amaliyot_otish_joyi": student.company,
        "boshlanish_sanasi": format_full_date_label(start_date_str),
        "tugash_sanasi": format_full_date_label(end_date_str),
        "amaliyotga_kelgan_sana": format_full_date_label(start_date_str),
        "amaliyotdan_ketgan_sana": format_full_date_label(end_date_str),
        "korxona_telefoni": student.company_phone,
        "korxona_rahbari": student.company_director,
        "Korxona_rahbari": student.company_director,
        "amaliyot_rahbari": student.company_director,
        "korxonadagi_amaliyot_rahbari": student.company_director,
        "universitetdagi_amaliyot_rahbari": student.practice_supervisor,
        "kafedra": student.faculty,
        "kafedra_mudiri": student.department_head,
        "Kafedra_mudiri": student.department_head,
        "Fakultet_dekani": student.faculty_dean,
        "talabalar_soni": str(get_company_student_count(student.company)),
        "korxonadagi_lavozimi": "Amaliyot rahbari",
        "korxona_vakili_fish_lavozimi": f"{student.company_director}, amaliyot rahbari",
        "Lavozimi": "Amaliyotchi",
        "Baho": "5",
        "Komissiya_raisi": student.practice_supervisor,
        "TATUdan_ketgan_sana": start_parts["day"],
        "TATUdan_ketgan_oy": start_parts["month"],
        "Korxonaga_kelgan_sana": start_parts["day"],
        "Korxonaga_kelgan_oy": start_parts["month"],
        "Korxonadan_ketgan_sana": end_parts["day"],
        "Korxonadan_ketgan_oy": end_parts["month"],
        "Tatuga_kelgan_sana": end_parts["day"],
        "Tatuga_kelgan_oy": end_parts["month"],
    }


def build_manual_template_document(template_key):
    template_path = get_available_runtime_template_path(template_key)
    if not template_path:
        return None
    doc = Document(template_path)
    normalize_document_formatting(doc)
    return doc


def generate_shartnoma_from_fixed_template(student, course, practice_type, start_date_str, end_date_str):
    doc = build_manual_template_document("shartnoma")
    if doc is None:
        return None

    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)
    company_count = get_company_student_count(student.company)

    if len(doc.paragraphs) > 9:
        doc.paragraphs[0].text = f"TALABALARNING {practice_type.upper()}NI TASHKIL ETISH BO‘YICHA SHARTNOMA"
        doc.paragraphs[3].text = f"Toshkent sh.                                                                 {start_parts['year']}-yil «{start_parts['day']}»  {start_parts['month']}"
        doc.paragraphs[6].text = student.company
        doc.paragraphs[9].text = student.company_director

    if len(doc.tables) >= 2:
        main_table = doc.tables[0]
        if len(main_table.rows) > 2 and len(main_table.rows[2].cells) >= 7:
            main_table.rows[2].cells[1].text = student.direction or student.faculty
            main_table.rows[2].cells[2].text = str(course)
            main_table.rows[2].cells[3].text = practice_type
            main_table.rows[2].cells[4].text = str(company_count)
            main_table.rows[2].cells[5].text = start_parts["short"]
            main_table.rows[2].cells[6].text = end_parts["short"]

        info_table = doc.tables[1]
        if len(info_table.rows) > 4 and len(info_table.rows[0].cells) >= 3:
            company_block = f"Korxona\n{student.company}\n{student.company_address}\nTel.: {student.company_phone}"
            info_table.rows[0].cells[1].text = company_block
            info_table.rows[0].cells[2].text = company_block
            info_table.rows[4].cells[0].text = (
                f"Kafedra mudiri: ____________ {student.department_head}\n"
                "(imzo)                                (F.I.Sh.)"
            )
            info_table.rows[4].cells[1].text = (
                f"Fakultet dekani: ____________ {student.faculty_dean}\n"
                "(imzo)                                (F.I.Sh.)"
            )
            info_table.rows[4].cells[2].text = (
                "Korxonadan ajratilgan\n"
                f"amaliyot rahbari: ____________ {student.company_director}\n"
                "(imzo)                      (F.I.Sh.)"
            )
    normalize_document_formatting(doc)
    return doc


def generate_kundalik_from_fixed_template(student, course, practice_type, start_date_str, end_date_str):
    doc = build_manual_template_document("kundalik")
    if doc is None:
        return None

    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)
    academic_year = f"{start_parts['year']}/{int(start_parts['year']) + 1}"
    direction_label = student.direction or student.faculty

    replacements = {
        5: f"{direction_label} ta’lim yo‘nalishi {course} - bosqich talabasi {student.full_name} ning {academic_year} o‘quv yilidagi {practice_type.lower()}",
        10: f"1.1. Amaliyot joyi va muddati {student.company}, {student.company_address} muddati: {start_parts['short']} dan   {end_parts['short']} gacha",
        12: f"Universitetdan: ____________ {student.practice_supervisor}",
        14: f"Korxonadan: ____________ {student.company_director}",
        16: f"1.3. Talaba {student.full_name} ga “{student.faculty}” kafedrasidan berilgan individual topshiriqlar {practice_type.lower()} bo‘yicha kundalik yuritish va hisobot tayyorlash.",
        17: f"1.4. Amaliyotga keldi: {start_parts['year']}-yil {start_parts['day']}-{start_parts['month']},  ketdi: {end_parts['year']}-yil {end_parts['day']}-{end_parts['month']}",
        39: f"Universitetdan __________ {student.practice_supervisor} {start_parts['year']}-yil “{start_parts['day']}”  {start_parts['month']}",
        42: f"Korxonadan    __________ {student.company_director} {end_parts['year']}-yil “{end_parts['day']}” {end_parts['month']}",
        56: f"Korxonadan rahbar: ____________ {student.company_director}",
        59: f"M.O‘. \t\t{end_parts['year']}-yil       “_____”  ____________________",
        63: f"Universitetdan amaliyot rahbari: _____________ {student.practice_supervisor}",
        65: f"Kafedra mudiri: _____________ {student.department_head}",
        67: f"{end_parts['year']}-yil “____” ___________",
    }
    for index, text in replacements.items():
        if index < len(doc.paragraphs):
            doc.paragraphs[index].text = text

    if course == 3 and 25 < len(doc.paragraphs):
        doc.paragraphs[25].text = "2.6. Amaliyot muddati 3-bosqichda o‘quv reja asosida belgilanadi."

    normalize_document_formatting(doc)
    return doc


def generate_yollanma_from_fixed_template(student, course, practice_type, start_date_str, end_date_str):
    doc = build_manual_template_document("yollanma")
    if doc is None:
        return None

    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)
    direction_label = student.direction or student.faculty

    replacements = {
        4: f"Muxammad al-Xorazmiy nomidagi Toshkent axborot texnologiyalari universiteti O‘zbekiston Respublikasi oliy va o‘rta maxsus ta’lim vazirligi tomonidan tasdiqlangan “Oliy ta’lim muassasalari talabalarining malaka amaliyotini o‘tash tartibi to‘g‘risidagi nizom” va {student.company}",
        6: f"{student.company} hamda Muxammad al-Xorazmiy nomidagi TATU o‘rtasidagi shartnoma asosida",
        7: f"{direction_label} ta’lim yo‘nalishida tahsil olayotgan talaba",
        9: f"{student.full_name} {practice_type.lower()}ni o‘tash uchun",
        12: f"{student.company} ga yubormoqda.",
        15: f"Amaliyot muddati:     {start_parts['short']}-yildan,       {end_parts['short']} yilgacha",
        19: f"TATU dan ketdi: “{start_parts['day']}” {start_parts['month']} {start_parts['year']}-yil                              Korxonaga keldi: “{start_parts['day']}” {start_parts['month']} {start_parts['year']}-yil",
        20: f"Fakultet dekani: _____________________ {student.faculty_dean}                                 Korxona rahbari: _____________________ {student.company_director}",
        23: "M.O‘.: ______________                                                              M.O‘.: ______________",
        26: f"Korxonadan ketdi: “{end_parts['day']}” {end_parts['month']} {end_parts['year']}-yil                               TATU ga keldi: “{end_parts['day']}” {end_parts['month']} {end_parts['year']}-yil",
        28: f"Korxona rahbari: _____________________ {student.company_director}                                 Fakultet dekani: _____________________ {student.faculty_dean}",
        31: "M.O‘.: ______________                                                              M.O‘.: ______________",
        37: student.full_name,
        39: f"{student.company} ga amaliyotga keldi.",
        41: "Texnika xavfsizligi bo‘yicha sinovni “5” bahoga topshirdi va amaliyotni o‘tashga ruxsat berildi.",
        44: f"Komissiya raisi: ____________ {student.practice_supervisor}",
        49: f"Talaba: {student.full_name}",
        52: "ishga qo‘yildi: Amaliyotchi",
        59: f"Korxonadan tayinlagan amaliyot rahbari: ____________ {student.company_director}",
    }
    for index, text in replacements.items():
        if index < len(doc.paragraphs):
            doc.paragraphs[index].text = text

    normalize_document_formatting(doc)
    return doc


def format_yollanma_rendered_document(doc, student, start_date_str, end_date_str):
    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)

    replacements = {
        11: (
            f"TATU dan ketdi: “{start_parts['day']}” {start_parts['month']} {start_parts['year']}-yil"
            f"                              Korxonaga keldi: “{start_parts['day']}” {start_parts['month']} {start_parts['year']}-yil"
        ),
        12: (
            f"Fakultet dekani: ____________ {student.faculty_dean}"
            f"                              Korxona rahbari: ____________ {student.company_director}"
        ),
        14: "M.O‘.: ____________                                                              M.O‘.: ____________",
        17: (
            f"Korxonadan ketdi: “{end_parts['day']}” {end_parts['month']} {end_parts['year']}-yil"
            f"                               TATU ga keldi: “{end_parts['day']}” {end_parts['month']} {end_parts['year']}-yil"
        ),
        19: (
            f"Korxona rahbari: ____________ {student.company_director}"
            f"                              Fakultet dekani: ____________ {student.faculty_dean}"
        ),
        21: "M.O‘.: ____________                                                              M.O‘.: ____________",
        27: f"Talaba: {student.full_name}    Korxona: {student.company}    amaliyotga keldi.",
        30: f"Komissiya raisi: ____________ {student.practice_supervisor}",
        35: f"Talaba: {student.full_name}    Lavozimi: Amaliyotchi",
        39: f"Korxonadan tayinlangan amaliyot rahbari: ____________ {student.company_director}",
        45: f"Baho: _____ ballga topshirdi.           Kafedra mudiri: ____________ {student.department_head}           Sana: ______",
    }

    for index, text in replacements.items():
        if index < len(doc.paragraphs):
            doc.paragraphs[index].text = text

    normalize_document_formatting(doc)
    return doc


def generate_shartnoma_document(student, course, practice_type, start_date_str, end_date_str):
    context = build_common_template_context(student, course, practice_type, start_date_str, end_date_str)
    rendered_document = render_docx_template("shartnoma", context)
    if rendered_document is not None:
        return rendered_document
    fixed_template_document = generate_shartnoma_from_fixed_template(student, course, practice_type, start_date_str, end_date_str)
    if fixed_template_document is not None:
        return fixed_template_document
    return build_fallback_document("Shartnoma", student, course, practice_type, start_date_str, end_date_str)


def generate_kundalik_document(student, course, practice_type, start_date_str, end_date_str):
    context = build_common_template_context(student, course, practice_type, start_date_str, end_date_str)
    rendered_document = render_docx_template("kundalik", context)
    if rendered_document is not None:
        return rendered_document
    fixed_template_document = generate_kundalik_from_fixed_template(student, course, practice_type, start_date_str, end_date_str)
    if fixed_template_document is not None:
        return fixed_template_document
    return build_fallback_document("Kundalik", student, course, practice_type, start_date_str, end_date_str)


def generate_yollanma_document(student, course, practice_type, start_date_str, end_date_str):
    context = build_common_template_context(student, course, practice_type, start_date_str, end_date_str)
    rendered_document = render_docx_template("yollanma", context)
    if rendered_document is not None:
        return format_yollanma_rendered_document(rendered_document, student, start_date_str, end_date_str)
    fixed_template_document = generate_yollanma_from_fixed_template(student, course, practice_type, start_date_str, end_date_str)
    if fixed_template_document is not None:
        return fixed_template_document
    return build_fallback_document("Yo'llanma", student, course, practice_type, start_date_str, end_date_str)


def generate_kundalik_document(student, course, practice_type, start_date_str, end_date_str):
    context = build_common_template_context(student, course, practice_type, start_date_str, end_date_str)
    rendered_document = render_docx_template("kundalik", context)
    if rendered_document is not None:
        return rendered_document

    doc = build_manual_template_document("kundalik")
    if doc is None:
        return build_fallback_document("Kundalik", student, course, practice_type, start_date_str, end_date_str)

    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)
    academic_year = f"{start_parts['year']}/{int(start_parts['year']) + 1}"
    direction_label = student.direction or student.faculty

    base_replacements = {
        5: f"{direction_label} ta'lim yo'nalishi {course} - bosqich talabasi {student.full_name} ning {academic_year} o'quv yilidagi {practice_type.lower()}",
        10: f"1.1. Amaliyot joyi va muddati {student.company}, {student.company_address} muddati: {start_parts['short']} dan   {end_parts['short']} gacha",
        12: "Universitetdan",
        14: "Korxonadan",
        16: f"1.3. Talaba {student.full_name} ga \"{student.faculty}\" kafedrasidan berilgan individual topshiriqlar {practice_type.lower()} bo'yicha kundalik yuritish va hisobot tayyorlash.",
        17: f"1.4. Amaliyotga keldi: {format_document_date(start_date_str)},  ketdi: {format_document_date(end_date_str)}",
        56: "Korxonadan rahbar",
        59: f"M.O'. \t\t{end_parts['year']}-yil       \"_____\"  ____________________",
        63: "Universitetdan amaliyot rahbari",
        65: "Kafedra mudiri",
        67: f"{end_parts['year']}-yil \"____\" ___________",
    }
    for index, text in base_replacements.items():
        if index < len(doc.paragraphs):
            doc.paragraphs[index].text = text

    if course == 3 and 25 < len(doc.paragraphs):
        doc.paragraphs[25].text = "2.6. Amaliyot muddati 3-bosqichda o'quv reja asosida belgilanadi."

    if 39 < len(doc.paragraphs):
        set_paragraph_segments(
            doc.paragraphs[39],
            [
                ("Universitetdan ", False),
                ("____________ ", False),
                (student.practice_supervisor, True),
                (" ", False),
                (format_document_date(start_date_str), True),
            ],
        )
    if 40 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[40], [("Imzo", False, True), ("\t\t         ", False), ("F.I.Sh.", False, True)])
    if 42 < len(doc.paragraphs):
        set_paragraph_segments(
            doc.paragraphs[42],
            [
                ("Korxonadan ", False),
                ("____________ ", False),
                (student.company_director, True),
                (" ", False),
                (format_document_date(end_date_str), True),
            ],
        )
    if 43 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[43], [("Imzo", False, True), ("\t\t         ", False), ("F.I.Sh.", False, True)])
    if 56 < len(doc.paragraphs):
        set_paragraph_segments(
            doc.paragraphs[56],
            [
                ("Korxonadan rahbar          ", False),
                ("________________ ", False),
                (student.company_director, True),
            ],
        )
    if 57 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[57], [("imzo", False, True), ("                                                              ", False), ("F.I.Sh.", False, True)])
    if 63 < len(doc.paragraphs):
        set_paragraph_segments(
            doc.paragraphs[63],
            [
                ("Universitetdan amaliyot rahbari      ", False),
                ("_____________ ", False),
                (student.practice_supervisor, True),
            ],
        )
    if 64 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[64], [("imzo", False, True), ("                                                              ", False), ("F.I.Sh.", False, True)])
    if 65 < len(doc.paragraphs):
        set_paragraph_segments(
            doc.paragraphs[65],
            [
                ("Kafedra mudiri                                ", False),
                ("_____________ ", False),
                (student.department_head, True),
            ],
        )
    if 66 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[66], [("imzo", False, True), ("                                                              ", False), ("F.I.Sh.", False, True)])

    normalize_document_formatting(doc)
    return doc





def generate_shartnoma_document(student, course, practice_type, start_date_str, end_date_str):
    doc = build_manual_template_document("shartnoma")
    if doc is None:
        return build_fallback_document("Shartnoma", student, course, practice_type, start_date_str, end_date_str)

    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)
    company_count = get_company_student_count(student.company)

    if len(doc.paragraphs) > 9:
        doc.paragraphs[0].text = f"TALABALARNING {practice_type.upper()}NI TASHKIL ETISH BO'YICHA SHARTNOMA"
        doc.paragraphs[3].text = f"Toshkent sh.                                                                 {start_parts['year']}-yil \"{start_parts['day']}\" {start_parts['month']}"
        doc.paragraphs[6].text = student.company
        doc.paragraphs[9].text = student.company_director

    if len(doc.tables) >= 2:
        main_table = doc.tables[0]
        if len(main_table.rows) > 2 and len(main_table.rows[2].cells) >= 7:
            row = main_table.rows[2].cells
            row[1].text = student.direction or student.faculty
            row[2].text = str(course)
            row[3].text = practice_type
            row[4].text = str(company_count)
            row[5].text = start_parts["short"]
            row[6].text = end_parts["short"]
            style_cell(row[0], font_size_pt=10)
            style_cell(row[1], font_size_pt=9)
            style_cell(row[2], font_size_pt=10)
            style_cell(row[3], font_size_pt=9)
            style_cell(row[4], font_size_pt=10)
            style_cell(row[5], font_size_pt=9)
            style_cell(row[6], font_size_pt=9)

        info_table = doc.tables[1]
        if len(info_table.rows) > 4 and len(info_table.rows[0].cells) >= 3:
            contact_block = f"{student.company_address}\nTel.: {student.company_phone}"
            info_table.rows[0].cells[1].text = contact_block
            style_cell(info_table.rows[0].cells[1], horizontal=WD_ALIGN_PARAGRAPH.CENTER, font_size_pt=10)

            info_table.rows[4].cells[1].text = (
                f"Fakultet dekani: ____________ {student.faculty_dean}\n"
                "(imzo)                                (F.I.Sh.)"
            )
            info_table.rows[4].cells[2].text = (
                "Korxonadan ajratilgan\n"
                f"amaliyot rahbari: ____________ {student.company_director}\n"
                "(imzo)                      (F.I.Sh.)"
            )
            style_cell(info_table.rows[4].cells[1], horizontal=WD_ALIGN_PARAGRAPH.CENTER, font_size_pt=10)
            style_cell(info_table.rows[4].cells[2], horizontal=WD_ALIGN_PARAGRAPH.CENTER, font_size_pt=10)

    normalize_document_formatting(doc)
    return doc


def generate_yollanma_document(student, course, practice_type, start_date_str, end_date_str):
    doc = build_manual_template_document("yollanma")
    if doc is None:
        return build_fallback_document("Yo'llanma", student, course, practice_type, start_date_str, end_date_str)

    start_parts = parse_date_parts(start_date_str)
    end_parts = parse_date_parts(end_date_str)
    direction_label = student.direction or student.faculty

    if 4 < len(doc.paragraphs):
        doc.paragraphs[4].text = (
            "Muxammad al-Xorazmiy nomidagi Toshkent axborot texnologiyalari universiteti O'zbekiston Respublikasi oliy va o'rta maxsus ta'lim vazirligi tomonidan tasdiqlangan "
            "\"Oliy ta'lim muassasalari talabalarining malaka amaliyotini o'tash tartibi to'g'risidagi nizom\" va"
        )
    if 5 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[5], [(student.company, True)])
        center_paragraph(doc.paragraphs[5], font_size_pt=11)
    if 6 < len(doc.paragraphs):
        doc.paragraphs[6].text = "hamda Muxammad al-Xorazmiy nomidagi TATU o'rtasidagi shartnoma asosida"
    if 7 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[7], [(direction_label, True)])
        center_paragraph(doc.paragraphs[7], font_size_pt=11)
    if 8 < len(doc.paragraphs):
        doc.paragraphs[8].text = ""
    if 9 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[9], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[9], font_size_pt=11)
    if 10 < len(doc.paragraphs):
        doc.paragraphs[10].text = ""
    if 12 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[12], [(student.company, True)])
        center_paragraph(doc.paragraphs[12], font_size_pt=11)
    if 13 < len(doc.paragraphs):
        doc.paragraphs[13].text = ""
    if 15 < len(doc.paragraphs):
        doc.paragraphs[15].text = f"Amaliyot muddati:     {start_parts['short']}-yildan,       {end_parts['short']} yilgacha"
    if 19 < len(doc.paragraphs):
        doc.paragraphs[19].text = f"TATU dan ketdi: \"{start_parts['day']}\" {start_parts['month']} {start_parts['year']}-yil                              Korxonaga keldi: \"{start_parts['day']}\" {start_parts['month']} {start_parts['year']}-yil"
    if 20 < len(doc.paragraphs):
        doc.paragraphs[20].text = f"Fakultet dekani: _____________________ {student.faculty_dean}                                 Korxona rahbari: _____________________ {student.company_director}"
    if 21 < len(doc.paragraphs):
        doc.paragraphs[21].text = ""
    if 26 < len(doc.paragraphs):
        doc.paragraphs[26].text = f"Korxonadan ketdi: \"{end_parts['day']}\" {end_parts['month']} {end_parts['year']}-yil                               TATU ga keldi: \"{end_parts['day']}\" {end_parts['month']} {end_parts['year']}-yil"
    if 28 < len(doc.paragraphs):
        doc.paragraphs[28].text = f"Korxona rahbari: _____________________ {student.company_director}                                 Fakultet dekani: _____________________ {student.faculty_dean}"
    if 29 < len(doc.paragraphs):
        doc.paragraphs[29].text = ""
    if 37 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[37], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[37], font_size_pt=11)
    if 38 < len(doc.paragraphs):
        doc.paragraphs[38].text = ""
    if 39 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[39], [(student.company, True)])
        center_paragraph(doc.paragraphs[39], font_size_pt=11)
    if 40 < len(doc.paragraphs):
        doc.paragraphs[40].text = ""
    if 41 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[41], [("Texnika xavfsizligi bo'yicha sinovni \"", False), ("5", True), ("\" bahoga topshirdi va amaliyotni o'tashga ruxsat berildi.", False)])
    if 44 < len(doc.paragraphs):
        doc.paragraphs[44].text = f"Komissiya raisi       ____________    {student.practice_supervisor}"
    if 45 < len(doc.paragraphs):
        doc.paragraphs[45].text = ""
    if 49 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[49], [(student.full_name, True)])
        center_paragraph(doc.paragraphs[49], font_size_pt=11)
    if 50 < len(doc.paragraphs):
        doc.paragraphs[50].text = ""
    if 52 < len(doc.paragraphs):
        set_paragraph_segments(doc.paragraphs[52], [("Amaliyotchi", True)])
        center_paragraph(doc.paragraphs[52], font_size_pt=11)
    if 53 < len(doc.paragraphs):
        doc.paragraphs[53].text = ""
    if 59 < len(doc.paragraphs):
        doc.paragraphs[59].text = f"Korxonadan tayinlagan amaliyot rahbari     ____________    {student.company_director}"

    normalize_document_formatting(doc)
    return doc



