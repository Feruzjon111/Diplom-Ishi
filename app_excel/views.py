from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
import openpyxl
from django.http import HttpResponse, JsonResponse, FileResponse
from rest_framework.decorators import api_view
from .models import Student, Payment, Profile
import io, os, zipfile, uuid
from django.conf import settings
from datetime import datetime
from decimal import Decimal, InvalidOperation
from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth import update_session_auth_hash
from django.contrib.auth.models import User
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from .serializers import StudentSerializer
from rest_framework import viewsets
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status
from django.views.decorators.csrf import csrf_exempt
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


def get_template_status_cards():
    runtime_paths = get_runtime_template_paths()
    return [
        {
            "label": "Shartnoma",
            "source_key": "shartnoma",
            "runtime_ready": os.path.exists(runtime_paths["shartnoma"]),
        },
        {
            "label": "Kundalik",
            "source_key": "kundalik",
            "runtime_ready": os.path.exists(runtime_paths["kundalik"]),
        },
        {
            "label": "Yo‘llanma",
            "source_key": "yollanma",
            "runtime_ready": os.path.exists(runtime_paths["yollanma"]),
        },
    ]


def validate_runtime_templates():
    missing = [
        name for name, path in get_runtime_template_paths().items()
        if not os.path.exists(path)
    ]
    return missing


def get_practice_type(course):
    return "Ishlab chiqarish amaliyoti" if int(course) == 3 else "Bitiruv oldi amaliyoti"


def sanitize_filename(value):
    safe = "".join(char if char.isalnum() or char in (" ", "-", "_") else "_" for char in (value or "").strip())
    return "_".join(safe.split()) or "talaba"


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


def generate_shartnoma_document(student, course, practice_type, start_date_str, end_date_str):
    doc = Document(get_runtime_template_paths()["shartnoma"])
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
    doc = Document(get_runtime_template_paths()["kundalik"])
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
        f"{practice_type.lower()} bo‘yicha kundalik yuritish va hisobot tayyorlash."
    )
    doc.paragraphs[17].text = (
        f"1.4. Amaliyotga keldi: {start_parts['year']}-yil «{start_parts['day']}» {start_parts['month']}, "
        f"ketdi: {end_parts['year']}-yil «{end_parts['day']}» {end_parts['month']}"
    )
    if course == 3:
        doc.paragraphs[25].text = "2.6. Amaliyot muddati 3-bosqichda o‘quv reja asosida belgilanadi."
    for idx in (12, 14):
        if idx < len(doc.paragraphs):
            center_paragraph(doc.paragraphs[idx], font_size_pt=11)
    for idx in (13, 15):
        if idx < len(doc.paragraphs):
            center_paragraph(doc.paragraphs[idx], font_size_pt=9)
    return doc


def generate_yollanma_document(student, course, practice_type, start_date_str, end_date_str):
    doc = Document(get_runtime_template_paths()["yollanma"])
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
    context = {
        "profile": profile,
        "student_count": students.count(),
        "enterprise_count": students.values("company").distinct().count(),
        "document_count": students.count() * 3,
        "recent_students": students[:5],
        "template_cards": get_template_status_cards(),
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
    sheet.title = "Namuna"

    headers = [
        "Talabaning F.I.Sh.",
        "Tug'ilgan sana",
        "Pasport seriyasi va raqami",
        "JSHSHIR",
        "Fakultet",
        "Yo'nalish",
        "Guruh",
        "Kurs",
        "Amaliyot turi",
        "Amaliyot o'tash joyi",
        "Korxona nomi",
        "Korxona manzili",
        "Korxona rahbari F.I.Sh.",
        "Universitet mas'ul rahbari",
        "Telefon raqami",
    ]

    sample_rows = [
        ["Aliyev Bekzod Anvar o'g'li", "2003-02-14", "AA1234567", "30302141234567", "Axborot texnologiyalari", "Dasturiy injiniring", "SE-401", 4, "Bitiruv oldi amaliyoti", "IT bo'lim", "TechSoft MCHJ", "Toshkent sh., Yunusobod tumani", "Karimov Sardor Rustamovich", "Rasulov Dilshod Qodirovich", "+998901112233"],
        ["Karimova Maftuna Jamshid qizi", "2003-06-21", "AB2345678", "30306211234568", "Axborot texnologiyalari", "Kompyuter injiniring", "KI-402", 4, "Bitiruv oldi amaliyoti", "Tarmoq markazi", "Digital Systems", "Toshkent sh., Chilonzor tumani", "Ergashev Oybek Bahodirovich", "Rasulov Dilshod Qodirovich", "+998901112234"],
        ["Toshpo'latov Azizbek Shavkat o'g'li", "2002-11-03", "AC3456789", "30211031234569", "Kompyuter texnologiyalari", "Axborot xavfsizligi", "AT-403", 4, "Bitiruv oldi amaliyoti", "Xavfsizlik bo'limi", "SecureNet Group", "Toshkent sh., Shayxontohur tumani", "Mamatqulov Anvar Tohirovich", "Sobirova Nargiza Akmalovna", "+998901112235"],
        ["Usmonova Shahnoza Akbar qizi", "2003-01-18", "AD4567890", "30301181234570", "Raqamli iqtisodiyot", "Biznes analitika", "BA-404", 4, "Bitiruv oldi amaliyoti", "Analitika bo'limi", "Insight Analytics", "Toshkent sh., Mirobod tumani", "Rahimov Temur Baxtiyorovich", "Sobirova Nargiza Akmalovna", "+998901112236"],
        ["Qodirov Shohruh Ilhom o'g'li", "2003-07-09", "AE5678901", "30307091234571", "Axborot texnologiyalari", "Sun'iy intellekt", "AI-405", 4, "Bitiruv oldi amaliyoti", "ML laboratoriya", "AI Vision Lab", "Toshkent sh., Olmazor tumani", "Jo'rayev Mirjalol Otabekovich", "Yo'ldoshev Bekzod Alimuhamedov", "+998901112237"],
        ["Nazarova Diyora Baxtiyor qizi", "2002-09-26", "AF6789012", "30209261234572", "Axborot texnologiyalari", "Dasturiy injiniring", "SE-406", 4, "Bitiruv oldi amaliyoti", "Frontend guruh", "Creative Apps", "Toshkent sh., Uchtepa tumani", "Azimov Alisher Habibovich", "Yo'ldoshev Bekzod Alimuhamedov", "+998901112238"],
        ["Madumarov Sarvar Ulug'bek o'g'li", "2003-03-30", "AG7890123", "30303301234573", "Axborot texnologiyalari", "Backend dasturlash", "BE-407", 4, "Bitiruv oldi amaliyoti", "Backend bo'lim", "Core Systems", "Toshkent sh., Sergeli tumani", "Yoqubov Sherzod G'ayratovich", "Tursunova Mohinur Islomovna", "+998901112239"],
        ["Rahmatullayeva Sevinch Otabek qizi", "2003-12-11", "AH8901234", "30312111234574", "Kompyuter texnologiyalari", "Mobil ilovalar", "MB-408", 4, "Bitiruv oldi amaliyoti", "Mobil dasturlash", "Mobile Hub", "Toshkent sh., Yakkasaroy tumani", "Abdug'aniyev Sherali Tohir o'g'li", "Tursunova Mohinur Islomovna", "+998901112240"],
        ["Abdullayev Jahongir Siroj o'g'li", "2002-08-05", "AI9012345", "30208051234575", "Raqamli texnologiyalar", "Ma'lumotlar ilmi", "DS-409", 4, "Bitiruv oldi amaliyoti", "Data engineering", "Data Flow", "Toshkent sh., Bektemir tumani", "Hasanov Diyor Dilmurodovich", "Niyozova Shahlo Abdusattorovna", "+998901112241"],
        ["Saidova Nilufar Elyor qizi", "2003-04-17", "AJ0123456", "30304171234576", "Axborot texnologiyalari", "Kompyuter injiniring", "KI-410", 4, "Bitiruv oldi amaliyoti", "Texnik qo'llab-quvvatlash", "Support Plus", "Toshkent sh., Mirzo Ulug'bek tumani", "Toshmatov Jasur Komilovich", "Niyozova Shahlo Abdusattorovna", "+998901112242"],
        ["Erkinov Doston Ravshan o'g'li", "2003-05-29", "AK1122334", "30305291234577", "Axborot texnologiyalari", "Axborot tizimlari", "AT-411", 4, "Bitiruv oldi amaliyoti", "ERP bo'lim", "Enterprise Solutions", "Toshkent sh., Yangihayot tumani", "Sattorov Behruz Xayrulla o'g'li", "Qosimov Murodjon Iskandarovich", "+998901112243"],
        ["Yusupova Mohira Shukhrat qizi", "2002-10-08", "AL2233445", "30210081234578", "Kompyuter texnologiyalari", "UX/UI dizayn", "UX-412", 4, "Bitiruv oldi amaliyoti", "Dizayn studiya", "Pixel Studio", "Toshkent sh., Mirobod tumani", "Qobilov Sanjar Akramovich", "Qosimov Murodjon Iskandarovich", "+998901112244"],
        ["Raxmonov Ibrohim Doniyor o'g'li", "2003-09-19", "AM3344556", "30309191234579", "Raqamli iqtisodiyot", "Fintech", "FT-413", 4, "Bitiruv oldi amaliyoti", "To'lov tizimlari", "PayTech Group", "Toshkent sh., Chilonzor tumani", "Xudoyberdiyev Jamshid Ilyosovich", "Abdurahmonova Zarina Habibullayevna", "+998901112245"],
        ["Bozorova Madina Murod qizi", "2003-02-02", "AN4455667", "30302021234580", "Axborot texnologiyalari", "Web dasturlash", "WD-414", 4, "Bitiruv oldi amaliyoti", "Fullstack guruh", "WebNova", "Toshkent sh., Yunusobod tumani", "Soliyev Farrux Shavkatovich", "Abdurahmonova Zarina Habibullayevna", "+998901112246"],
        ["Hakimov Kamron Akmal o'g'li", "2002-07-27", "AO5566778", "30207271234581", "Kompyuter texnologiyalari", "Bulutli texnologiyalar", "CL-415", 4, "Bitiruv oldi amaliyoti", "Cloud infratuzilma", "Cloud Prime", "Toshkent sh., Olmazor tumani", "Rizoqulov Umid Raufovich", "Jo'rayeva Feruza Akbarovna", "+998901112247"],
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
        course = int(request.POST.get("course", 4))
        start_date = request.POST.get("start_date", "")
        end_date = request.POST.get("end_date", "")

        Student.objects.all().delete()

        try:
            if filename.endswith('.xlsx'):
                wb = openpyxl.load_workbook(uploaded_file)
                sheet = wb.active
                for row in sheet.iter_rows(min_row=2, values_only=True):
                    if not row or not row[0]:
                        continue
                    Student.objects.create(
                        full_name=(row[0] or "").strip(),
                        direction=(row[5] or "").strip(),
                        faculty=(row[4] or "").strip(),
                        group=str(row[6] or "").strip(),
                        company=(row[10] or "").strip(),
                        company_address=(row[11] or "").strip(),
                        company_director=(row[12] or "").strip(),
                        practice_supervisor=(row[13] or "").strip(),
                        company_phone=(row[14] or "").strip(),
                    )

            elif filename.endswith('.docx'):
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
                    )
            else:
                return render(request, 'app_excel/upload.html', {
                    'error': 'Faqat .xlsx yoki .docx fayl yuklash mumkin.',
                    'profile': profile,
                    'template_sources': get_template_sources(),
                    'template_cards': get_template_status_cards(),
                })

            request.session['uploaded'] = True
            request.session['course'] = course
            request.session['start_date'] = start_date
            request.session['end_date'] = end_date
            return redirect('upload_excel')

        except Exception as e:
            return render(request, 'app_excel/upload.html', {
                'error': str(e),
                'profile': profile,
                'template_sources': get_template_sources(),
                'template_cards': get_template_status_cards(),
            })

    uploaded = request.session.pop('uploaded', False)
    return render(request, 'app_excel/upload.html', {
        'uploaded': uploaded,
        'profile': profile,
        'template_sources': get_template_sources(),
        'template_cards': get_template_status_cards(),
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

    Student.objects.all().delete()
    return save_document_response(doc, f"{sanitize_filename(first_student.company)}_shartnoma.docx")




@login_required(login_url='login')
def export_all_documents_zip(request):
    missing_templates = validate_runtime_templates()
    if missing_templates:
        missing_labels = ", ".join(missing_templates)
        return HttpResponse(f"ZIP yaratish uchun faol DOCX shablonlar topilmadi: {missing_labels}.", status=400)

    user_profile = request.user.profile
    narx = 30000  # hujjatlarni yuklab olish narxi

    # Mablag' yetarli emas
    if user_profile.balance < narx:
        return HttpResponse("Hisobingizda yetarli mablag‘ mavjud emas.")

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

    user_profile.balance -= narx
    user_profile.save()

    Student.objects.all().delete()

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


@csrf_exempt
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
            return redirect("login")

    return render(request, "app_excel/register.html")


def logout_view(request):
    logout(request)
    return redirect("login")


@login_required(login_url='login')
def pay_view(request):
    user = request.user
    amount = 30000
    merchant_id = '123456'
    service_id = '654321'
    callback_url = 'http://127.0.0.1:8000/payment/callback/'

    merchant_trans_id = str(uuid.uuid4())

    payment = Payment.objects.create(
        user=user,
        amount=amount,
        merchant_trans_id=merchant_trans_id,
        status='pending'  # optional
    )

    context = {
        'payment': payment,
        'merchant_id': merchant_id,
        'service_id': service_id,
        'amount': amount,
        'merchant_trans_id': merchant_trans_id,
        'callback_url': callback_url,
    }

    return render(request, 'app_excel/pay.html', context)


@csrf_exempt
def click_prepare(request):
    return JsonResponse({'error': 0, 'error_note': 'Success'})


@csrf_exempt
def click_result(request):
    merchant_trans_id = request.POST.get("merchant_trans_id")
    click_trans_id = request.POST.get("click_trans_id")

    try:
        payment = Payment.objects.get(merchant_trans_id=merchant_trans_id)

        if not payment.paid:
            payment.paid = True
            payment.click_trans_id = click_trans_id
            payment.save()

            zip_buffer = generate_documents_zip()

            response = HttpResponse(zip_buffer, content_type='application/zip')
            response['Content-Disposition'] = f'attachment; filename=hujjatlar.zip'
            return response

        return JsonResponse({'error': 0, 'error_note': 'To‘lov avval amalga oshirilgan'})
    except Payment.DoesNotExist:
        return JsonResponse({'error': -5, 'error_note': 'To‘lov topilmadi'})



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


@csrf_exempt
def payment_callback(request):
    if request.method == "POST" or request.method == "GET":

        request.session['paid'] = True

        return redirect('export_all_documents_zip')

    return HttpResponse("To‘lov bekor qilindi yoki noto‘g‘ri so‘rov!", status=400)



@login_required
def top_up_balance(request):
    profile = ensure_profile(request.user)
    if request.method == "POST":
        raw_amount = (request.POST.get("amount") or "").strip()
        try:
            amount = Decimal(raw_amount)
        except (InvalidOperation, TypeError):
            messages.error(request, "To‘ldirish summasini to‘g‘ri kiriting.")
            return redirect('balance')

        if amount <= 0:
            messages.error(request, "Summa 0 dan katta bo‘lishi kerak.")
            return redirect('balance')

        profile.balance += amount
        profile.save(update_fields=["balance"])
        messages.success(request, f"Demo to‘lov muvaffaqiyatli: balansga {amount:,.0f} so‘m qo‘shildi.".replace(",", " "))
        return redirect('balance')

    return render(request, 'app_excel/balance.html', {"profile": profile})


@login_required
def balance_view(request):
    return render(request, 'app_excel/balance.html', {"profile": ensure_profile(request.user)})



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







