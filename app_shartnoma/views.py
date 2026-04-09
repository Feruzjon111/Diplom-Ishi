from django.http import HttpResponse
from docx import Document
from app_excel.models import Student
import io
import os
from django.conf import settings

def generate_contracts_download(request):
    students = Student.objects.all()

    if not students.exists():
        return HttpResponse(" Talabalar mavjud emas.")

    template_path = os.path.join(
        settings.BASE_DIR,
        'app_shartnoma',
        'contract_templates',
        'shartnoma_template.docx'
    )

    if not os.path.exists(template_path):
        return HttpResponse(f"Fayl topilmadi: {template_path}", status=404)

    s = students.first()

    doc = Document(template_path)

    for p in doc.paragraphs:
        p.text = p.text.replace('FISH_STUDENT', s.full_name)
        p.text = p.text.replace('GROUP_STUDENT', s.group)
        p.text = p.text.replace('COMPANY_NAME', s.company)
        p.text = p.text.replace('COMPANY_ADDRESS', s.company_address)
        p.text = p.text.replace('COMPANY_PHONE', s.company_phone)
        p.text = p.text.replace('FISH_DIRECTOR', s.company_director)
        p.text = p.text.replace('FISH_SUPERVISOR', s.practice_supervisor)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    response = HttpResponse(buffer.getvalue(),
                            content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    response['Content-Disposition'] = f'attachment; filename="{s.full_name.replace(" ", "_")}_shartnoma.docx"'
    return response
