from django.contrib import admin
from .models import Student


@admin.register(Student)
class StudentAdmin(admin.ModelAdmin):
    list_display = ("full_name", "direction", "group", "company", "faculty")
    search_fields = ("full_name", "direction", "group", "company", "faculty")
