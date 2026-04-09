from django.core.management.base import BaseCommand
from django.contrib.auth import get_user_model

class Command(BaseCommand):
    def handle(self, *args, **options):
        User = get_user_model()
        admin, _ = User.objects.get_or_create(
            username='admin',
            defaults={'email': 'admin@amaliyotdocx.uz', 'is_staff': True, 'is_superuser': True},
        )
        admin.email = 'admin@amaliyotdocx.uz'
        admin.is_staff = True
        admin.is_superuser = True
        admin.set_password('Admin12345')
        admin.save()

        operator, _ = User.objects.get_or_create(
            username='operator',
            defaults={'email': 'operator@amaliyotdocx.uz', 'first_name': 'Default', 'last_name': 'Operator'},
        )
        operator.email = 'operator@amaliyotdocx.uz'
        operator.first_name = 'Default'
        operator.last_name = 'Operator'
        operator.set_password('Operator12345')
        operator.save()

        self.stdout.write("Admin va default operator foydalanuvchilari yaratildi.")
        self.stdout.write("Admin login: admin / Admin12345")
        self.stdout.write("Operator login: operator / Operator12345")
