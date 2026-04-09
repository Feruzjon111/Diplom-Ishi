from django.contrib.auth.models import User
from django.db import models


class Profile(models.Model):
    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name="profile")
    balance = models.DecimalField(max_digits=10, decimal_places=2, default=0)

    def __str__(self):
        return f"{self.user.username} profili"


class Student(models.Model):
    full_name = models.CharField(max_length=255)
    direction = models.CharField(max_length=255, blank=True, default="")
    group = models.CharField(max_length=50)
    company = models.CharField(max_length=255)
    company_address = models.CharField(max_length=255)
    company_director = models.CharField(max_length=255)
    company_phone = models.CharField(max_length=50)
    practice_supervisor = models.CharField(max_length=255)
    faculty = models.CharField(max_length=150)

    def __str__(self):
        return self.full_name


class Payment(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    amount = models.DecimalField(max_digits=10, decimal_places=2)
    merchant_trans_id = models.CharField(max_length=255)
    status = models.CharField(max_length=50, default='pending')
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"To‘lov {self.user.username} - {self.amount} so‘m ({self.status})"
