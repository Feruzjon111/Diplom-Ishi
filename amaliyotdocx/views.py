from django.shortcuts import redirect, render

def home_redirect(request):
    return render(request, 'app_excel/home.html')
