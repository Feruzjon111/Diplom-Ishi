from django.shortcuts import redirect, render

def home_redirect(request):
    if request.user.is_authenticated:
        return redirect('dashboard')
    return render(request, 'app_excel/home.html')
