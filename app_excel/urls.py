from django.urls import include, path
from . import views
from rest_framework.routers import DefaultRouter
from .views import StudentViewSet
from rest_framework.authtoken.views import obtain_auth_token


router = DefaultRouter()
router.register('students', StudentViewSet, basename='student')

urlpatterns = [
    path('', views.dashboard_view, name='dashboard'),
    path('students/', views.students_archive_view, name='students_archive'),
    path('companies/', views.companies_archive_view, name='companies_archive'),
    path('documents/', views.documents_archive_view, name='documents_archive'),
    path('api/', include(router.urls)),
    path('api/chat/', views.ai_chat_view, name='ai_chat'),
    path('api/token/', obtain_auth_token, name='api_token_auth'),

    path('upload/', views.upload_excel, name='upload_excel'),
    path('upload/sample-excel/', views.download_sample_excel, name='download_sample_excel'),
    path('templates/<str:filename>/', views.download_template_source, name='download_template_source'),
    path('export/', views.export_all_documents_zip, name='export_all_documents_zip'),
    path('generate/<str:company_name>/', views.generate_contract_for_company, name='generate_contract_for_company'),
    path('export/one/', views.export_to_word, name='export_to_word'),
    path('documents/<int:student_id>/<str:document_type>/', views.download_student_document, name='download_student_document'),

    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    path('register/', views.register_view, name='register'),

    path('profile/', views.profile_view, name='profile'),
    path('settings/', views.account_settings_view, name='account_settings'),
]

handler403 = 'app_excel.views.handler403'
handler404 = 'app_excel.views.handler404'
handler500 = 'app_excel.views.handler500'
