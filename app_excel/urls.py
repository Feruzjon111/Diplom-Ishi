from django.urls import path, include
from . import views
from rest_framework.routers import DefaultRouter
from .views import StudentViewSet, StudentListCreateAPIView, StudentRetrieveUpdateDestroyAPIView
from rest_framework.authtoken.views import obtain_auth_token


router = DefaultRouter()
router.register('students', StudentViewSet, basename='student')

urlpatterns = [
    path('', views.dashboard_view, name='dashboard'),
    path('api/', include(router.urls)),
    path('api/token/', obtain_auth_token, name='api_token_auth'),

    path('upload/', views.upload_excel, name='upload_excel'),
    path('upload/sample-excel/', views.download_sample_excel, name='download_sample_excel'),
    path('templates/<str:filename>/', views.download_template_source, name='download_template_source'),
    path('export/', views.export_all_documents_zip, name='export_all_documents_zip'),
    path('generate/<str:company_name>/', views.generate_contract_for_company, name='generate_contract_for_company'),
    path('export/one/', views.export_to_word, name='export_to_word'),

    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    path('register/', views.register_view, name='register'),

    path('api/', include(router.urls)),
    path('api2/students/', StudentListCreateAPIView.as_view()),
    path('api2/students/<int:pk>/', StudentRetrieveUpdateDestroyAPIView.as_view()),

    path('pay/', views.pay_view, name='pay_view'),
    path('click/result/', views.click_result, name='click_result'),
    path('click/prepare/', views.click_prepare, name='click_prepare'),
    path('payment/callback/', views.payment_callback, name='payment_callback'),

    path('balance/', views.balance_view, name='balance'),
    path('balance/top-up/', views.top_up_balance, name='top_up_balance'),
    path('profile/', views.profile_view, name='profile'),
    path('settings/', views.account_settings_view, name='account_settings'),


]



# handler401 = 'app_excel.views.custom_401_view'
# handler403 = 'app_excel.views.csrf_failure'
# handler404 = 'app_excel.views.custom_404_view'
# handler500 = 'app_excel.views.custom_500_view'


handler403 = 'app_excel.views.handler403'
handler404 = 'app_excel.views.handler404'
handler500 = 'app_excel.views.handler500'
