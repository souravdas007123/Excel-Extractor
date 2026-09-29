from django.urls import path
from . import views

urlpatterns = [
    # Landing Page
    path('', views.landing_page, name='home'),
    
    # Main Tool (Aapka Excel Dashboard)
    path('app/', views.excel_dashboard, name='app'),
    
    # Authentication Pages
    path('login/', views.user_login, name='login'),
    path('register/', views.register, name='register'),
    path('logout/', views.user_logout, name='logout'),
]