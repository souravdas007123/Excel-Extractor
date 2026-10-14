from django.urls import path
from . import views

urlpatterns = [
    path('', views.landing_page, name='home'),
    path('app/', views.excel_dashboard, name='app'),
    path('download/<str:token>/', views.download_result, name='download'),
    path('jobs/start/', views.start_job, name='start_job'),
    path('jobs/<str:job_id>/status/', views.job_status, name='job_status'),
    path('jobs/<str:job_id>/result/', views.job_result, name='job_result'),

    path('login/', views.user_login, name='login'),
    path('register/', views.register, name='register'),
    path('logout/', views.user_logout, name='logout'),
    path('upgrade/', views.upgrade, name='upgrade'),
]