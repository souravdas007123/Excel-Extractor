from django.urls import path
from . import views

urlpatterns = [
    path('', views.excel_dashboard, name='excel_dashboard'),
    
]