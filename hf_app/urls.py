from django.urls import path
from hf_app import views

urlpatterns = [
    path('', views.index, name='index'),
    path('upload/', views.upload, name='upload'),
    path('history/', views.history, name='history'),
    path('status/<int:record_id>/', views.status, name='status'),
    path('download/<int:record_id>/', views.download, name='download'),
]