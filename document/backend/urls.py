
from django.contrib import admin
from django.urls import path
from . import views
urlpatterns = [
    path('', views.home),
    path('process', views.process),
    path('download', views.download, name='download-word-document'),
]
