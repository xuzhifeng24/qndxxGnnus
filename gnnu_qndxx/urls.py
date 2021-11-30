# -*- coding: utf-8 -*-
# @Time    : 2021/11/28 16:34
# @Author  : xuzhifeng
# @File    : urls.py
from django.urls import path
from . import views

urlpatterns = [
    path('', views.login),
    path('index/', views.userLogin),
    path('details/', views.userdownload),
]