# -*- coding: utf-8 -*-
# @Time    : 2021/11/27 9:43
# @Author  : xuzhifeng
# @File    : index.py
from django.http.response import HttpResponse

def index(request):
    return HttpResponse(b'Hello World')