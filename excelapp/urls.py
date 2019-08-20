# !/usr/bin/python3
# -*- coding: utf-8 -*-

"""
@Author: 闫继龙
@Version: ??
@License: Apache Licence
@CreateTime: 2019/8/20 16:26
@Describe：

"""
from django.urls import path

from . import views

app_name = 'excelapp'

urlpatterns = [
    path('', views.index, name='index')
]
