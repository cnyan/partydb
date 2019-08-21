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
    path('', views.index, name='index'),
    path('write/excel/', views.write_excel, name='writeExcel'),
    path('read/excel/', views.read_excel, name='readExcel'),
    path('delete/', views.delete_table, name='deleteTable'),
    path('read/work1/', views.read_work1, name='readWork1'),
    path('read/work2/', views.read_work2, name='readWork2'),
    path('read/work3/', views.read_work3, name='readWork3'),
    path('read/work4/', views.read_work4, name='readWork4'),
    path('read/work5/', views.read_work5, name='readWork5'),
    path('read/work6/', views.read_work6, name='readWork6'),
    path('read/work7/', views.read_work7, name='readWork7'),
]
