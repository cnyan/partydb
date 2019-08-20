# !/usr/bin/python3
# -*- coding: utf-8 -*-

"""
@Author: 闫继龙
@Version: ??
@License: Apache Licence
@CreateTime: 2019/8/20 14:54
@Describe：

"""
from django.urls import path

from . import views

app_name = 'pollapp'
urlpatterns = [
    path('', views.index, name='index'),
    path('<int:question_id>/detail/', views.detail, name='detail'),
    path('<int:question_id>/results/', views.results, name='results'),
    path('<int:question_id>/vote/', views.vote, name='vote'),
]
