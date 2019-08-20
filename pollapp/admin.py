from django.contrib import admin

# Register your models here.
from django.contrib import admin
from .models import Question

# 在admin中注册投票应用
admin.site.register(Question)