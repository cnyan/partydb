from django.db import models

# 在 settings.py 中启用模型
# Create your models here.

class excelUser(models.Model):
    seq = models.IntegerField(default=0)
    name = models.CharField(max_length=200)
    major = models.CharField(max_length=240)
    idcard = models.CharField(max_length=250)

    def __str__(self):
        return '{0} . {1} ,{2}'.format(self.name, self.major, self.idcard)
