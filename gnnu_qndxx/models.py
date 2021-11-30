from django.db import models

# Create your models here.

class User(models.Model):
    xy_name = models.CharField(max_length=50)
    use_name = models.CharField(max_length=50)
    password = models.CharField(max_length=50)
