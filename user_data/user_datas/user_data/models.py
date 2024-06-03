from django.db import models

# Create your models here.
class Detail(models.Model):
    name=models.CharField(max_length=250)
    website=models.CharField(max_length=250)
    email_id=models.EmailField(max_length=250)
    email_id_2=models.EmailField(max_length=250,blank=True)
    tel_no=models.CharField(max_length=15)
    address=models.TextField(max_length=250)
    address2=models.TextField(max_length=250,blank=True)
    city=models.CharField(max_length=25)
    town=models.CharField(max_length=25)
    country=models.CharField(max_length=75)
    technology=models.CharField(max_length=250)
    technology_2=models.CharField(max_length=250,blank=True)
    technology_3=models.CharField(max_length=250,blank=True)
    contact_person=models.CharField(max_length=250)

    class Meta:
        ordering=('name',)
        verbose_name='detail'
        verbose_name_plural='details'

    def __str__(self):
        return '{}'.format(self.name)