from django.contrib import admin
from .models import Detail

# Register your models here.
class DetailAdmin(admin.ModelAdmin):
    list_display = ['id','name','website','email_id','email_id_2','tel_no','address','address2','city','town','country',
                    'technology','technology_2','technology_3','contact_person']
    list_editable = ['name','website','email_id','email_id_2','tel_no','address','address2','city','town','country',
                    'technology','technology_2','technology_3','contact_person']
    list_filter = ['city','town','country']

admin.site.register(Detail,DetailAdmin)