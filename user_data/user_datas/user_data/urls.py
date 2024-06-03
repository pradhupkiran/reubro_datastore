from django.urls import path
from . import views

urlpatterns = [

    path('',views.contact,name='contact'),
    path('data_list/',views.data_list,name='data_list'),
    path('thank-you/', views.thank_you_view, name='thank_you'),
    path('search-by-country/', views.search_by_country, name='search_by_country'),
    # path('download_excel/', views.download_excel, name='download_excel'),


]