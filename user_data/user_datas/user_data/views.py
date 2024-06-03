from django.http import HttpResponse
from django.shortcuts import render, redirect
from .models import Detail
import pandas as pd
import openpyxl


# Create your views here.
def contact(request):
    if request.method=='POST':
        name=request.POST.get('name')
        website = request.POST.get('website')
        email_id = request.POST.get('email')
        email_id_2 = request.POST.get('email_2')
        tel_no = request.POST.get('tel')
        address = request.POST.get('address')
        address2 = request.POST.get('address_2')
        city = request.POST.get('city')
        town = request.POST.get('town')
        country = request.POST.get('country')
        technology = request.POST.get('technology')
        technology_2 = request.POST.get('technology_2')
        technology_3 = request.POST.get('technology_3')
        contact_person = request.POST.get('contact_person')

        contact_details=Detail(name=name,website=website,email_id=email_id,email_id_2=email_id_2,tel_no=tel_no,address=address,
                                       address2=address2,city=city,town=town,country=country,technology=technology,technology_2=technology_2,
                                       technology_3=technology_3,contact_person=contact_person)
        contact_details.save()
        return redirect('thank_you')

    return render(request,'contact.html')

def data_list(request):
    all_data=Detail.objects.all()
    context={
        'all_data':all_data
    }
    return render(request,'data_list.html',context)

def filter_by_country(request):
    country = request.GET.get('country', None)
    if country:
        contacts = Detail.objects.filter(country=country)
    else:
        contacts = Detail.objects.all()
    return render(request, '/filter_by_country.html', {'contacts': contacts, 'country': country})

# def contact_list_view(request):
#     contacts = Detail.objects.all()
#     query = request.GET.get('q')
#     if query:
#         contacts = contacts.filter(name__icontains=query)
#     return render(request, 'myapp/contact_list.html', {'contacts': contacts})

def thank_you_view(request):

    return render(request, 'thank.html')

# def search_by_country(request):
#     query = request.GET.get('q', '')
#     if query:
#         contacts = Detail.objects.filter(country__icontains=query)
#     else:
#         contacts = Detail.objects.all()
#     return render(request, 'search.html', {'contacts': contacts, 'query': query})
#
# def download_excel(request):
#     # Create a workbook and a worksheet
#     wb = openpyxl.Workbook()
#     ws = wb.active
#     ws.title = "Data"
#
#     # Add headers to the worksheet
#     ws.append(['id','name','website','email_id','email_id_2','tel_no','address','address2','city','town','country',
#                     'technology','technology_2','contact_person'])
#
#     # Retrieve data from the database
#
#     persons = Detail.objects.all()
#
#     # Add data to the worksheet
#     for person in persons:
#         ws.append([person.id, person.name,person.website, person.email_id,person.email_id_2,person.tel_no,person.address,
#                    person.address2,person.city,person.town,person.country,person.technology,person.technology_2,
#                    person.contact_person])
#
#     # Save the workbook to a bytes buffer
#     response = HttpResponse(
#         content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
#     )
#     response['Content-Disposition'] = 'attachment; filename="data.xlsx"'
#     wb.save(response)
#     return response

# def download_excel(request):
#     query = request.GET.get('q')
#     contacts = Detail.objects.all()
#     if query:
#         contacts = contacts.filter(name__icontains=query)
#
#     data = list(contacts.values())
#     df = pd.DataFrame(data)
#     response = HttpResponse(content_type='application/ms-excel')
#     response['Content-Disposition'] = 'attachment; filename="contacts.xls"'
#     df.to_excel(response, index=False)
#     return response


def search_by_country(request):
    query = request.GET.get('q', '')
    if query:
        contacts = Detail.objects.filter(country__icontains=query)
    else:
        contacts = Detail.objects.all()

    if 'download' in request.GET:
        # Create a workbook and a worksheet
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"

        # Add headers to the worksheet
        ws.append(['id', 'name', 'website', 'email_id', 'email_id_2', 'tel_no', 'address', 'address2', 'city', 'town', 'country',
                   'technology', 'technology_2', 'contact_person'])

        # Add data to the worksheet
        for person in contacts:
            ws.append([person.id, person.name, person.website, person.email_id, person.email_id_2, person.tel_no, person.address,
                       person.address2, person.city, person.town, person.country, person.technology, person.technology_2,
                       person.contact_person])

        # Save the workbook to a bytes buffer
        response = HttpResponse(
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = 'attachment; filename="data.xlsx"'
        wb.save(response)
        return response

    return render(request, 'search.html', {'contacts': contacts, 'query': query})