from django.shortcuts import render
from django.http import HttpResponse
import csv

from .bulk_upload import process_jobcard_upload


def home(request):
    return render(request, 'home.html')


def bulk_upload_jobcards(request):
    if request.method == "POST":
        file = request.FILES['file']

        result = process_jobcard_upload(file)

        return render(request, "upload_result.html", result)

    return render(request, "upload.html")


def download_template(request):
    response = HttpResponse(content_type='text/csv')
    response['Content-Disposition'] = 'attachment; filename="jobcard_template.csv"'

    writer = csv.writer(response)

    headers = [
        'JC Number',
        'SKU',
        'PO Number',
        'PO Date',
        'Month',
        'Material',
        'Colour',
        'Application',
        'Order Quantity',
        'Ups',
        'Print Sheet Size',
        'Wastage (%)',
        'Actual Sheet Required',
        'Purchase Sheet Size',
        'Purchase Sheet Ups',
        'Remarks',
        'Destination',
        'Machine',
        'Department',
        'Die Cutting (Yes/No)'
    ]

    writer.writerow(headers)

    writer.writerow([
        'JC-26-1001',
        'SKU-01',
        'PO-7788',
        '4/13/2026',
        'April',
        'Bleach230',
        '4',
        'UV',
        '10000',
        '12',
        '20x30',
        '5',
        '10500',
        '20x30',
        '6',
        'Urgent job',
        'SITE 1',
        'GTO 1A',
        'Pillow',
        'Yes'
    ])

    return response