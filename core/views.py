from django.shortcuts import render
from .bulk_upload import process_jobcard_upload
import csv
from django.http import HttpResponse

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

    # Header row
    writer.writerow([
        'job_card_no',
        'SKU',
        'order_qty',
        'material',
        'colour',
        'machine_name',
        'department'
    ])

    # Sample row
    writer.writerow([
        'JC001',
        'SKU-A',
        '1000',
        'Paper',
        '1',
        'Machine-1',
        'Printing'
    ])

    return response