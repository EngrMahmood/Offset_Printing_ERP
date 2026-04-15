from django.shortcuts import render
from django.http import HttpResponse
import csv
import io

from .bulk_upload import process_jobcard_upload, get_template_headers, get_template_example

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False


def home(request):
    return render(request, 'home.html')


def bulk_upload_jobcards(request):
    if request.method == "POST":
        file = request.FILES['file']

        result = process_jobcard_upload(file)

        return render(request, "upload_result.html", result)

    return render(request, "upload.html")


def download_template(request):
    """Download template in CSV or Excel format"""
    file_format = request.GET.get('format', 'csv').lower()
    
    headers = get_template_headers()
    example = get_template_example()
    
    if file_format == 'excel' and EXCEL_AVAILABLE:
        # Generate Excel file
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.title = "Job Cards"
        
        # Add headers with styling
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        
        for col_num, header in enumerate(headers, 1):
            cell = worksheet.cell(row=1, column=col_num, value=header)
            cell.fill = header_fill
            cell.font = header_font
        
        # Add example row
        for col_num, value in enumerate(example, 1):
            worksheet.cell(row=2, column=col_num, value=value)
        
        # Auto-adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
        
        # Send file
        response = HttpResponse(
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = 'attachment; filename="jobcard_template.xlsx"'
        workbook.save(response)
        return response
    
    else:
        # Generate CSV file (default)
        response = HttpResponse(content_type='text/csv')
        response['Content-Disposition'] = 'attachment; filename="jobcard_template.csv"'
        
        writer = csv.writer(response)
        writer.writerow(headers)
        writer.writerow(example)
        
        return response