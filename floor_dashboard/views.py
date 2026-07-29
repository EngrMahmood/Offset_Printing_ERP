from django.http import JsonResponse
from django.shortcuts import render

from .services import get_dashboard_data


def dashboard_view(request):
    data = get_dashboard_data()
    return render(request, 'floor_dashboard/dashboard.html', {'data': data})


def dashboard_data_api(request):
    return JsonResponse(get_dashboard_data())
