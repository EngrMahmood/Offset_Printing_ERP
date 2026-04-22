from django.conf import settings


def software_version(request):
    return {
        'erp_software_version': getattr(settings, 'ERP_SOFTWARE_VERSION', '0.0.0'),
        'erp_software_release_date': getattr(settings, 'ERP_SOFTWARE_RELEASE_DATE', ''),
    }
