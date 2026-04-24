from django.contrib import admin

from .models import PlanningDispatchRun, PlanningJob, PlanningPrintRun, PoDocument, SkuRecipe


class PlanningPrintRunInline(admin.TabularInline):
    model = PlanningPrintRun
    extra = 0


class PlanningDispatchRunInline(admin.TabularInline):
    model = PlanningDispatchRun
    extra = 0


@admin.register(PlanningJob)
class PlanningJobAdmin(admin.ModelAdmin):
    list_display = (
        'jc_number',
        'po_number',
        'sku',
        'order_qty',
        'machine_name',
        'department',
        'status',
        'updated_at',
    )
    search_fields = ('jc_number', 'po_number', 'sku', 'job_name')
    list_filter = ('status', 'plan_month', 'department', 'machine_name')
    inlines = [PlanningPrintRunInline, PlanningDispatchRunInline]


@admin.register(PoDocument)
class PoDocumentAdmin(admin.ModelAdmin):
    list_display = ('id', 'planning_job', 'extraction_status', 'uploaded_by', 'created_at')
    list_filter = ('extraction_status', 'created_at')
    search_fields = ('planning_job__jc_number', 'planning_job__po_number')


@admin.register(SkuRecipe)
class SkuRecipeAdmin(admin.ModelAdmin):
    list_display = ('sku', 'job_name', 'material', 'machine_name', 'updated_at')
    search_fields = ('sku', 'job_name', 'material', 'machine_name')
    list_filter = ('machine_name', 'updated_at')
