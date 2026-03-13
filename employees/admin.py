from django.contrib import admin
from .models import Department, Position, Employee, AttendanceRecord, LeaveRequest, SalaryRecord


@admin.register(Department)
class DepartmentAdmin(admin.ModelAdmin):
    list_display = ['code', 'name', 'created_at']
    search_fields = ['name', 'code']


@admin.register(Position)
class PositionAdmin(admin.ModelAdmin):
    list_display = ['code', 'name', 'department', 'created_at']
    search_fields = ['name', 'code']
    list_filter = ['department']


@admin.register(Employee)
class EmployeeAdmin(admin.ModelAdmin):
    list_display = ['employee_id', 'full_name', 'department', 'position', 'hire_date', 'status']
    search_fields = ['employee_id', 'full_name', 'id_number', 'phone']
    list_filter = ['department', 'status', 'gender']


@admin.register(AttendanceRecord)
class AttendanceRecordAdmin(admin.ModelAdmin):
    list_display = ['employee', 'date', 'check_in', 'check_out', 'status']
    search_fields = ['employee__full_name', 'employee__employee_id']
    list_filter = ['status', 'date']


@admin.register(LeaveRequest)
class LeaveRequestAdmin(admin.ModelAdmin):
    list_display = ['employee', 'leave_type', 'start_date', 'end_date', 'days', 'status']
    search_fields = ['employee__full_name']
    list_filter = ['status', 'leave_type']


@admin.register(SalaryRecord)
class SalaryRecordAdmin(admin.ModelAdmin):
    list_display = ['employee', 'month', 'year', 'basic_salary', 'net_salary']
    search_fields = ['employee__full_name']
    list_filter = ['month', 'year']
