from django import forms
from .models import Department, Position, Employee, AttendanceRecord, LeaveRequest, SalaryRecord


class DepartmentForm(forms.ModelForm):
    class Meta:
        model = Department
        fields = ['name', 'code', 'description']
        widgets = {
            'name': forms.TextInput(attrs={'class': 'form-control'}),
            'code': forms.TextInput(attrs={'class': 'form-control'}),
            'description': forms.Textarea(attrs={'class': 'form-control', 'rows': 3}),
        }


class PositionForm(forms.ModelForm):
    class Meta:
        model = Position
        fields = ['name', 'code', 'department', 'description']
        widgets = {
            'name': forms.TextInput(attrs={'class': 'form-control'}),
            'code': forms.TextInput(attrs={'class': 'form-control'}),
            'department': forms.Select(attrs={'class': 'form-select'}),
            'description': forms.Textarea(attrs={'class': 'form-control', 'rows': 3}),
        }


class EmployeeForm(forms.ModelForm):
    class Meta:
        model = Employee
        fields = [
            'employee_id', 'full_name', 'date_of_birth', 'gender', 'id_number',
            'phone', 'email', 'address', 'department', 'position',
            'hire_date', 'status', 'basic_salary', 'photo',
        ]
        widgets = {
            'employee_id': forms.TextInput(attrs={'class': 'form-control'}),
            'full_name': forms.TextInput(attrs={'class': 'form-control'}),
            'date_of_birth': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
            'gender': forms.Select(attrs={'class': 'form-select'}),
            'id_number': forms.TextInput(attrs={'class': 'form-control'}),
            'phone': forms.TextInput(attrs={'class': 'form-control'}),
            'email': forms.EmailInput(attrs={'class': 'form-control'}),
            'address': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),
            'department': forms.Select(attrs={'class': 'form-select'}),
            'position': forms.Select(attrs={'class': 'form-select'}),
            'hire_date': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
            'status': forms.Select(attrs={'class': 'form-select'}),
            'basic_salary': forms.NumberInput(attrs={'class': 'form-control'}),
            'photo': forms.FileInput(attrs={'class': 'form-control'}),
        }


class AttendanceForm(forms.ModelForm):
    class Meta:
        model = AttendanceRecord
        fields = ['employee', 'date', 'check_in', 'check_out', 'status', 'note']
        widgets = {
            'employee': forms.Select(attrs={'class': 'form-select'}),
            'date': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
            'check_in': forms.TimeInput(attrs={'class': 'form-control', 'type': 'time'}),
            'check_out': forms.TimeInput(attrs={'class': 'form-control', 'type': 'time'}),
            'status': forms.Select(attrs={'class': 'form-select'}),
            'note': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),
        }


class LeaveRequestForm(forms.ModelForm):
    class Meta:
        model = LeaveRequest
        fields = ['employee', 'leave_type', 'start_date', 'end_date', 'days', 'reason', 'status', 'approved_by']
        widgets = {
            'employee': forms.Select(attrs={'class': 'form-select'}),
            'leave_type': forms.Select(attrs={'class': 'form-select'}),
            'start_date': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
            'end_date': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
            'days': forms.NumberInput(attrs={'class': 'form-control'}),
            'reason': forms.Textarea(attrs={'class': 'form-control', 'rows': 3}),
            'status': forms.Select(attrs={'class': 'form-select'}),
            'approved_by': forms.TextInput(attrs={'class': 'form-control'}),
        }


class SalaryForm(forms.ModelForm):
    class Meta:
        model = SalaryRecord
        fields = [
            'employee', 'month', 'year', 'basic_salary',
            'allowance', 'bonus', 'deduction',
            'working_days', 'actual_days', 'note',
        ]
        widgets = {
            'employee': forms.Select(attrs={'class': 'form-select'}),
            'month': forms.NumberInput(attrs={'class': 'form-control', 'min': 1, 'max': 12}),
            'year': forms.NumberInput(attrs={'class': 'form-control', 'min': 2000}),
            'basic_salary': forms.NumberInput(attrs={'class': 'form-control'}),
            'allowance': forms.NumberInput(attrs={'class': 'form-control'}),
            'bonus': forms.NumberInput(attrs={'class': 'form-control'}),
            'deduction': forms.NumberInput(attrs={'class': 'form-control'}),
            'working_days': forms.NumberInput(attrs={'class': 'form-control'}),
            'actual_days': forms.NumberInput(attrs={'class': 'form-control'}),
            'note': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),
        }
