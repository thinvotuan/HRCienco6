from django.shortcuts import render, get_object_or_404, redirect
from django.views.generic import ListView, DetailView, CreateView, UpdateView, DeleteView
from django.urls import reverse_lazy
from django.contrib import messages
from django.db.models import Q, Count, Sum
from django.utils import timezone
import datetime

from .models import Department, Position, Employee, AttendanceRecord, LeaveRequest, SalaryRecord
from .forms import (DepartmentForm, PositionForm, EmployeeForm,
                    AttendanceForm, LeaveRequestForm, SalaryForm)


def dashboard(request):
    total_employees = Employee.objects.count()
    active_employees = Employee.objects.filter(status='active').count()
    probation_employees = Employee.objects.filter(status='probation').count()
    inactive_employees = Employee.objects.filter(status='inactive').count()

    departments = Department.objects.annotate(emp_count=Count('employees')).order_by('name')

    today = timezone.now().date()
    week_ago = today - datetime.timedelta(days=7)
    recent_attendance = AttendanceRecord.objects.filter(
        date__gte=week_ago
    ).select_related('employee').order_by('-date')[:20]

    pending_leaves = LeaveRequest.objects.filter(status='pending').count()

    now = timezone.now()
    monthly_salary = SalaryRecord.objects.filter(
        month=now.month, year=now.year
    ).aggregate(total=Sum('net_salary'))['total'] or 0

    context = {
        'total_employees': total_employees,
        'active_employees': active_employees,
        'probation_employees': probation_employees,
        'inactive_employees': inactive_employees,
        'departments': departments,
        'recent_attendance': recent_attendance,
        'pending_leaves': pending_leaves,
        'monthly_salary': monthly_salary,
        'current_month': now.month,
        'current_year': now.year,
    }
    return render(request, 'dashboard.html', context)


# --- Department Views ---

class DepartmentListView(ListView):
    model = Department
    template_name = 'employees/department_list.html'
    context_object_name = 'departments'
    paginate_by = 10

    def get_queryset(self):
        return Department.objects.annotate(emp_count=Count('employees')).order_by('name')


class DepartmentCreateView(CreateView):
    model = Department
    form_class = DepartmentForm
    template_name = 'employees/department_form.html'
    success_url = reverse_lazy('department_list')

    def form_valid(self, form):
        messages.success(self.request, 'Thêm phòng ban thành công!')
        return super().form_valid(form)

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['title'] = 'Thêm phòng ban'
        return ctx


class DepartmentUpdateView(UpdateView):
    model = Department
    form_class = DepartmentForm
    template_name = 'employees/department_form.html'
    success_url = reverse_lazy('department_list')

    def form_valid(self, form):
        messages.success(self.request, 'Cập nhật phòng ban thành công!')
        return super().form_valid(form)

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['title'] = 'Cập nhật phòng ban'
        return ctx


class DepartmentDeleteView(DeleteView):
    model = Department
    template_name = 'employees/department_confirm_delete.html'
    success_url = reverse_lazy('department_list')

    def form_valid(self, form):
        messages.success(self.request, 'Xóa phòng ban thành công!')
        return super().form_valid(form)


# --- Position Views ---

class PositionListView(ListView):
    model = Position
    template_name = 'employees/position_list.html'
    context_object_name = 'positions'
    paginate_by = 10

    def get_queryset(self):
        return Position.objects.select_related('department').order_by('name')


class PositionCreateView(CreateView):
    model = Position
    form_class = PositionForm
    template_name = 'employees/position_form.html'
    success_url = reverse_lazy('position_list')

    def form_valid(self, form):
        messages.success(self.request, 'Thêm chức vụ thành công!')
        return super().form_valid(form)

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['title'] = 'Thêm chức vụ'
        return ctx


class PositionUpdateView(UpdateView):
    model = Position
    form_class = PositionForm
    template_name = 'employees/position_form.html'
    success_url = reverse_lazy('position_list')

    def form_valid(self, form):
        messages.success(self.request, 'Cập nhật chức vụ thành công!')
        return super().form_valid(form)

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['title'] = 'Cập nhật chức vụ'
        return ctx


class PositionDeleteView(DeleteView):
    model = Position
    template_name = 'employees/position_confirm_delete.html'
    success_url = reverse_lazy('position_list')

    def form_valid(self, form):
        messages.success(self.request, 'Xóa chức vụ thành công!')
        return super().form_valid(form)


# --- Employee Views ---

class EmployeeListView(ListView):
    model = Employee
    template_name = 'employees/employee_list.html'
    context_object_name = 'employees'
    paginate_by = 10

    def get_queryset(self):
        qs = Employee.objects.select_related('department', 'position').order_by('employee_id')
        q = self.request.GET.get('q', '')
        dept = self.request.GET.get('department', '')
        status = self.request.GET.get('status', '')
        if q:
            qs = qs.filter(Q(full_name__icontains=q) | Q(employee_id__icontains=q))
        if dept:
            qs = qs.filter(department__id=dept)
        if status:
            qs = qs.filter(status=status)
        return qs

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['departments'] = Department.objects.all()
        ctx['status_choices'] = Employee.STATUS_CHOICES
        ctx['q'] = self.request.GET.get('q', '')
        ctx['selected_dept'] = self.request.GET.get('department', '')
        ctx['selected_status'] = self.request.GET.get('status', '')
        return ctx


class EmployeeDetailView(DetailView):
    model = Employee
    template_name = 'employees/employee_detail.html'
    context_object_name = 'employee'

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        emp = self.get_object()
        ctx['recent_attendance'] = emp.attendance_records.order_by('-date')[:10]
        ctx['leave_requests'] = emp.leave_requests.order_by('-created_at')[:5]
        ctx['salary_records'] = emp.salary_records.order_by('-year', '-month')[:5]
        return ctx


class EmployeeCreateView(CreateView):
    model = Employee
    form_class = EmployeeForm
    template_name = 'employees/employee_form.html'
    success_url = reverse_lazy('employee_list')

    def form_valid(self, form):
        messages.success(self.request, 'Thêm nhân viên thành công!')
        return super().form_valid(form)

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['title'] = 'Thêm nhân viên mới'
        return ctx


class EmployeeUpdateView(UpdateView):
    model = Employee
    form_class = EmployeeForm
    template_name = 'employees/employee_form.html'
    success_url = reverse_lazy('employee_list')

    def form_valid(self, form):
        messages.success(self.request, 'Cập nhật nhân viên thành công!')
        return super().form_valid(form)

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['title'] = 'Cập nhật nhân viên'
        return ctx


class EmployeeDeleteView(DeleteView):
    model = Employee
    template_name = 'employees/employee_confirm_delete.html'
    success_url = reverse_lazy('employee_list')

    def form_valid(self, form):
        messages.success(self.request, 'Xóa nhân viên thành công!')
        return super().form_valid(form)


# --- Attendance Views ---

class AttendanceListView(ListView):
    model = AttendanceRecord
    template_name = 'employees/attendance_list.html'
    context_object_name = 'records'
    paginate_by = 20

    def get_queryset(self):
        qs = AttendanceRecord.objects.select_related('employee').order_by('-date')
        emp_id = self.request.GET.get('employee', '')
        date_from = self.request.GET.get('date_from', '')
        date_to = self.request.GET.get('date_to', '')
        if emp_id:
            qs = qs.filter(employee__id=emp_id)
        if date_from:
            qs = qs.filter(date__gte=date_from)
        if date_to:
            qs = qs.filter(date__lte=date_to)
        return qs

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['employees'] = Employee.objects.filter(status__in=['active', 'probation'])
        ctx['selected_emp'] = self.request.GET.get('employee', '')
        ctx['date_from'] = self.request.GET.get('date_from', '')
        ctx['date_to'] = self.request.GET.get('date_to', '')
        return ctx


class AttendanceCreateView(CreateView):
    model = AttendanceRecord
    form_class = AttendanceForm
    template_name = 'employees/attendance_form.html'
    success_url = reverse_lazy('attendance_list')

    def form_valid(self, form):
        messages.success(self.request, 'Thêm chấm công thành công!')
        return super().form_valid(form)

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['title'] = 'Thêm chấm công'
        return ctx


class AttendanceUpdateView(UpdateView):
    model = AttendanceRecord
    form_class = AttendanceForm
    template_name = 'employees/attendance_form.html'
    success_url = reverse_lazy('attendance_list')

    def form_valid(self, form):
        messages.success(self.request, 'Cập nhật chấm công thành công!')
        return super().form_valid(form)

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['title'] = 'Cập nhật chấm công'
        return ctx


class AttendanceDeleteView(DeleteView):
    model = AttendanceRecord
    template_name = 'employees/attendance_confirm_delete.html'
    success_url = reverse_lazy('attendance_list')

    def form_valid(self, form):
        messages.success(self.request, 'Xóa chấm công thành công!')
        return super().form_valid(form)


# --- Leave Views ---

class LeaveListView(ListView):
    model = LeaveRequest
    template_name = 'employees/leave_list.html'
    context_object_name = 'leaves'
    paginate_by = 15

    def get_queryset(self):
        qs = LeaveRequest.objects.select_related('employee').order_by('-created_at')
        status = self.request.GET.get('status', '')
        if status:
            qs = qs.filter(status=status)
        return qs

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['status_choices'] = LeaveRequest.STATUS_CHOICES
        ctx['selected_status'] = self.request.GET.get('status', '')
        return ctx


class LeaveCreateView(CreateView):
    model = LeaveRequest
    form_class = LeaveRequestForm
    template_name = 'employees/leave_form.html'
    success_url = reverse_lazy('leave_list')

    def form_valid(self, form):
        messages.success(self.request, 'Tạo đơn nghỉ phép thành công!')
        return super().form_valid(form)

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['title'] = 'Tạo đơn xin nghỉ phép'
        return ctx


class LeaveUpdateView(UpdateView):
    model = LeaveRequest
    form_class = LeaveRequestForm
    template_name = 'employees/leave_form.html'
    success_url = reverse_lazy('leave_list')

    def form_valid(self, form):
        messages.success(self.request, 'Cập nhật đơn nghỉ phép thành công!')
        return super().form_valid(form)

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['title'] = 'Cập nhật đơn nghỉ phép'
        return ctx


class LeaveDeleteView(DeleteView):
    model = LeaveRequest
    template_name = 'employees/leave_confirm_delete.html'
    success_url = reverse_lazy('leave_list')

    def form_valid(self, form):
        messages.success(self.request, 'Xóa đơn nghỉ phép thành công!')
        return super().form_valid(form)


# --- Salary Views ---

class SalaryListView(ListView):
    model = SalaryRecord
    template_name = 'employees/salary_list.html'
    context_object_name = 'salaries'
    paginate_by = 15

    def get_queryset(self):
        qs = SalaryRecord.objects.select_related('employee').order_by('-year', '-month')
        month = self.request.GET.get('month', '')
        year = self.request.GET.get('year', '')
        if month:
            qs = qs.filter(month=month)
        if year:
            qs = qs.filter(year=year)
        return qs

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['months'] = range(1, 13)
        now = timezone.now()
        ctx['years'] = range(now.year - 3, now.year + 2)
        ctx['selected_month'] = self.request.GET.get('month', '')
        ctx['selected_year'] = self.request.GET.get('year', '')
        return ctx


class SalaryCreateView(CreateView):
    model = SalaryRecord
    form_class = SalaryForm
    template_name = 'employees/salary_form.html'
    success_url = reverse_lazy('salary_list')

    def form_valid(self, form):
        messages.success(self.request, 'Thêm bảng lương thành công!')
        return super().form_valid(form)

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['title'] = 'Thêm bảng lương'
        return ctx


class SalaryUpdateView(UpdateView):
    model = SalaryRecord
    form_class = SalaryForm
    template_name = 'employees/salary_form.html'
    success_url = reverse_lazy('salary_list')

    def form_valid(self, form):
        messages.success(self.request, 'Cập nhật bảng lương thành công!')
        return super().form_valid(form)

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx['title'] = 'Cập nhật bảng lương'
        return ctx


class SalaryDeleteView(DeleteView):
    model = SalaryRecord
    template_name = 'employees/salary_confirm_delete.html'
    success_url = reverse_lazy('salary_list')

    def form_valid(self, form):
        messages.success(self.request, 'Xóa bảng lương thành công!')
        return super().form_valid(form)
