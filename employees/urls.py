from django.urls import path
from . import views

urlpatterns = [
    path('', views.dashboard, name='dashboard'),

    path('employees/', views.EmployeeListView.as_view(), name='employee_list'),
    path('employees/add/', views.EmployeeCreateView.as_view(), name='employee_add'),
    path('employees/<int:pk>/', views.EmployeeDetailView.as_view(), name='employee_detail'),
    path('employees/<int:pk>/edit/', views.EmployeeUpdateView.as_view(), name='employee_edit'),
    path('employees/<int:pk>/delete/', views.EmployeeDeleteView.as_view(), name='employee_delete'),

    path('departments/', views.DepartmentListView.as_view(), name='department_list'),
    path('departments/add/', views.DepartmentCreateView.as_view(), name='department_add'),
    path('departments/<int:pk>/edit/', views.DepartmentUpdateView.as_view(), name='department_edit'),
    path('departments/<int:pk>/delete/', views.DepartmentDeleteView.as_view(), name='department_delete'),

    path('positions/', views.PositionListView.as_view(), name='position_list'),
    path('positions/add/', views.PositionCreateView.as_view(), name='position_add'),
    path('positions/<int:pk>/edit/', views.PositionUpdateView.as_view(), name='position_edit'),
    path('positions/<int:pk>/delete/', views.PositionDeleteView.as_view(), name='position_delete'),

    path('attendance/', views.AttendanceListView.as_view(), name='attendance_list'),
    path('attendance/add/', views.AttendanceCreateView.as_view(), name='attendance_add'),
    path('attendance/<int:pk>/edit/', views.AttendanceUpdateView.as_view(), name='attendance_edit'),
    path('attendance/<int:pk>/delete/', views.AttendanceDeleteView.as_view(), name='attendance_delete'),

    path('leaves/', views.LeaveListView.as_view(), name='leave_list'),
    path('leaves/add/', views.LeaveCreateView.as_view(), name='leave_add'),
    path('leaves/<int:pk>/edit/', views.LeaveUpdateView.as_view(), name='leave_edit'),
    path('leaves/<int:pk>/delete/', views.LeaveDeleteView.as_view(), name='leave_delete'),

    path('salary/', views.SalaryListView.as_view(), name='salary_list'),
    path('salary/add/', views.SalaryCreateView.as_view(), name='salary_add'),
    path('salary/<int:pk>/edit/', views.SalaryUpdateView.as_view(), name='salary_edit'),
    path('salary/<int:pk>/delete/', views.SalaryDeleteView.as_view(), name='salary_delete'),
]
