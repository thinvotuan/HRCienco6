from django.test import TestCase, Client
from django.urls import reverse
from django.utils import timezone
import datetime

from .models import Department, Position, Employee, AttendanceRecord, LeaveRequest, SalaryRecord


class DepartmentModelTest(TestCase):
    def setUp(self):
        self.dept = Department.objects.create(name='Phòng Kỹ thuật', code='KT')

    def test_department_creation(self):
        self.assertEqual(self.dept.name, 'Phòng Kỹ thuật')
        self.assertEqual(self.dept.code, 'KT')
        self.assertEqual(str(self.dept), 'Phòng Kỹ thuật')

    def test_department_unique_code(self):
        from django.db import IntegrityError
        with self.assertRaises(Exception):
            Department.objects.create(name='Test', code='KT')


class PositionModelTest(TestCase):
    def setUp(self):
        self.dept = Department.objects.create(name='Phòng Kỹ thuật', code='KT')
        self.pos = Position.objects.create(name='Kỹ sư', code='KS', department=self.dept)

    def test_position_creation(self):
        self.assertEqual(self.pos.name, 'Kỹ sư')
        self.assertEqual(self.pos.department, self.dept)
        self.assertEqual(str(self.pos), 'Kỹ sư')


class EmployeeModelTest(TestCase):
    def setUp(self):
        self.dept = Department.objects.create(name='Phòng Kỹ thuật', code='KT')
        self.employee = Employee.objects.create(
            employee_id='NV001',
            full_name='Nguyễn Văn An',
            date_of_birth='1990-01-01',
            gender='male',
            id_number='012345678901',
            phone='0901234567',
            hire_date='2020-01-01',
            status='active',
            basic_salary=10000000,
            department=self.dept,
        )

    def test_employee_creation(self):
        self.assertEqual(self.employee.full_name, 'Nguyễn Văn An')
        self.assertEqual(self.employee.status, 'active')
        self.assertIn('NV001', str(self.employee))

    def test_employee_status_badge(self):
        self.assertEqual(self.employee.get_status_display_badge(), 'success')
        self.employee.status = 'inactive'
        self.assertEqual(self.employee.get_status_display_badge(), 'danger')
        self.employee.status = 'probation'
        self.assertEqual(self.employee.get_status_display_badge(), 'warning')


class SalaryRecordModelTest(TestCase):
    def setUp(self):
        self.dept = Department.objects.create(name='Phòng Test', code='TEST')
        self.employee = Employee.objects.create(
            employee_id='NV999',
            full_name='Test User',
            date_of_birth='1990-01-01',
            gender='male',
            id_number='999999999999',
            phone='0999999999',
            hire_date='2020-01-01',
        )

    def test_salary_net_calculation(self):
        sal = SalaryRecord.objects.create(
            employee=self.employee,
            month=1,
            year=2025,
            basic_salary=10000000,
            allowance=500000,
            bonus=200000,
            deduction=100000,
            working_days=26,
            actual_days=26,
        )
        expected = (10000000 / 26 * 26) + 500000 + 200000 - 100000
        self.assertAlmostEqual(float(sal.net_salary), expected, places=1)


class DashboardViewTest(TestCase):
    def test_dashboard_status_code(self):
        c = Client()
        response = c.get(reverse('dashboard'))
        self.assertEqual(response.status_code, 200)

    def test_dashboard_template(self):
        c = Client()
        response = c.get(reverse('dashboard'))
        self.assertTemplateUsed(response, 'dashboard.html')


class EmployeeListViewTest(TestCase):
    def setUp(self):
        self.client = Client()
        self.dept = Department.objects.create(name='Phòng Test', code='TST')
        Employee.objects.create(
            employee_id='NV001', full_name='Nguyễn Test',
            date_of_birth='1990-01-01', gender='male',
            id_number='111111111111', phone='0900000001',
            hire_date='2020-01-01', status='active',
        )

    def test_employee_list_status_code(self):
        response = self.client.get(reverse('employee_list'))
        self.assertEqual(response.status_code, 200)

    def test_employee_list_template(self):
        response = self.client.get(reverse('employee_list'))
        self.assertTemplateUsed(response, 'employees/employee_list.html')

    def test_employee_list_search(self):
        response = self.client.get(reverse('employee_list'), {'q': 'Nguyễn'})
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Nguyễn Test')


class EmployeeCRUDTest(TestCase):
    def setUp(self):
        self.client = Client()
        self.dept = Department.objects.create(name='Phòng Test', code='TST')
        self.employee = Employee.objects.create(
            employee_id='NV001', full_name='Nguyễn Test',
            date_of_birth='1990-01-01', gender='male',
            id_number='111111111111', phone='0900000001',
            hire_date='2020-01-01', status='active',
        )

    def test_employee_detail_view(self):
        response = self.client.get(reverse('employee_detail', args=[self.employee.pk]))
        self.assertEqual(response.status_code, 200)

    def test_employee_create_view_get(self):
        response = self.client.get(reverse('employee_add'))
        self.assertEqual(response.status_code, 200)

    def test_employee_create_view_post(self):
        response = self.client.post(reverse('employee_add'), {
            'employee_id': 'NV002',
            'full_name': 'Trần Thị Test',
            'date_of_birth': '1995-05-05',
            'gender': 'female',
            'id_number': '222222222222',
            'phone': '0900000002',
            'hire_date': '2022-01-01',
            'status': 'active',
            'basic_salary': '10000000',
        })
        self.assertIn(response.status_code, [200, 302])

    def test_employee_update_view(self):
        response = self.client.get(reverse('employee_edit', args=[self.employee.pk]))
        self.assertEqual(response.status_code, 200)

    def test_employee_delete_view(self):
        response = self.client.get(reverse('employee_delete', args=[self.employee.pk]))
        self.assertEqual(response.status_code, 200)


class DepartmentViewTest(TestCase):
    def setUp(self):
        self.client = Client()
        self.dept = Department.objects.create(name='Phòng Test', code='TST')

    def test_department_list(self):
        response = self.client.get(reverse('department_list'))
        self.assertEqual(response.status_code, 200)

    def test_department_create_get(self):
        response = self.client.get(reverse('department_add'))
        self.assertEqual(response.status_code, 200)
