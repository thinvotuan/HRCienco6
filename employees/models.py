from django.db import models


class Department(models.Model):
    name = models.CharField(max_length=200, verbose_name='Tên phòng ban')
    code = models.CharField(max_length=50, unique=True, verbose_name='Mã phòng ban')
    description = models.TextField(blank=True, verbose_name='Mô tả')
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = 'Phòng ban'
        verbose_name_plural = 'Phòng ban'
        ordering = ['name']

    def __str__(self):
        return self.name


class Position(models.Model):
    name = models.CharField(max_length=200, verbose_name='Tên chức vụ')
    code = models.CharField(max_length=50, unique=True, verbose_name='Mã chức vụ')
    department = models.ForeignKey(Department, on_delete=models.CASCADE,
                                   related_name='positions', verbose_name='Phòng ban')
    description = models.TextField(blank=True, verbose_name='Mô tả')
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = 'Chức vụ'
        verbose_name_plural = 'Chức vụ'
        ordering = ['name']

    def __str__(self):
        return self.name


class Employee(models.Model):
    GENDER_CHOICES = [
        ('male', 'Nam'),
        ('female', 'Nữ'),
        ('other', 'Khác'),
    ]
    STATUS_CHOICES = [
        ('active', 'Đang làm việc'),
        ('inactive', 'Nghỉ việc'),
        ('probation', 'Thử việc'),
    ]

    employee_id = models.CharField(max_length=20, unique=True, verbose_name='Mã nhân viên')
    full_name = models.CharField(max_length=200, verbose_name='Họ và tên')
    date_of_birth = models.DateField(verbose_name='Ngày sinh')
    gender = models.CharField(max_length=10, choices=GENDER_CHOICES, verbose_name='Giới tính')
    id_number = models.CharField(max_length=20, verbose_name='CCCD/CMND')
    phone = models.CharField(max_length=20, verbose_name='Số điện thoại')
    email = models.EmailField(blank=True, verbose_name='Email')
    address = models.TextField(blank=True, verbose_name='Địa chỉ')
    department = models.ForeignKey(Department, on_delete=models.SET_NULL,
                                   null=True, blank=True,
                                   related_name='employees', verbose_name='Phòng ban')
    position = models.ForeignKey(Position, on_delete=models.SET_NULL,
                                 null=True, blank=True,
                                 related_name='employees', verbose_name='Chức vụ')
    hire_date = models.DateField(verbose_name='Ngày vào làm')
    status = models.CharField(max_length=20, choices=STATUS_CHOICES,
                               default='active', verbose_name='Trạng thái')
    basic_salary = models.DecimalField(max_digits=12, decimal_places=2,
                                       default=0, verbose_name='Lương cơ bản')
    photo = models.ImageField(upload_to='employee_photos/', blank=True,
                              null=True, verbose_name='Ảnh')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = 'Nhân viên'
        verbose_name_plural = 'Nhân viên'
        ordering = ['employee_id']

    def __str__(self):
        return f'{self.employee_id} - {self.full_name}'

    def get_status_display_badge(self):
        badges = {
            'active': 'success',
            'inactive': 'danger',
            'probation': 'warning',
        }
        return badges.get(self.status, 'secondary')


class AttendanceRecord(models.Model):
    STATUS_CHOICES = [
        ('present', 'Có mặt'),
        ('absent', 'Vắng mặt'),
        ('late', 'Đi muộn'),
        ('leave', 'Nghỉ phép'),
        ('holiday', 'Nghỉ lễ'),
    ]

    employee = models.ForeignKey(Employee, on_delete=models.CASCADE,
                                  related_name='attendance_records', verbose_name='Nhân viên')
    date = models.DateField(verbose_name='Ngày')
    check_in = models.TimeField(null=True, blank=True, verbose_name='Giờ vào')
    check_out = models.TimeField(null=True, blank=True, verbose_name='Giờ ra')
    status = models.CharField(max_length=20, choices=STATUS_CHOICES,
                               default='present', verbose_name='Trạng thái')
    note = models.TextField(blank=True, verbose_name='Ghi chú')
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = 'Chấm công'
        verbose_name_plural = 'Chấm công'
        unique_together = [['employee', 'date']]
        ordering = ['-date']

    def __str__(self):
        return f'{self.employee.full_name} - {self.date}'


class LeaveRequest(models.Model):
    LEAVE_TYPE_CHOICES = [
        ('annual', 'Nghỉ phép năm'),
        ('sick', 'Nghỉ ốm'),
        ('personal', 'Nghỉ việc riêng'),
        ('unpaid', 'Nghỉ không lương'),
    ]
    STATUS_CHOICES = [
        ('pending', 'Chờ duyệt'),
        ('approved', 'Đã duyệt'),
        ('rejected', 'Từ chối'),
    ]

    employee = models.ForeignKey(Employee, on_delete=models.CASCADE,
                                  related_name='leave_requests', verbose_name='Nhân viên')
    leave_type = models.CharField(max_length=20, choices=LEAVE_TYPE_CHOICES,
                                   verbose_name='Loại nghỉ phép')
    start_date = models.DateField(verbose_name='Ngày bắt đầu')
    end_date = models.DateField(verbose_name='Ngày kết thúc')
    days = models.IntegerField(verbose_name='Số ngày nghỉ')
    reason = models.TextField(verbose_name='Lý do')
    status = models.CharField(max_length=20, choices=STATUS_CHOICES,
                               default='pending', verbose_name='Trạng thái')
    approved_by = models.CharField(max_length=200, blank=True, verbose_name='Người duyệt')
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = 'Đơn xin nghỉ phép'
        verbose_name_plural = 'Đơn xin nghỉ phép'
        ordering = ['-created_at']

    def __str__(self):
        return f'{self.employee.full_name} - {self.get_leave_type_display()} ({self.start_date})'

    def get_status_badge(self):
        badges = {
            'pending': 'warning',
            'approved': 'success',
            'rejected': 'danger',
        }
        return badges.get(self.status, 'secondary')


class SalaryRecord(models.Model):
    employee = models.ForeignKey(Employee, on_delete=models.CASCADE,
                                  related_name='salary_records', verbose_name='Nhân viên')
    month = models.IntegerField(verbose_name='Tháng')
    year = models.IntegerField(verbose_name='Năm')
    basic_salary = models.DecimalField(max_digits=12, decimal_places=2, verbose_name='Lương cơ bản')
    allowance = models.DecimalField(max_digits=12, decimal_places=2,
                                    default=0, verbose_name='Phụ cấp')
    bonus = models.DecimalField(max_digits=12, decimal_places=2,
                                default=0, verbose_name='Thưởng')
    deduction = models.DecimalField(max_digits=12, decimal_places=2,
                                    default=0, verbose_name='Khấu trừ')
    working_days = models.IntegerField(default=26, verbose_name='Ngày công chuẩn')
    actual_days = models.IntegerField(default=0, verbose_name='Ngày công thực tế')
    net_salary = models.DecimalField(max_digits=12, decimal_places=2,
                                     default=0, verbose_name='Lương thực nhận')
    note = models.TextField(blank=True, verbose_name='Ghi chú')
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = 'Bảng lương'
        verbose_name_plural = 'Bảng lương'
        unique_together = [['employee', 'month', 'year']]
        ordering = ['-year', '-month']

    def __str__(self):
        return f'{self.employee.full_name} - {self.month}/{self.year}'

    def save(self, *args, **kwargs):
        if self.working_days and self.working_days > 0:
            self.net_salary = (
                (self.basic_salary / self.working_days * self.actual_days)
                + self.allowance + self.bonus - self.deduction
            )
        else:
            self.net_salary = self.allowance + self.bonus - self.deduction
        super().save(*args, **kwargs)
