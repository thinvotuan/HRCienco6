using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using BatDongSan.Models.HeThong;
using BatDongSan.Models.NhanSu;

namespace BatDongSan.Models.PhieuDeNghi
{
    public class PhieuCapNhatNgayCong
    {
        public string maPhieuCapNhatNgayCong { get; set; }

        public NhanVienModel nguoiLap { get; set; }

        public DateTime ngayLap { get; set; }

        public int maQuiTrinhDuyet { get; set; }

        public int thang { get; set; }

        public int nam { get; set; }

        public DMNguoiDuyet Duyet { get; set; }

        public PhieuCapNhatNgayCongNhanVien chiTiet { get; set; }

        public string ghiChu { get; set; }

        public PhieuCapNhatNgayCong()
        {
            Duyet = new DMNguoiDuyet();
            chiTiet = new PhieuCapNhatNgayCongNhanVien();
        }
    }

    public class PhieuCapNhatNgayCongNhanVien
    {
        public NhanVienModel nhanVien { get; set; }

        public double soNgayCong { get; set; }

        public double soNgayCongChoViec { get; set; }

        public double soNgayNghiPhep { get; set; }

        public double soNgayLe { get; set; }

        public double soNgayTangCaThuong { get; set; }

        public double soNgayTangChuNhat { get; set; }

        public double soNgayTangCaLe { get; set; }

        public PhieuCapNhatNgayCongNhanVien()
        {
        }

    }
}