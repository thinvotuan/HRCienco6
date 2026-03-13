using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace BatDongSan.Models.NhanSu
{
    public class PhepNamNhanVien
    {
        public int Nam { get; set; }
        public NhanVienModel NhanVien { get; set; }
        public System.Nullable<System.DateTime> ngayLap { get; set; }
        IList<DanhSachNhanVien> DanhSach { get; set; }
        public string maPhongBan { get; set; }
        public string tenPhongBan { get; set; }

        public string maChiNhanh { get; set; }
        public string tenChiNhanh { get; set; }

        public string maChucVu { get; set; }
        public string tenChucVu { get; set; }

        public string maNhanVien { get; set; }
        public string tenNhanVien { get; set; }


        public System.Nullable<decimal> soNgayPhepTonNamTruoc { get; set; }
        public System.Nullable<decimal> soNgayPhepTonDuocDuyet { get; set; }
        public System.Nullable<decimal> soNgayPhepQuiDinh { get; set; }
        public System.Nullable<decimal> soNgayPhepThamNien { get; set; }
        public System.Nullable<decimal> tongSoNgayPhep { get; set; }

        public decimal soNgayPhepDacBiet { get; set; }
    }
    public class DanhSachNhanVien
    {
        public string maPhongBan { get; set; }
        public string tenPhongBan { get; set; }

        public string maChiNhanh { get; set; }
        public string tenChiNhanh { get; set; }

        public string maChucVu { get; set; }
        public string tenChucVu { get; set; }

        public string maNhanVien { get; set; }
        public string tenNhanVien { get; set; }

        public System.Nullable<decimal> soNgayPhepTonNamTruoc { get; set; }
        public System.Nullable<decimal> soNgayPhepTonDuocDuyet { get; set; }
        public System.Nullable<decimal> soNgayPhepQuiDinh { get; set; }
        public System.Nullable<decimal> soNgayPhepThamNien { get; set; }
        public System.Nullable<decimal> tongSoNgayPhep { get; set; }

        public decimal soNgayPhepDacBiet { get; set; }

    }
}