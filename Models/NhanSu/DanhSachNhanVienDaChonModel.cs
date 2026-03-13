using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace BatDongSan.Models.NhanSu
{
    public class DanhSachNhanVienDaChonModel
    {
        public string MaNhanVien { get; set; }
        public string TenNhanVien { get; set; }
        public string MaPhongBan { get; set; }
        public string TenPhongBan { get; set; }
        public string MaChucDanh { get; set; }
        public string TenChucDanh { get; set; }
        public string Avatar { get; set; }
        public int? CapBac { get; set; }
        public DateTime? NgaySinh { get; set; }
        public int STT { get; set; }

    }
}