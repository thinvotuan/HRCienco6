using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using BatDongSan.Models.NhanSu;

namespace BatDongSan.Models.HeThong
{
    public class DMNguoiDuyet
    {
        public string maPhieu { get; set; }
        public string maCongViec { get; set; }
        public int tinhTrang { get; set; }
        public NhanVienModel nguoiDuyet { get; set; }
        public int? trangThai { get; set; }
        public string lyDo { get; set; }
        public DateTime? ngayDuyet { get; set; }

        public string maNhanVien { get; set; }

        public string tenNhanVien { get; set; }

        public int soNgayNghiPhep { get; set; }

        public string tenBuocDuyet { get; set; }

        public string maBuocDuyet { get; set; }

        public int Id { get; set; }

        public DMNguoiDuyet() { }

        public DMNguoiDuyet(string maPhieu)
        {
            this.maPhieu = maPhieu;
            this.nguoiDuyet = new NhanVienModel(string.Empty, string.Empty);
        }
    }
}