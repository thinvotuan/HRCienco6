using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using BatDongSan.Models.HeThong;
using BatDongSan.Models.NhanSu;

namespace BatDongSan.Models.NhanSu
{
    public class DeNghiChiLuong
    {
        public string soPhieu { get; set; }
        public DateTime ngayLap { get; set; }
        public string tenNguoiLap { get; set; }
        public string maNguoiLap { get; set; }
        public int? thang { get; set; }
        public int? nam { get; set; }
        public string noiDung { get; set; }
        public string maCongTrinh { get; set; }
        public string doiTuongBaoHiem { get; set; }
        public string doiTuongThue { get; set; }
        public string doiTuongLuong { get; set; }

        public int? trangThai { get; set; }
        public decimal? tongTienThucChuyen { get; set; }
        public string tenBuocDuyet { get; set; }

        public int maQuiTrinhDuyet { get; set; }

        public DMNguoiDuyet Duyet { get; set; }

        public List<DeNghiChiLuongChiTiet> DeNghiChiLuongChiTiet { get; set; }
    }
    public class DeNghiChiLuongChiTiet
    {
        public int id { get; set; }
        public string TenBoPhanTinhLuong { get; set; }
        public double? LuongThang { get; set; }
        public double? PhuCapCT { get; set; }
        public double? PhuCapKhac { get; set; }
        public double? BHXH { get; set; }
        public double? ThueTNCN { get; set; }
        public double? TruyLanhTruyThu { get; set; }
        public double? ChuyenKhoan { get; set; }
        public string MaBoPhan { get; set; }
        public double? DoanPhi { get; set; }
    }
    public class QuiTrinhDuyet
    {

        public int Id { get; set; }

        public string TenQuiTrinh { get; set; }

        public string ChuoiBuocDuyet { get; set; }
    }
}