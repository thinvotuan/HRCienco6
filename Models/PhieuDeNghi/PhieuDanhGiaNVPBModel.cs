using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace BatDongSan.Models.PhieuDeNghi
{
    public class PhieuDanhGiaNVPBModel
    {
        public string MaPhieu { get; set; }

        public string MaNhanVienLP { get; set; }

        public string HoTenNVLP { get; set; }

        public DateTime? TuNgay { get; set; }

        public DateTime? DenNgay { get; set; }

        public DateTime NgayLap { get; set; }

        public int XacNhan { get; set; }

        public string NoiDung { get; set; }
    }

    public class PhieuDanhGiaChiTiet
    {
        public string MaPhieu { get; set; }

        public string MaNhanVien { get; set; }

        public string HoTenNV { get; set; }

        public string GhiChu { get; set; }

        public double TyLeHoanThanh { get; set; }

        public string TenPhongBan { get; set; }

        public string KeHoachLamViecTaiNha { get; set; }
    }
}