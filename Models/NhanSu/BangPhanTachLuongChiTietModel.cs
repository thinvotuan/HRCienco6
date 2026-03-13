using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace BatDongSan.Models.NhanSu
{
    public class BangPhanTachLuongModel
    {
        public BangPhanTachLuongModel() { }

        public string maMucLuong { get; set; }

        public string tenMucLuong { get; set; }

        public decimal mucLuongTu { get; set; }

        public decimal mucLuongDen { get; set; }

        public decimal luongCoBan { get; set; }

        public decimal phuCapKhongTheoTiLe { get; set; }

        public decimal phuCapTheoTiLe { get; set; }

        public string ghiChu { get; set; }

        public IList<BangPhanTachLuongChiTietModel> chiTiets { get; set; }
    }

    public class BangPhanTachLuongChiTietModel
    {
        public int id { get; set; }

        public string maMucLuong { get; set; }

        public string tenMucLuong { get; set; }

        public string maPhuCap { get; set; }

        public string tenPhuCap { get; set; }

        public int idLoaiPhuCap { get; set; }

        public string tenLoaiPhuCap { get; set; }

        public string loaiTyLe { get; set; }

        public System.Nullable<decimal> tyLe { get; set; }

        public string salaryTemplate { get; set; }

        public string ghiChu { get; set; }
    }
}