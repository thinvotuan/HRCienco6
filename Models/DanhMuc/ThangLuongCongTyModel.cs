using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace BatDongSan.Models.DanhMuc
{
    public class ThangLuongCongTyModel
    {
        public int? bac { get; set; }

        public int? capBacChucVu { get; set; }

        public System.Nullable<decimal> tongLuongKhoiDiem { get; set; }

        public System.Nullable<decimal> tongLuongChinhThuc { get; set; }

        public System.Nullable<decimal> tongLuongThucThu { get; set; }

        public System.Nullable<decimal> luongCanBanKhoiDiem { get; set; }

        public System.Nullable<decimal> luongCanBanChinhThuc { get; set; }

        public System.Nullable<decimal> luongCanBanThucThu { get; set; }

        public System.Nullable<decimal> luongThanhTichKhoiDiem { get; set; }

        public System.Nullable<decimal> luongThanhTichChinhThuc { get; set; }

        public System.Nullable<decimal> luongThanhTichThucThu { get; set; }
    }
}