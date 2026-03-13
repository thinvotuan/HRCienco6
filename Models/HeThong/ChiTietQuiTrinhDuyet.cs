using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace BatDongSan.Models.HeThong
{
    public class ChiTietQuiTrinhDuyet
    {
        public int ThuTuDuyet { get; set; }

        public int TrangThai { get; set; }

        public string MaBuocDuyet { get; set; }

        public string TenBuocDuyet { get; set; }

        public string TenHienThiFooter { get; set; }

        public string MaNhanVien { get; set; }

        public string TenNhanVien { get; set; }

        public string MaMau { get; set; }

        public ChiTietQuiTrinhDuyet()
        {
        }
    }
}