using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace BatDongSan.Models.VIP
{
    public class BangLuongNVVIPModel 
    {
        public IList<BLNVVIP> blVips { get; set; }
        
    }
    public class BLNVVIP
    {
    public string maNhanVien { get; set; }

        public int thang { get; set; }

        public decimal? tongLuong { get; set; }

        public double? ngayCongChuan { get; set; }

        public double? ngayTinhLuong { get; set; }

        public decimal? thueTNCN { get; set; }

        public decimal? baoHiemXH { get; set; }

        public decimal? conLaiPhaiChuyen { get; set; }

        public string tenNhanVien { get; set; }
    }
}
