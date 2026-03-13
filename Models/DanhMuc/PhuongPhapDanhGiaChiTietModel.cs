using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using BatDongSan.Models.HeThong;

namespace BatDongSan.Models.DanhMuc
{
    public class PhuongPhapDanhGiaChiTietModel
    {
        public string MaPhuongPhapDanhGia { get; set; }

        public string MaNhomDanhGia { get; set; }

        public string TenNhomDanhGia { get; set; }

        public double SoDiem { get; set; }

        public int ThuHang { get; set; }
      
    }
}