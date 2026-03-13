using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using BatDongSan.Models.HeThong;

namespace BatDongSan.Models.DanhMuc
{
    public class PhongBanModel
    {
        public string MaPhongBan { get; set; }

        public string Ten { get; set; }

        public string ChucNang { get; set; }

        public System.Nullable<int> CapBac { get; set; }

        public string GhiChu { get; set; }        

        public string maCha { get; set; }

        public int? maPhanCa { get; set; }

        public string maNhanVienDuyet { get; set; }

        public string tenNhanVienDuyet { get; set; }

        public string maNhomDanhGia { get; set; }
    }
}