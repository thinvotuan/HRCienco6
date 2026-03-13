using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using BatDongSan.Models.HeThong;
using BatDongSan.Models.NhanSu;

namespace BatDongSan.Models.PhieuDeNghi
{
    public class PhieuThoiViec
    {
        public string maPhieu { get; set; }

        public string maNhanVien { get; set; }

        public System.Nullable<System.DateTime> ngayLap { get; set; }
        
        public System.Nullable<System.DateTime> ngayThoiViec { get; set; }

        public int? maLoaiLyDo { get; set; }

        public string ghiChu { get; set; }

        public byte trangThai { get; set; }        

        public string nguoiLap { get; set; }

        public int? TrangThaiDuyet { get; set; }

        public NhanVienModel NhanVien { get; set; }

        public NhanVienModel NhanVienLapPhieu { get; set; }      

        public DMNguoiDuyet Duyet { get; set; }

        public int maQuiTrinhDuyet { get; set; }

        public PhieuThoiViec()
        {
            Duyet = new DMNguoiDuyet();
        }
    }
}