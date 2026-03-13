using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using BatDongSan.Models.HeThong;
using BatDongSan.Models.NhanSu;

namespace BatDongSan.Models.NhanSu
{
    public class ThuNhapKhac
    {
        public string maPhieu { get; set; }

        public string maNhanVien { get; set; }

        public System.Nullable<System.DateTime> ngayLap { get; set; }

        public System.Nullable<System.Decimal> soTien { get; set; }
        
        public string maLoaiThuNhapKhac { get; set; }

        public int? thang { get; set; }
        public int? nam { get; set; }

        public string lyDo { get; set; }

        public byte trangThai { get; set; }        

        public string nguoiLap { get; set; }

        public int? TrangThaiDuyet { get; set; }

        public NhanVienModel NhanVien { get; set; }

        public NhanVienModel NhanVienLapPhieu { get; set; }      

        public DMNguoiDuyet Duyet { get; set; }

        public int maQuiTrinhDuyet { get; set; }

        public int id { get; set; }

        public ThuNhapKhac()
        {
            Duyet = new DMNguoiDuyet();
        }
    }
}