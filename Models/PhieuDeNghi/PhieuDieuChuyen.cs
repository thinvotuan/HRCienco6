using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using BatDongSan.Models.HeThong;
using BatDongSan.Models.NhanSu;

namespace BatDongSan.Models.PhieuDeNghi
{
    public class PhieuDieuChuyen
    {
        public string maPhieu { get; set; }
        public System.Nullable<System.DateTime> ngayLap { get; set; }
        public NhanVienModel NhanVienLapPhieu { get; set; }
        public NhanVienModel NhanVien { get; set; }
        public System.Nullable<System.DateTime> ngayChuyen { get; set; }

        public string maPhongBan { get; set; }
        public string tenPhongBan { get; set; }

        public string maChucDanhMoi { get; set; }
        public string tenChucDanhMoi { get; set; }

        public string maBoPhanTinhLuongMoi { get; set; }
        public string tenBoPhanTinhLuongMoi { get; set; }

        public string maCoSoMoi { get; set; }
        public string tenCoSoMoi { get; set; }


        public string maPhongBanCu { get; set; }
        public string tenPhongBanCu { get; set; }

        public string maChucDanhCu { get; set; }
        public string tenChucDanhCu { get; set; }

        public string maBoPhanTinhLuongCu { get; set; }
        public string tenBoPhanTinhLuongCu { get; set; }

        public string maCoSoCu { get; set; }
        public string tenCoSoCu { get; set; }

        public short maPhanCaCu { get; set; }
        public string tenPhanCaCu { get; set; }
        public string tenPhanCaMoi { get; set; }
        public short maPhanCaMoi { get; set; }
        public string noiDung { get; set; }
        public string loaiDieuChuyen { get; set; }

        public int? TrangThaiDuyet { get; set; }

        public int maQuiTrinhDuyet { get; set; }

        public DMNguoiDuyet Duyet { get; set; }

        public string maNhanVienLapPhieu { get; set; }

        public PhieuDieuChuyen()
        {
            Duyet = new DMNguoiDuyet();
        }
    }
}