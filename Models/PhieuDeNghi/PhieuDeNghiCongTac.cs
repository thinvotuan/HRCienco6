using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using BatDongSan.Models.HeThong;
using BatDongSan.Models.NhanSu;

namespace BatDongSan.Models.PhieuDeNghi
{
    public class PhieuDeNghiCongTac
    {
        public string maPhieu { get; set; }


        public string tenPhongBan { get; set; }

        public string tenChucDanh { get; set; }

        public System.Nullable<System.DateTime> ngayBatDau { get; set; }

        public System.Nullable<System.DateTime> ngayKetThuc { get; set; }

        public int maNoiCongTac { get; set; }

        public string tenNoiCongTac { get; set; }
        public string phongBanCongTac { get; set; }

        public string lyDo { get; set; }

        public string phuongTien { get; set; }

        public System.Nullable<System.Decimal> phuCap { get; set; }

        public byte trangThai { get; set; }

        public string ghiChu { get; set; }

        public string nguoiLap { get; set; }

        public System.Nullable<System.DateTime> ngayLap { get; set; }

        public int? TrangThaiDuyet { get; set; }

        public NhanVienModel NhanVien { get; set; }

        public NhanVienModel NhanVienLapPhieuTT { get; set; }

        public string maTinh { get; set; }
        public string tenTinh { get; set; }

        public string maQuocGia { get; set; }
        public string tenQuocGia { get; set; }

        public DateTime? gioBatDau { get; set; }

        public DateTime? gioKetThuc { get; set; }

        public bool? loaiCongTac { get; set; }

        public int? soNgayCongTac { get; set; }

        public DMNguoiDuyet Duyet { get; set; }

        public int maQuiTrinhDuyet { get; set; }

        public PhieuDeNghiCongTac()
        {
            Duyet = new DMNguoiDuyet();
        }
    }
}