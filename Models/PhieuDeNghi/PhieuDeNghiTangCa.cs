using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using BatDongSan.Models.HeThong;
using BatDongSan.Models.NhanSu;

namespace BatDongSan.Models.PhieuDeNghi
{
    public class PhieuDeNghiTangCa
    {
        public string maPhieu { get; set; }

        public string maPhongBan { get; set; }

        public string tenPhongBan { get; set; }

        public string tenChucDanh { get; set; }

        public System.Nullable<System.DateTime> thoiGianBatDau { get; set; }

        public System.Nullable<System.DateTime> thoiGianKetThuc { get; set; }

        public System.Nullable<System.DateTime> ngayTangCa { get; set; }

        public byte trangThai { get; set; }

        public string ghiChu { get; set; }

        public string nguoiLap { get; set; }

        public System.Nullable<System.DateTime> ngayLap { get; set; }

        public int? TrangThaiDuyet { get; set; }

        public NhanVienModel  NhanVien { get; set; }

        public DateTime? thoiGianNghiGiuaCa { get; set; }

        public DateTime? thoiGianNghiKetThucGiuaCa { get; set; }

        public DateTime? gioBatDau { get; set; }

        public DateTime? gioKetThuc { get; set; }

        public int? maLoaiTangCa { get; set; }

        public int? soNgayCongTac { get; set; }

        public DMNguoiDuyet Duyet { get; set; }

        public DateTime? soGioTangCa { get; set; }

        public double heSoTangCa { get; set; }

        public string tenLoaiTangCa { get; set; }

        public int soGioNghiGiuaCa { get; set; }

        public string hinhThucTangCa { get; set; }
        public int maQuiTrinhDuyet { get; set; }

        public bool? tangCaTinhChamCong { get; set; }

        public PhieuDeNghiTangCa()
        {
            Duyet = new DMNguoiDuyet();
        }
    }
}