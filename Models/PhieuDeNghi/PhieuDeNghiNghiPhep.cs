using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using BatDongSan.Models.HeThong;
using BatDongSan.Models.NhanSu;

namespace BatDongSan.Models.PhieuDeNghi
{
    public class PhieuDeNghiNghiPhep
    {
        public string maPhieu { get; set; }

        public NhanVienModel NhanVien { get; set; }

        public NhanVienModel NhanVienLapPhieuTT { get; set; }

        public decimal? soNgayNamVien { get; set; }

        public bool? thoiGianNghi { get; set; }

        public System.Nullable<System.DateTime> ngayBatDau { get; set; }

        public int? loaiNgayBatDau { get; set; }

        public System.Nullable<System.DateTime> ngayKetThuc { get; set; }

        public int? loaiNgayKetThuc { get; set; }

        public string loaiNghiPhep { get; set; }

        public int? soNgayRangBuoc { get; set; }

        public string lyDo { get; set; }

        public decimal? phanTramHuongNamVien { get; set; }

        public decimal? soNgayNghiThucTe { get; set; }

        public int? trangThai { get; set; }

        public string ghiChu { get; set; }

        public string nguoiLap { get; set; }

        public System.Nullable<System.DateTime> ngayLap { get; set; }

        public decimal? soNgayNghiPhep { get; set; }

        public decimal? phanTramTinhHuong { get; set; }

        public bool? loaiNghiTrucTuyen { get; set; }

        public IList<PhieuDeNghiNghiPhepChiTiet> DeNghiChiTiet { get; set; }

        public System.Nullable<decimal> soNgayTangCa { get; set; }

        public DMNguoiDuyet Duyet { get; set; }

        public int maQuiTrinhDuyet { get; set; }

        public PhieuDeNghiNghiPhep()
        {
            Duyet = new DMNguoiDuyet();
        }
    }
    public class PhieuDeNghiNghiPhepChiTiet
    {
    }
}