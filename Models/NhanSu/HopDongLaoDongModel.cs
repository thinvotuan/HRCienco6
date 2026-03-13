using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using BatDongSan.Models.HeThong;

namespace BatDongSan.Models.NhanSu
{
    public class HopDongLaoDongModel
    {
        public int id { get; set; }

        public string soHopDong { get; set; }

        public string tenHopDong { get; set; }

        public System.Nullable<System.DateTime> ngayKy { get; set; }

        public string maNhanVien { get; set; }

        public string tenNhanVien { get; set; }

        public byte idThoiHanHopDong { get; set; }

        public string tenThoiHanHopDong { get; set; }

        public short soThang { get; set; }

        public System.Nullable<decimal> luongHopDong { get; set; }

        public string maChucDanh { get; set; }

        public string tenChucDanh { get; set; }

        public System.Nullable<System.DateTime> ngayBatDau { get; set; }

        public System.Nullable<System.DateTime> ngayKetThuc { get; set; }

        public string ghiChu { get; set; }

        public string nguoiLap { get; set; }

        public System.Nullable<System.DateTime> ngayLap { get; set; }

        public System.Nullable<bool> tinhTrang { get; set; }

        public string maLoaiHopDong { get; set; }

        public string tenLoaiHopDong { get; set; }

        public System.Nullable<decimal> tongLuong { get; set; }

        public System.Nullable<decimal> luongThoaThuan { get; set; }

        public System.Nullable<decimal> tyLeHuong { get; set; }

        public int maQuiTrinhDuyet { get; set; }

        public DMNguoiDuyet Duyet { get; set; }

        public decimal luongDongBH { get; set; }

        public decimal khoanBoSungLuong { get; set; }

        public double doanPhi { get; set; }

        public double dangPhi { get; set; }

        public bool tienAnGiuaCa { get; set; }

        public decimal tienDienThoai { get; set; }

        public double phuCapLuong { get; set; }

        public string maLuongChoViec { get; set; }

        public double tyLeLuongChoViec { get; set; }

        public double congChoViec { get; set; }

        public decimal luongChoViec { get; set; }

        public string soPhuLuc { get; set; }

        public NhanVienModel NguoiLap { get; set; }

        public NhanVienModel NhanVien { get; set; }

        public double phuCapKhac { get; set; }

        public HopDongLaoDongModel()
        {
            Duyet = new DMNguoiDuyet();
            //NguoiLap = new NhanVienModel();
            //NhanVien = new NhanVienModel();
        }
    }
}