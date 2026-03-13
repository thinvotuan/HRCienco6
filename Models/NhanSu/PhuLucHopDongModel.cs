using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using BatDongSan.Models.HeThong;

namespace BatDongSan.Models.NhanSu
{
    public class PhuLucHopDongModel
    {

        public int id { get; set; }

        public string soPhuLuc { get; set; }

        public string soHopDong { get; set; }

        public string maNhanVien { get; set; }

        public string tenNhanVien { get; set; }

        public string tenChucDanh { get; set; }

        public int idHopDong { get; set; }

        public string noiDungThayDoi { get; set; }

        public System.Nullable<System.DateTime> ngayHieuLuc { get; set; }
        public System.Nullable<System.DateTime> giaHanDen { get; set; }

        public System.Nullable<decimal> soLuongCu { get; set; }

        public System.Nullable<decimal> mucDieuChinh { get; set; }

        public System.Nullable<decimal> mucLuongMoi { get; set; }

        public string ghiChu { get; set; }

        public string nguoiLap { get; set; }

        public string nguoiCapNhat { get; set; }

        public System.Nullable<System.DateTime> ngayLap { get; set; }

        public int maQuiTrinhDuyet { get; set; }

        public DMNguoiDuyet Duyet { get; set; }

        public decimal luongDongBH { get; set; }

        public decimal khoanBoSungLuong { get; set; }

        public double doanPhi { get; set; }

        public double dangPhi { get; set; }

        public bool tienAnGiuaCa { get; set; }

        public decimal tienDienThoai { get; set; }

        public decimal luong { get; set; }

        public double phuCapLuong { get; set; }

        public decimal tongLuong { get; set; }

        public double phuCapKhac { get; set; }

        public PhuLucHopDongModel()
        {
            Duyet = new DMNguoiDuyet();
        }

        public byte? idThoiHanHopDong { get; set; }

        public int soThang { get; set; }

        public string tenThoiHanHopDong { get; set; }

        public string maLoaiHopDong { get; set; }

        public string tenLoaiHopDong { get; set; }
    }
}