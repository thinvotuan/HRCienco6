using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace BatDongSan.Models.DanhMuc
{
    public class DMPhanca
    {
        public int maPhanca { get; set; }
        public string tenPhanca { get; set; }
        public int maKhoiTinhLuong { get; set; }
        public string tenKhoiTinhLuong { get; set; }
        public System.Nullable<System.DateTime> ngayLap { get; set; }
        public string ghiChu { get; set; }
        public List<DMPhanCaNhanVien> ListNhanVien { get; set; }
        public int soLuongNhanVienTrongCa { get; set; }
    }
    public class DMPhanCaNhanVien
    {
        public string  maNhanVien { get; set; }
        public string hoVaTen { get; set; }
        public System.Nullable<System.DateTime> ngayVaoLam { get; set; }
        public System.Nullable<System.DateTime> ngayApDung  { get; set; }
        public string tenPhongBan { get; set; }
        public string tenChucDanh { get; set; }
        public string tenKhoiTinhLuong { get; set; }
        public string hinhThucTinhLuong { get; set; }
        public string tenPhanCa { get; set; }
        public int? CapBac { get; set; }
    }
    public class DMPhanCaChiTiet
    {
        public decimal? soGio { get; set; }
        public decimal? heSo { get; set; }
        public System.Nullable<System.DateTime> gioBatDau { get; set; }
        public System.Nullable<System.DateTime> gioKetThuc { get; set; }
        public int? diTreChoPhep { get; set; }
        public int? veSomChoPhep { get; set; }
        public bool nghiGiuaCa { get; set; }
        public int? thoiGianNghiGiuaCa { get; set; }
        public System.Nullable<System.DateTime> batDauNghiGiuaCa { get; set; }
        public System.Nullable<System.DateTime> ketThucNghiGiuaCa { get; set; }
        public string ghiChu { get; set; }

        //public string ghiChu { get; set; }

    }
    public partial class tbl_NS_PhanCaChiTiet
    {
        public string hourStartShift { get; set; }
        public string hourEndShift { get; set; }
        public string hourStartBetweenShift { get; set; }
        public string hourEndBetweenShift { get; set; }
    }
}