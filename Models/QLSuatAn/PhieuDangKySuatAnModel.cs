using BatDongSan.Models.QLSuatAn.LuuVet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace BatDongSan.Models.QLSuatAn
{
    public class PhieuDangKySuatAnModel
    {
        
        public List<tbl_SA_CuaHang> lstCuaHang = new List<tbl_SA_CuaHang>();
        public int id { get; set; }
        public DateTime? ngayLapPhieu { get; set; }
        public String maCuaHang { get; set; }
        public String hoTenNhanVien { get; set; }
        public String tenCuaHang { get; set; }
        

        public string maNhanVien { get; set; }

        public string avatar { get; set; }

        public string maPhieu { get; set; }

        //public int? soSuatAnTrongNgay { get; set; }

        //public int? suatAnTrenNhanVien { get; set; }

        //public int? tongSACuaHangConLai { get; set; }

        //public int? tongSANhanVienCuaHang { get; set; }

        //public int? tongSACuaHangDaDK { get; set; }

        //public int? soLanNhanVienCuaHangDK { get; set; }

        //public int? soThuTuPhucVuTaiQuan { get; set; }

        //public int? soSuatAnConLaiTaiQuan { get; set; }

        public string maVach { get; set; }

        public int? soSuatAnTrongNgayDaDangKyCH { get; set; }

        public int? soLanAnCuaNhanVienTrongThang { get; set; }

        public int? soSuatAnTrongNgayCuaCuaHang { get; set; }

        public int? suatAnTrenNhanVien { get; set; }

        public int soSuatAnConLaiTrongNgayCH { get; set; }

        public int soLanAnConLaiTaiCuaHangNV { get; set; }

        public string avartar { get; set; }

        public string ghiChu { get; set; }

        public bool themNhanVien { get; set; }

        public string hoTenNhanVienThuHuong { get; set; }

        public string maNhanVienThuHuong { get; set; }

        public int? soThuTuPhucVuTaiQuan { get; set; }

        public int? soSuatAnConLaiCuaNVTaiQuan { get; set; }

        public string hinhAnhQuan { get; set; }

        public string tenMonAn { get; set; }

        public string maMonAn { get; set; }

        public int? soLuongMonAnBanDau { get; set; }

        public int soLuongMonAnDaDat { get; set; }

        public string hinhMonAn { get; set; }

        public string maCuaHangMonAn { get; set; }

        public string thoiGianDuKienPhucVu { get; set; }

        public bool? tinhTrangHetHieuLuc { get; set; }
    }
}