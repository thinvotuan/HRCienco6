using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using BatDongSan.Models.HeThong;

namespace BatDongSan.Models.VIP
{
    public class NhanVienVIPModel
    {
        public int id { get; set; }
        public string maNhanVien { get; set; }
        public string soHopDong { get; set; }
        public string tenHopDong { get; set; }
        public string soPhuLuc { get; set; }
        public System.Nullable<decimal> tongLuong { get; set; }
        public System.Nullable<decimal> luongCoBan { get; set; }
        public System.Nullable<decimal> phuCapThamNien { get; set; }
        public System.Nullable<decimal> phuCapChucVu { get; set; }
        public System.Nullable<decimal> phuCapCongTrinh { get; set; }
        public System.Nullable<decimal> phuCapThuHut { get; set; }
        public System.Nullable<decimal> phuCapDacBiet { get; set; }
        public System.Nullable<decimal> phuCapPhatSinh { get; set; }
        public string tenNhanVien { get; set; }
        public string nguoiLap { get; set; }
        public System.Nullable<System.DateTime> ngayVaoLam { get; set; }
        public System.Nullable<System.DateTime> ngayLap { get; set; }
        public string email{get;set;}
        public string ghiChu { get; set; }
        public float thoiGianCongTac { get; set; }
        public string soTaiKhoan { get; set; }

        public string maChiNhanhNganHang { get; set; }
        public System.Nullable<decimal> truyLanh { get; set; }
        public System.Nullable<decimal> truyThu { get; set; }
       
        public bool Net { get; set; }
        
        

        

        public NhanVienVIPModel()
        {
           
        }

        public bool khongChamCong { get; set; }
    }
}