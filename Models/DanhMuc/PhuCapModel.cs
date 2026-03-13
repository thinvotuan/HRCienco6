using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace BatDongSan.Models.DanhMuc
{
    public class PhuCapModel
    {
        public PhuCapModel() { }

        public string maPhuCap { get; set; }

        public string tenPhuCap { get; set; }

        public int loaiPhuCap { get; set; }

        public System.Nullable<decimal> soTien { get; set; }

        public bool coTinhThueTNCN { get; set; }
        public bool coTinhBaoHiem { get; set; }

        public string ghiChu { get; set; }

        public string nguoiLap { get; set; }

        public System.Nullable<System.DateTime> ngayLap { get; set; }

        public System.Nullable<System.DateTime> ngayCapNhat { get; set; }

        public string maPhongBan { get; set; }

        public string tenLoaiPhuCap { get; set; }

        public string tenPhongBan { get; set; }

        //public string tenPhongBan { get; set; }
    }
    public class LyDoThoiViecModel
     {

        public string tenLyDoThoiViec { get; set; }

        public string nguoiLap { get; set; }

        public int maLyDoThoiViec { get; set; }

        public DateTime? ngayLap { get; set; }
    }
    public class LoaiTangCaModel
    {

        public string loaiTangCa { get; set; }

        public string nguoiLap { get; set; }

        public int id { get; set; }
        public double?  heSoTangCa { get; set; }

        public DateTime? ngayLap { get; set; }
    }
}