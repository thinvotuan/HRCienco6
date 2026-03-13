using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace BatDongSan.Models.NhanSu
{
    public class ThangTinhThueTNCNModel
    {
        public int id { get; set; }

        public string moTa { get; set; }

        public System.Nullable<System.DateTime> ngayApDung { get; set; }

        public byte tinhTrang { get; set; }

        public string ghiChu { get; set; }

        public string nguoiLap { get; set; }

        public System.Nullable<System.DateTime> ngayCapNhat { get; set; }

        public System.Nullable<decimal> giamTruCaNhan { get; set; }
       
        public System.Nullable<decimal> giamTruGiaCanh { get; set; }

        public IList<tbl_NS_ThangTinhThueTNCNChiTiet> chiTiets { get; set; }
    }
}