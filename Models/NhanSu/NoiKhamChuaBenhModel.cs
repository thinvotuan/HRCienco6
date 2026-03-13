using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace BatDongSan.Models.NhanSu
{
    public class NoiKhamChuaBenhModel
    {
        public string maNoiKham { get; set; }

        public string tenNoiKham { get; set; }

        public string diaChi { get; set; }

        public string maTinhThanh { get; set; }

        public string ghiChu { get; set; }

        public string nguoiLap { get; set; }

        public System.Nullable<System.DateTime> ngayCapNhat { get; set; }

        public System.Nullable<int> loai { get; set; }

        public string tenLoaiBenhVien { get; set; }
        
        public string tenTinhThanh { get; set; }

    }
}