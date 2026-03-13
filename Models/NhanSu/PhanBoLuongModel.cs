using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace BatDongSan.Models.NhanSu
{
    public class PhanBoLuongModel
    {
        public string tenNhanVien { get; set; }

        public string maNhanVien { get; set; }

        public decimal? mucLuongThoaThuan { get; set; }

        public decimal? tongPhuCapKhongTheoTyLe { get; set; }

        public decimal? tongPhuCapTheoTyLe { get; set; }

        public DateTime? ngayApDung { get; set; }

        public string mucTinhPhanBo { get; set; }

        public decimal? tongPhuCap { get; set; }

        public decimal? luongCoBan { get; set; }

        public decimal? tongTyle { get; set; }

        public int flag { get; set; }

        public IList<BangPhanTachLuongChiTietModel> chiTiets { get; set; }
    }
}