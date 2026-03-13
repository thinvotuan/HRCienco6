using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Utils.Paging;
using BatDongSan.Helper.Utils;
using System.Data;
using System.Globalization;
using System.Text;
using System.IO;
using BatDongSan.Models.NhanSu;
namespace BatDongSan.Controllers.TinhCong
{
    public class ChamCongController : ApplicationController
    {

        private LinqNhanSuDataContext nhanSuContext = new LinqNhanSuDataContext();
        private IList<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan> phongBans;
        private StringBuilder buildTree;
        private readonly string MCV = "ChamCongAdmin";
        private bool? permission;
        public ActionResult XemTinhHinhRaVao()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            thang(DateTime.Now.Month);
            nam(DateTime.Now.Year);
            return View("");
        }
        public ActionResult LoadXemTinhHinhRaVao(string qSearch, int thang, int nam, int _page = 0)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            //BatDongSan.Models.ChamCong.LinqChamCongServerDataContext contextCC = new BatDongSan.Models.ChamCong.LinqChamCongServerDataContext();
            string maNhanVien = GetUser().manv;
            int page = _page == 0 ? 1 : _page;
            int pIndex = page;
            ////Get ma cham cong
            //var getMaCC = nhanSuContext.tbl_NS_NhanViens.Where(d => d.maNhanVien == maNhanVien).FirstOrDefault();
            //if (getMaCC != null)
            //{
            //    var maChamCong = getMaCC.maChamCong;

                int total = nhanSuContext.sp_NS_XemTinhHinhRaVao(maNhanVien, thang, nam, qSearch).Count();
                PagingLoaderController("/ChamCong/XemTinhHinhRaVao/", total, page, "?qsearch=" + qSearch + "&maNhanVien=" + maNhanVien);
                ViewData["lsDanhSach"] = nhanSuContext.sp_NS_XemTinhHinhRaVao(maNhanVien, thang, nam, qSearch).Skip(start).Take(offset).ToList();

                ViewData["qSearch"] = qSearch;
                return PartialView("_LoadXemTinhHinhRaVao");
            //}
            //else {
            //    return View("error");
            //}
        }
        //public ActionResult XemTinhHinhRaVaoCongNhan()
        //{
        //    #region Role user
        //    permission = GetPermission("XemRaVaoToanCTy", BangPhanQuyen.QuyenXem);
        //    if (!permission.HasValue)
        //        return View("LogIn");
        //    if (!permission.Value)
        //        return View("AccessDenied");
        //    #endregion
        //    thang(DateTime.Now.Month);
        //    nam(DateTime.Now.Year);
        //    return View("");
        //}
        //public ActionResult LoadXemTinhHinhRaVaoCongNhan(string qSearch, int thang, int nam, int _page = 0)
        //{
        //    #region Role user
        //    permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
        //    if (!permission.HasValue)
        //        return View("LogIn");
        //    if (!permission.Value)
        //        return View("AccessDenied");
        //    #endregion
        //    //BatDongSan.Models.ChamCong.LinqChamCongServerDataContext contextCC = new BatDongSan.Models.ChamCong.LinqChamCongServerDataContext();
        //    string maNhanVien = GetUser().manv;
        //    int page = _page == 0 ? 1 : _page;
        //    int pIndex = page;
        //    ////Get ma cham cong
        //    //var getMaCC = nhanSuContext.tbl_NS_NhanViens.Where(d => d.maNhanVien == maNhanVien).FirstOrDefault();
        //    //if (getMaCC != null)
        //    //{
        //    //    var maChamCong = getMaCC.maChamCong;

        //    int total = nhanSuContext.sp_NS_XemTinhHinhRaVao(maNhanVien, thang, nam, qSearch).Count();
        //    PagingLoaderController("/ChamCong/XemTinhHinhRaVao/", total, page, "?qsearch=" + qSearch + "&maNhanVien=" + maNhanVien);
        //    ViewData["lsDanhSach"] = nhanSuContext.sp_NS_XemTinhHinhRaVao(maNhanVien, thang, nam, qSearch).Skip(start).Take(offset).ToList();

        //    ViewData["qSearch"] = qSearch;
        //    return PartialView("_LoadXemTinhHinhRaVao");
        //    //}
        //    //else {
        //    //    return View("error");
        //    //}
        //}
        public ActionResult XemBangLuong()
        {
            #region Role user
            permission = GetPermission("XemBangLuong", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            nam(DateTime.Now.Year);
            return View("");
        }
        public ActionResult ViewChiTietLuong(string thang, string nam)
        {
            #region Role user
            permission = GetPermission("XemBangLuong", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion    
            string maNhanVien = GetUser().manv;
            var dsMauIn = nhanSuContext.GetTable<BatDongSan.Models.HeThong.Sys_PrintTemplate>().Where(d => d.maMauIn == "MIBLNV").FirstOrDefault();
            string noiDung = string.Empty;
            var ds = nhanSuContext.tbl_NS_BangLuongNhanViens
                               .Where(t => t.maNhanVien == maNhanVien && t.thang.ToString() == thang && nam == t.nam.ToString()).FirstOrDefault();
            ViewData["chiTiet"] = dsMauIn;
            if (dsMauIn != null)
            {
                var tongKhauTru = (ds.baoHiem) + (ds.thue ?? 0) + (ds.doanPhi ?? 0) + (ds.dangPhi ?? 0);
                noiDung = dsMauIn.html.Replace("{$thang}", thang)
                    .Replace("{$nam}", nam)
                    .Replace("{$hoVaTen}", ds.hoTen)
                    .Replace("{$tongLuong}", String.Format("{0:###,##0}", ds.tongLuong))
                    .Replace("{$khoanBSLVSLuongBH}", String.Format("{0:###,##0}", ds.luongDongBaoHiem))
                    .Replace("{$luongThoaThuan}", String.Format("{0:###,##0}", ds.khoanBoSungLuong))
                    .Replace("{$ngayCongChuan}", Convert.ToString(ds.ngayCongChuan))
                    .Replace("{$ngayCongTinhLuong}", Convert.ToString(ds.tongNgayCong))
                    .Replace("{$ngayCong}", Convert.ToString(ds.soNgayQuet))
                    .Replace("{$ngayCongTac}", Convert.ToString(ds.soNgayCongTac))
                    .Replace("{$ngayCongTangCa}", Convert.ToString(ds.soNgayTangCaThuong + ds.soNgayTangCaChuNhat + ds.soNgayTangCaNgayLe))
                    .Replace("{$soNgayPhepLuyKeThangTruoc}", Convert.ToString(ds.soNgayPhepLuyKeThangTruoc))

                    .Replace("{$nghiPhep}", Convert.ToString(ds.soNgayNghiPhep))
                    .Replace("{$nghiLeTet}", Convert.ToString(ds.soNgayNghiLe))

                    .Replace("{$luongTangCa}", String.Format("{0:###,##0}", (ds.luongTangCa ?? 0)))
                    .Replace("{$luongTheoCongChinh}", String.Format("{0:###,##0}", (ds.luongThucTe ?? 0) + (ds.luongNghiPhep ?? 0)))

                    .Replace("{$luongTheoNgayCong}", String.Format("{0:###,##0}", (ds.luongThucTe ?? 0) + (ds.luongNghiPhep ?? 0) + (ds.luongTangCa ?? 0)))

                    .Replace("{$congChoViec}", Convert.ToString(ds.soNgayCongChoViec))
                    .Replace("{$luongChoViec}", String.Format("{0:###,##0}", ds.luongChoViec))
                    .Replace("{$tongTienPhuCap}", String.Format("{0:###,##0}", ds.phuCapTienAn + ds.phuCapDienThoai))
                    .Replace("{$tienAnGiuaCa}", String.Format("{0:###,##0}", ds.phuCapTienAn))
                    .Replace("{$tienDienThoai}", String.Format("{0:###,##0}", ds.phuCapDienThoai))
                    .Replace("{$tienBHXH}", String.Format("{0:###,##0}", ds.baoHiem > 0 ? ds.baoHiemXH : 0))
                    .Replace("{$tienBHYT}", String.Format("{0:###,##0}", ds.baoHiem > 0 ? ds.baoHiemYTe : 0))
                    .Replace("{$tienBHTN}", String.Format("{0:###,##0}", ds.baoHiem > 0 ? ds.baoHiemTN : 0))
                    .Replace("{$tienTTNCT}", String.Format("{0:###,##0}", ds.thue))
                    .Replace("{$tienDoanPhi}", String.Format("{0:###,##0}", ds.doanPhi))
                    .Replace("{$tienDangPhi}", String.Format("{0:###,##0}", ds.dangPhi))
                    .Replace("{$tienPhuCap}", String.Format("{0:###,##0}", ds.phuCapKhac))
                    //double truyLanh = (item1.truyLanh ?? 0);
                    //    if ((item1.TTTLLuong ?? 0) > 0) truyLanh += (item1.TTTLLuong ?? 0);
                    //    if ((item1.TTTLThue ?? 0) > 0) truyLanh += (item1.TTTLThue ?? 0);

                    //    double truyThu = (item1.truyThu ?? 0);
                    //    if ((item1.TTTLLuong ?? 0) < 0) truyThu += (item1.TTTLLuong ?? 0);
                    //    if ((item1.TTTLThue ?? 0) < 0) truyThu += (item1.TTTLThue ?? 0);
                    .Replace("{$tienTruyThu}", String.Format("{0:###,##0}", ds.truyThu + ((ds.TTTLLuong ?? 0) < 0 ? ds.TTTLLuong : 0) + ((ds.TTTLThue ?? 0) < 0 ? ds.TTTLThue : 0)))
                    .Replace("{tongPhuCapTL}", String.Format("{0:###,##0}", ds.truyLanh - ds.truyThu + ds.phuCapKhac))
                    .Replace("{$tienTruyLanh}", String.Format("{0:###,##0}", ds.truyLanh + ((ds.TTTLLuong ?? 0) > 0 ? ds.TTTLLuong : 0) + ((ds.TTTLThue ?? 0) > 0 ? ds.TTTLThue : 0)))
                    .Replace("{$tienGiamTruGC}", String.Format("{0:###,##0}", (ds.giamTruBanThan ?? 0) + (ds.giamTruNguoiPhuThuoc ?? 0)))
                    .Replace("{$tongThucNhanLuong}", String.Format("{0:###,##0}", ds.thucLanh))
                    .Replace("{$tongKhauTru}", String.Format("{0:###,##0}", tongKhauTru));

            }
            ViewBag.NoiDung = noiDung;
            // return PartialView("_ViewChiTietLuong");
            return PartialView("_ViewChiTietLuongTemplate");
        }
        //public ActionResult ViewChiTietLuong(string thang, string nam)
        //{
        //    #region Role user
        //    permission = GetPermission("XemBangLuong", BangPhanQuyen.QuyenXem);
        //    if (!permission.HasValue)
        //        return View("LogIn");
        //    if (!permission.Value)
        //        return View("AccessDenied");
        //    #endregion         
        //    var dsMauIn = nhanSuContext.GetTable<BatDongSan.Models.HeThong.Sys_PrintTemplate>().Where(d => d.maMauIn == "MIBLNV").FirstOrDefault();
        //    string noiDung = string.Empty;
        //    var ds = nhanSuContext.tbl_NS_BangLuongNhanViens
        //                       .Where(t => t.maNhanVien == GetUser().manv && t.thang.ToString() == thang && nam == t.nam.ToString()).FirstOrDefault();
        //    ViewData["chiTiet"] = dsMauIn;
        //    if (dsMauIn != null)
        //    {
        //        var tongKhauTru = (ds.baoHiemXH ?? 0) + (ds.baoHiemYTe ?? 0) + (ds.baoHiemTN ?? 0) + (ds.thue ?? 0) + (ds.doanPhi ?? 0) + (ds.dangPhi ?? 0);
        //        noiDung = dsMauIn.html.Replace("{$thang}", thang)
        //            .Replace("{$nam}", nam)
        //            .Replace("{$hoVaTen}", ds.hoTen)
        //            .Replace("{$tongLuong}", String.Format("{0:###,##0}", ds.tongLuong))
        //            .Replace("{$khoanBSLVSLuongBH}", String.Format("{0:###,##0}", ds.luongDongBaoHiem))
        //            .Replace("{$luongThoaThuan}", String.Format("{0:###,##0}", ds.khoanBoSungLuong))
        //            .Replace("{$ngayCongChuan}", Convert.ToString(ds.ngayCongChuan))
        //            .Replace("{$ngayCongTinhLuong}", Convert.ToString(ds.tongNgayCong))
        //            .Replace("{$ngayCong}", Convert.ToString(ds.soNgayQuet))
        //            .Replace("{$ngayCongTac}", Convert.ToString(ds.soNgayCongTac))
        //            .Replace("{$nghiPhep}", Convert.ToString(ds.soNgayNghiPhep))
        //            .Replace("{$nghiLeTet}", Convert.ToString(ds.soNgayNghiLe))
        //            .Replace("{$luongTheoNgayCong}", String.Format("{0:###,##0}", ds.luongThang))
        //            .Replace("{$congChoViec}", Convert.ToString(ds.soNgayCongChoViec))
        //            .Replace("{$luongChoViec}", String.Format("{0:###,##0}", ds.luongChoViec))
        //            .Replace("{$tongTienPhuCap}", String.Format("{0:###,##0}", ds.phuCapTienAn + ds.phuCapDienThoai))
        //            .Replace("{$tienAnGiuaCa}", String.Format("{0:###,##0}", ds.phuCapTienAn))
        //            .Replace("{$tienDienThoai}", String.Format("{0:###,##0}", ds.phuCapDienThoai))
        //            .Replace("{$tienBHXH}", String.Format("{0:###,##0}", ds.baoHiemXH))
        //            .Replace("{$tienBHYT}", String.Format("{0:###,##0}", ds.baoHiemYTe))
        //            .Replace("{$tienBHTN}", String.Format("{0:###,##0}", ds.baoHiemTN))
        //            .Replace("{$tienTTNCT}", String.Format("{0:###,##0}", ds.thue))
        //            .Replace("{$tienDoanPhi}", String.Format("{0:###,##0}", ds.doanPhi))
        //            .Replace("{$tienDangPhi}", String.Format("{0:###,##0}", ds.dangPhi))
        //            .Replace("{$tienPhuCap}", String.Format("{0:###,##0}", ds.phuCapKhac))
        //            .Replace("{$tienTruyThu}", String.Format("{0:###,##0}", ds.truyThu))
        //            .Replace("{tongPhuCapTL}", String.Format("{0:###,##0}", ds.truyLanh-ds.truyThu+ds.phuCapKhac))
        //            .Replace("{$tienTruyLanh}", String.Format("{0:###,##0}", ds.truyLanh))
        //            .Replace("{$tienGiamTruGC}", String.Format("{0:###,##0}", (ds.giamTruBanThan ?? 0) + (ds.giamTruNguoiPhuThuoc ?? 0)))
        //            .Replace("{$tongThucNhanLuong}", String.Format("{0:###,##0}", ds.thucLanh))
        //            .Replace("{$tongKhauTru}", String.Format("{0:###,##0}", tongKhauTru));

        //    }
        //    ViewBag.NoiDung = noiDung;
        //    // return PartialView("_ViewChiTietLuong");
        //    return PartialView("_ViewChiTietLuongTemplate");
        //}
        public ActionResult LoadXemBangLuong(int nam, int _page = 0)
        {
            #region Role user
            permission = GetPermission("XemBangLuong", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            string maNhanVien =GetUser().manv;
            int page = _page == 0 ? 1 : _page;
            int pIndex = page;
            int total = nhanSuContext.sp_NS_BangLuongDanhChoNhanVien(maNhanVien, nam).Count();
            PagingLoaderController("/ChamCong/XemBangLuong/", total, page, "?maNhanVien=" + maNhanVien);
            ViewData["lsDanhSach"] = nhanSuContext.sp_NS_BangLuongDanhChoNhanVien(maNhanVien, nam).Skip(start).Take(offset).ToList();

            return PartialView("_LoadXemBangLuong");
        }
        public ActionResult XemBangChamCongChiTiet()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            buildTree = new StringBuilder();
            phongBans = nhanSuContext.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan>().ToList();
            buildTree = TreePhongBanAjax.BuildTreeDepartment(phongBans);
            ViewBag.PhongBans = buildTree.ToString();
            thang(DateTime.Now.Month);
            nam(DateTime.Now.Year);
            return View("");
        }
        public ActionResult LoadBangChamCongChiTiet(string qSearch, int thang, int nam, int _page = 0)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            string maNhanVien = GetUser().manv;
            int page = _page == 0 ? 1 : _page;
            int pIndex = page;
            int total = nhanSuContext.sp_NS_BangChamCongChiTiet(thang, nam, qSearch, maNhanVien).Count();
            PagingLoaderController("/ChamCong/LoadBangChamCongChiTiet/", total, page, "?qsearch=" + qSearch);
            ViewData["lsDanhSach"] = nhanSuContext.sp_NS_BangChamCongChiTiet(thang, nam, qSearch, maNhanVien).Skip(start).Take(offset).ToList();

            ViewData["qSearch"] = qSearch;
            return PartialView("_LoadBangChamCongChiTiet");
        }





        public ActionResult LoadBangChamCongTongHop(string qSearch, string maPhongBan, int thang, int nam, int _page = 0)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            int page = _page == 0 ? 1 : _page;
            int pIndex = page;
          //  string maNhanVien = GetUser().manv;
            string maNhanVien = GetUser().manv;
            int total = nhanSuContext.sp_NS_BangTongHopCongThang(thang, nam, qSearch, maNhanVien, 0, maPhongBan).Count();
            PagingLoaderController("/BangChamCongTongHop/LoadBangChamCongTongHop/", total, page, "?qsearch=" + qSearch);
            ViewData["lsDanhSach"] = nhanSuContext.sp_NS_BangTongHopCongThang(thang, nam, qSearch, maNhanVien, 0, maPhongBan).Skip(start).Take(offset).ToList();

            ViewData["qSearch"] = qSearch;
            return PartialView("_LoadBangChamCongTongHop");
        }
        private void thang(int value)
        {
            Dictionary<int, string> dics = new Dictionary<int, string>();
            for (int i = 0; i < 13; i++)
            {
                dics[i] = i.ToString();
            }
            ViewData["thang"] = new SelectList(dics, "Key", "Value", value);
            ViewData["thangtc"] = new SelectList(dics, "Key", "Value", value);
        }
        private void nam(int value)
        {
            Dictionary<int, string> dics = new Dictionary<int, string>();
            for (int i = 2015; i < 2031; i++)
            {
                dics[i] = i.ToString();
            }
            ViewData["nam"] = new SelectList(dics, "Key", "Value", value);
            ViewData["namtc"] = new SelectList(dics, "Key", "Value", value);
        }


    }
}
