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
using Worldsoft.Mvc.Web.Util;
using NPOI.HSSF.UserModel.Contrib;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using BatDongSan.Models.VIP;

namespace BatDongSan.Controllers.TinhCong
{
    public class DieuChinhCongVipController : ApplicationController
    {
        private LinqVIPDataContext vipContext = new LinqVIPDataContext();
        private StringBuilder buildTree;
        private readonly string MCV = "DieuChinhCongVip";
        private bool? permission;
        //
        // GET: /DieuChinhCongVip/

        public ActionResult Index()
        {
            return View();
        }

        //
        // GET: /DieuChinhCongVip/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /DieuChinhCongVip/Create

        public ActionResult Create()
        {
            return View();
        }

        //
        // POST: /DieuChinhCongVip/Create

        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {
                // TODO: Add insert logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /DieuChinhCongVip/Edit/5

        public ActionResult Edit(int id)
        {
            return View();
        }

        //
        // POST: /DieuChinhCongVip/Edit/5

        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /DieuChinhCongVip/Delete/5

        public ActionResult Delete(int id)
        {
            return View();
        }

        //
        // POST: /DieuChinhCongVip/Delete/5

        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
        //Admin
        public ActionResult XemBangChamCongTongHop()
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
        public ActionResult LoadBangChamCongTongHop(string qSearch, int thang, int nam, int? trangThai, int _page = 0)
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
            int total = vipContext.sp_XemDieuChinhCongVIP(thang, nam, qSearch, "", trangThai).Count();
            PagingLoaderFullController("/DieuChinhCongVip/LoadBangChamCongTongHop/", total, page, "?thang=" + thang + "&nam=" + nam + "&qSearch=" + qSearch + "&trangThai=" + trangThai + "&maNhanVien=" + qSearch);
            ViewData["lsDanhSach"] = vipContext.sp_XemDieuChinhCongVIP(thang, nam, qSearch, "", trangThai).Skip(start).Take(1000).ToList();

            ViewData["qSearch"] = qSearch;
            return PartialView("_LoadBangChamCongTongHop");
        }
        public string DuyetNhieuPhieu(string maNhanVien, int thang, int nam)
        {
            try
            {
                #region Role user
                permission = GetPermission(MCV, BangPhanQuyen.QuyenDuyet);
                if (!permission.HasValue)
                    return "false";
                if (!permission.Value)
                    return "false";
                #endregion
               
                        var dieuChinhCong = vipContext.tbl_BangLuongVIPs.Where(d => d.thang == thang && d.nam == nam && d.maNhanVien == maNhanVien.Trim()).FirstOrDefault();
                        dieuChinhCong.trangThai = 1;    
                        vipContext.SubmitChanges();
                        
                 
                return "true";
            }
            catch
            {
                return "false";
            }
        }
        public string TraNhieuPhieu(string maNhanVien, int thang, int nam)
        {
            try
            {
                #region Role user
                permission = GetPermission(MCV, BangPhanQuyen.QuyenDuyet);
                if (!permission.HasValue)
                    return "false";
                if (!permission.Value)
                    return "false";
                #endregion
                
                        var dieuChinhCong = vipContext.tbl_BangLuongVIPs.Where(d => d.thang == thang && d.nam == nam && d.maNhanVien ==maNhanVien.Trim()).FirstOrDefault();
                        dieuChinhCong.trangThai = 0;
                        dieuChinhCong.guiDuyet = 0;
                        vipContext.SubmitChanges();
                        
                   
                return "true";
            }
            catch
            {
                return "false";
            }
        }
        
        //End Admin
        //Nhan vien
        public ActionResult XemBangChamCong()
        {

            #region Role user
            permission = GetPermission("DieuChinhCongVipNV", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
           
            nam(DateTime.Now.Year);
            return View("");
        }
        public ActionResult LoadBangChamCong(string qSearch, int nam, int? trangThai, int _page = 0)
        {
            #region Role user
            permission = GetPermission("DieuChinhCongVipNV", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            int page = _page == 0 ? 1 : _page;
            int pIndex = page;
            int total = vipContext.sp_XemDieuChinhCongVIP(null, nam, qSearch, GetUser().manv, trangThai).Where(d=>d.maNhanVien == GetUser().manv).Count();
            PagingLoaderFullController("/DieuChinhCongVip/LoadBangChamCong/", total, page, "?nam=" + nam + "&qSearch=" + qSearch + "&trangThai=" + trangThai + "&maNhanVien=" + GetUser().manv);
            ViewData["lsDanhSach"] = vipContext.sp_XemDieuChinhCongVIP(null, nam, qSearch, GetUser().manv, trangThai).Where(d => d.maNhanVien == GetUser().manv).Skip(start).Take(1000).ToList();

            ViewData["qSearch"] = qSearch;
            return PartialView("_LoadBangChamCong");
        }
        public JsonResult GetCongChuan(int thang, int nam, float congDieuChinh, int change, string maNhanVien)
        {
            double tongCongSauKhiDC = 0;
            double giaTriMax = 0;
            int trangThaiGet = 0;
            int chapNhan = 0;
            string VuotCongChuan = string.Empty;
            var MaNV = maNhanVien;
            if(maNhanVien == null || maNhanVien == ""){
                MaNV = GetUser().manv;
            }

            var thongTin = vipContext.tbl_BangLuongVIPs.Where(d => d.maNhanVien == MaNV && d.thang == thang && d.nam == nam).FirstOrDefault();
            if (thongTin != null) {
                trangThaiGet = 1;

                var tongCongMoi = Math.Round((thongTin.ngayCong ?? 0) + (thongTin.ngayNghiPhep ?? 0)
                    + (thongTin.ngayNghiLeTet??0) + (thongTin.ngayLuyKeThangTruoc??0) + congDieuChinh - (thongTin.nghiKhongLuong ??0),1);
                if (tongCongMoi <= thongTin.ngayCongChuan)
                {
                    chapNhan = 1;
                    if (change == 1)
                    {
                        thongTin.congDieuChinh = Math.Round(congDieuChinh, 1);

                        thongTin.tongCongSauKhiDC = Math.Round(tongCongMoi, 1);
                        vipContext.SubmitChanges();
                    }
                    tongCongSauKhiDC = Convert.ToDouble(tongCongMoi);
                }
                else {
                    chapNhan = 0;
                    giaTriMax = Math.Round(((thongTin.ngayCongChuan??0)-(Math.Round((thongTin.ngayCong ?? 0) + (thongTin.ngayNghiPhep ?? 0)
                    + (thongTin.ngayNghiLeTet??0) + (thongTin.ngayLuyKeThangTruoc??0) - (thongTin.nghiKhongLuong ??0),1))),1);
                }
            }
            var result = new { congDieuChinh = congDieuChinh, trangThaiGet = trangThaiGet, chapNhan = chapNhan, tongCongSauKhiDC = tongCongSauKhiDC, giaTriMax = giaTriMax, ngayCongChuan = thongTin.ngayCongChuan };
            return Json(result, JsonRequestBehavior.AllowGet);
        }
        public JsonResult CapNhatLyDo(int thang, int nam, string textLyDo, string maNhanVien)
        {
            int chapNhan = 0;
            var MaNV = maNhanVien;
            if (maNhanVien == null || maNhanVien == "")
            {
                MaNV = GetUser().manv;
            }

            var thongTin = vipContext.tbl_BangLuongVIPs.Where(d => d.maNhanVien == MaNV && d.thang == thang && d.nam == nam).FirstOrDefault();
            if (thongTin != null)
            {
                thongTin.lyDoXinDC = textLyDo;
                chapNhan = 1;
                vipContext.SubmitChanges();
                
            }
            var result = new {  chapNhan = chapNhan};
            return Json(result, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GuiDuyet(int thang, int nam)
        {
            int chapNhan = 0;
            var thongTin = vipContext.tbl_BangLuongVIPs.Where(d => d.maNhanVien == GetUser().manv && d.thang == thang && d.nam == nam).FirstOrDefault();
            if (thongTin != null)
            {
                thongTin.guiDuyet = 1;
                chapNhan = 1;
                vipContext.SubmitChanges();

            }
            var result = new { chapNhan = chapNhan };
            return Json(result, JsonRequestBehavior.AllowGet);
        }

        // End nhan vien
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
