using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Utils;
using BatDongSan.Models.NhanSu;
using System.Text;
using BatDongSan.Utils.Paging;
using BatDongSan.Helper.Common;
using System.Globalization;
using System.Configuration;
using System.Net.Mail;

namespace BatDongSan.Controllers.BaoCao
{
    public class BaoCaoTongHopController : ApplicationController
    {
        private LinqNhanSuDataContext context = new LinqNhanSuDataContext();
        private readonly string MCV = "BaoCaoTongHop";
        private bool? permission;
        //
        // GET: /BaoCaoNhanSu/

        public ActionResult Index()
        {
            return View();
        }
        private void thang(int value)
        {
            Dictionary<int, string> dics = new Dictionary<int, string>();
            for (int i = 0; i < 13; i++)
            {
                dics[i] = i.ToString();
            }
            ViewData["lstThangFrom"] = new SelectList(dics, "Key", "Value", value);
            ViewData["lstThangTo"] = new SelectList(dics, "Key", "Value", value);

        }
        private void nam(int value)
        {
            Dictionary<int, string> dics = new Dictionary<int, string>();
            for (int i = 2015; i < 2031; i++)
            {
                dics[i] = i.ToString();
            }
            ViewData["lstNamFrom"] = new SelectList(dics, "Key", "Value", value);
            ViewData["lstNamTo"] = new SelectList(dics, "Key", "Value", value);

        }
        public ActionResult BieuDoTongSoNV()
        {

            #region Role user
            permission = GetPermission("BaoCaoTongHop", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
         
            return PartialView("_BieuDoTongSoNV");

        }
        
        public ActionResult GetListBieuDoTongSoNV()
        {
          try
            {

                var list = context.sp_BC_NS_TongHop_TongSoNhanVien().ToList();
                var listThang = context.sp_BC_NS_TongHop_TuyenMoiNghiViec(1).ToList();
                var listNam = context.sp_BC_NS_TongHop_TuyenMoiNghiViec(0).ToList();


                if (list.Count == 0)
                {
                    list = null;
                }
                var result = new { list = list, listThang = listThang, listNam = listNam };
                return Json(result, JsonRequestBehavior.AllowGet);

            }
            catch
            {
                return View();
            }

        }
        public ActionResult BieuDoTuyenMoiNghiViec()
        {

            #region Role user
            permission = GetPermission("BaoCaoTongHop", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            return PartialView("_BieuDoTuyenMoiNghiViec");

        }
        public ActionResult GetListBieuDoTuyenMoiNghiViec(int? thang)
        {
            try
            {

                var list = context.sp_BC_NS_TongHop_TuyenMoiNghiViec(thang).ToList();



                if (list.Count == 0)
                {
                    list = null;
                }
                var result = new { list = list };
                return Json(result, JsonRequestBehavior.AllowGet);

            }
            catch
            {
                return View();
            }

        }
        public ActionResult BieuDoLuong()
        {

            #region Role user
            permission = GetPermission("BaoCaoTongHop", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            return PartialView("_BieuDoLuong");

        }
        public ActionResult GetListBieuDoLuong()
        {
            try
            {

                var list = context.sp_BC_NS_TongHop_BieuDoLuong().ToList();



                if (list.Count == 0)
                {
                    list = null;
                }
                var result = new { list = list };
                return Json(result, JsonRequestBehavior.AllowGet);

            }
            catch
            {
                return View();
            }

        }
        public ActionResult BieuDoTuoi()
        {

            #region Role user
            permission = GetPermission("BaoCaoTongHop", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            return PartialView("_BieuDoTuoi");

        }
        public ActionResult GetListBieuDoTuoi()
        {
            try
            {

                var list = context.sp_BC_NS_TongHop_DoTuoi().ToList();



                if (list.Count == 0)
                {
                    list = null;
                }
                var result = new { list = list };
                return Json(result, JsonRequestBehavior.AllowGet);

            }
            catch
            {
                return View();
            }

        }
        public ActionResult BieuDoTrinhDo()
        {

            #region Role user
            permission = GetPermission("BaoCaoTongHop", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            return PartialView("_BieuDoTrinhDo");

        }
        public ActionResult GetListBieuDoTrinhDo()
        {
            try
            {

                var list = context.sp_BC_NS_TongHop_TrinhDo().ToList();



                if (list.Count == 0)
                {
                    list = null;
                }
                var result = new { list = list };
                return Json(result, JsonRequestBehavior.AllowGet);

            }
            catch
            {
                return View();
            }

        }
       

        public ActionResult BieuDoHopDongSapHetHan()
        {

            #region Role user
            permission = GetPermission("BaoCaoTongHop", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            return PartialView("_BieuDoHopDongSapHetHan");
        }
        public ActionResult LoadBieuDoHopDongSapHetHan()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            int total = context.sp_BC_NS_TongHop_HopDongSapDenHan().Count();
         
            ViewData["lsDanhSach"] = context.sp_BC_NS_TongHop_HopDongSapDenHan().ToList();
            return PartialView("_LoadBieuDoHopDongSapHetHan");
        }
        public ActionResult BieuDoBienDongNVThang()
        {

            #region Role user
            permission = GetPermission("BaoCaoTongHop", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            return PartialView("_BieuDoBienDongNVThang");

        }
        public ActionResult GetBieuDoBienDongNVThang()
        {
            try
            {

                var list = context.sp_BC_NS_TongHop_TrinhDo().ToList();



                if (list.Count == 0)
                {
                    list = null;
                }
                var result = new { list = list };
                return Json(result, JsonRequestBehavior.AllowGet);

            }
            catch
            {
                return View();
            }

        }
        

    }
}
