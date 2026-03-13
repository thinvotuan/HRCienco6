using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data.OleDb;
using System.Data.SqlClient;
using BatDongSan.Models.HeThong;
using BatDongSan.Utils.Paging;
using BatDongSan.Helper.Common;
using System.IO;
using BatDongSan.Helper.Utils;
using System.Text.RegularExpressions;
using System.Data;

namespace BatDongSan.Controllers.HeThong
{
    public class GhiNhanLogController : ApplicationController
    {
        private LinqHeThongDataContext contextHT = new LinqHeThongDataContext();
     
       
        private bool? permission;

        public ActionResult Index()
        {
            #region Role user
            permission = GetPermission("GhiNhanLog", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            return View("");


        }

        public ActionResult GetTimeLine(string tuNgay, string denNgay, string qsearch, int? maTL = null)
        {
            try
            {

                DateTime? fromDate = String.IsNullOrEmpty(tuNgay) ? (DateTime?)null : DateTime.ParseExact(tuNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                DateTime? toDate = String.IsNullOrEmpty(denNgay) ? (DateTime?)null : DateTime.ParseExact(denNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                var list = contextHT.sp_Sys_LichSuHoatDong_Index(maTL, fromDate, toDate, qsearch).ToList();

                if (list.Count == 0)
                {
                    list = null;
                }
                return Json(list, JsonRequestBehavior.AllowGet);
            }
            catch
            {
                return Json(string.Empty, JsonRequestBehavior.AllowGet);
            }


        }

        //
        // GET: /SanPham/Details/5

        public ActionResult Details(int id)
        {
            #region Role user
            permission = GetPermission("GhiNhanLog", BangPhanQuyen.QuyenXemChiTiet);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            return View("");
        }



    }
}
