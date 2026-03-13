using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Models.NhanSu;

namespace BatDongSan.Controllers.NhanSu
{
    public class TyLeDongBHController : ApplicationController
    {
        private LinqNhanSuDataContext context;
        private IList<tbl_NS_TyLeDongBH> baoHiems;
        private tbl_NS_TyLeDongBH baoHiem;
        private readonly string MCV = "TyLeDongBH";
        private bool? permission;

        public ActionResult Index()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            context = new LinqNhanSuDataContext();
            baoHiems = context.tbl_NS_TyLeDongBHs.ToList();
            return View(baoHiems);
        }

        //
        // GET: /TyLeDongBHXH/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /TyLeDongBHXH/Create

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion


            baoHiem = new tbl_NS_TyLeDongBH();
            return PartialView("Create", baoHiem);
        }

        //
        // POST: /TyLeDongBHXH/Create

        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion


            try
            {
                context = new LinqNhanSuDataContext();
                baoHiem = new tbl_NS_TyLeDongBH();
                baoHiem.baoHiemTNDoanhNghiep = Convert.ToDecimal(collection["baoHiemTNDoanhNghiep"]);
                baoHiem.baoHiemTNNhanVien = Convert.ToDecimal(collection["baoHiemTNNhanVien"]);
                baoHiem.baoHiemXHDoanhNghiep = Convert.ToDecimal(collection["baoHiemXHDoanhNghiep"]);
                baoHiem.baoHiemXHNhanVien = Convert.ToDecimal(collection["baoHiemXHNhanVien"]);
                baoHiem.baoHiemYTeDoanhNghiep = Convert.ToDecimal(collection["baoHiemYTeDoanhNghiep"]);
                baoHiem.baoHiemYTeNhanVien = Convert.ToDecimal(collection["baoHiemYTeNhanVien"]);
                baoHiem.ghiChu = collection["ghiChu"];
                baoHiem.ngayApDung = DateTime.ParseExact(collection["ngayApDung"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                baoHiem.ngayLap = DateTime.Now;
                baoHiem.nguoiLap = GetUser().manv;
                context.tbl_NS_TyLeDongBHs.InsertOnSubmit(baoHiem);
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /TyLeDongBHXH/Edit/5

        public ActionResult Edit(int id)
        {
            return View();
        }

        //
        // POST: /TyLeDongBHXH/Edit/5

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
        // GET: /TyLeDongBHXH/Delete/5

        public ActionResult Delete(int id)
        {
            return View();
        }

        //
        // POST: /TyLeDongBHXH/Delete/5

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
    }
}
