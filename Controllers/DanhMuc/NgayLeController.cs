using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Models.DanhMuc;

namespace BatDongSan.Controllers.DanhMuc
{
    public class NgayLeController : ApplicationController
    {
        private LinqDanhMucDataContext context;
        private IList<tbl_DM_NghiLe> ngayLes;
        private tbl_DM_NghiLe ngayLe;
        private readonly string MCV = "NgayLe";
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

            using (context = new LinqDanhMucDataContext())
            {
                ngayLes = context.tbl_DM_NghiLes.ToList();
            }
            return View(ngayLes);
        }

        //
        // GET: /NgayLe/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /NgayLe/Create

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            context = new LinqDanhMucDataContext();
            ngayLe = new tbl_DM_NghiLe();
            ngayLe.nam = Convert.ToInt16(DateTime.Now.Year);
            ngayLe.loaiNgayNghi = true;
            var nams = context.Sys_Nams;
            ViewBag.Nams = new SelectList(nams, "id", "nam", ngayLe.nam);

            return PartialView("Create", ngayLe);
        }

        //
        // POST: /NgayLe/Create

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
                using (context = new LinqDanhMucDataContext())
                {
                    ngayLe = new tbl_DM_NghiLe();
                    ngayLe.ngayNghiLe = String.IsNullOrEmpty(collection["ngayNghiLe"]) ? (DateTime?)null : DateTime.ParseExact(collection["ngayNghiLe"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    ngayLe.ngayDuongLich = String.IsNullOrEmpty(collection["ngayDuongLich"]) ? (DateTime?)null : DateTime.ParseExact(collection["ngayDuongLich"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    ngayLe.tenNgayNghiLe = collection["tenNgayNghiLe"];
                    ngayLe.loaiNgayNghi = collection["loaiNgayNghi"] == "true" ? true : false;
                    ngayLe.nam = Convert.ToInt16(collection["nam"]);
                    ngayLe.ngayAmLich = String.IsNullOrEmpty(collection["ngayAmLich"]) ? (DateTime?)null : DateTime.ParseExact(collection["ngayAmLich"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    ngayLe.ghiChu = collection["ghiChu"];
                    ngayLe.nguoiLap = GetUser().manv;
                    ngayLe.ngayLap = DateTime.Now;
                    context.tbl_DM_NghiLes.InsertOnSubmit(ngayLe);
                    context.SubmitChanges();
                }
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /NgayLe/Edit/5

        public ActionResult Edit(int id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            context = new LinqDanhMucDataContext();
            ngayLe = context.tbl_DM_NghiLes.Where(s => s.id == id).FirstOrDefault();            
            var nams = context.Sys_Nams;
            ViewBag.Nams = new SelectList(nams, "id", "nam", ngayLe.nam);

            return PartialView("Edit", ngayLe);            
        }

        //
        // POST: /NgayLe/Edit/5

        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            try
            {
                using (context = new LinqDanhMucDataContext())
                {
                    ngayLe = context.tbl_DM_NghiLes.Where(s => s.id == id).FirstOrDefault();
                    ngayLe.ngayNghiLe = String.IsNullOrEmpty(collection["ngayNghiLe"]) ? (DateTime?)null : DateTime.ParseExact(collection["ngayNghiLe"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    ngayLe.ngayDuongLich = String.IsNullOrEmpty(collection["ngayDuongLich"]) ? (DateTime?)null : DateTime.ParseExact(collection["ngayDuongLich"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    ngayLe.tenNgayNghiLe = collection["tenNgayNghiLe"];
                    ngayLe.loaiNgayNghi = collection["loaiNgayNghi"] == "true" ? true : false;
                    ngayLe.nam = Convert.ToInt16(collection["nam"]);
                    ngayLe.ngayAmLich = String.IsNullOrEmpty(collection["ngayAmLich"]) ? (DateTime?)null : DateTime.ParseExact(collection["ngayAmLich"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    ngayLe.ghiChu = collection["ghiChu"];                    
                    ngayLe.ngayCapNhat = DateTime.Now;
                    context.SubmitChanges();
                }

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // POST: /NgayLe/Delete/5

        [HttpPost]
        public ActionResult Delete(int id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXoa);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            try
            {
                using (context = new LinqDanhMucDataContext())
                {
                    ngayLe = context.tbl_DM_NghiLes.Where(s => s.id == id).FirstOrDefault();
                    context.tbl_DM_NghiLes.DeleteOnSubmit(ngayLe);
                    context.SubmitChanges();
                }

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
    }
}
