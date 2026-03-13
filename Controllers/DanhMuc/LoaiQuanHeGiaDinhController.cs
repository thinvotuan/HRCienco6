using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Models.DanhMuc;

namespace BatDongSan.Controllers.DanhMuc
{
    public class LoaiQuanHeGiaDinhController : ApplicationController
    {
        private LinqDanhMucDataContext context;
        private IList<tbl_DM_LoaiQuanHeGiaDinh> loaiQuanHes;
        private tbl_DM_LoaiQuanHeGiaDinh loaiQuanHe;
        private readonly string MCV = "LoaiQuanHeGiaDinh";
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
                loaiQuanHes = context.tbl_DM_LoaiQuanHeGiaDinhs.ToList();
            }
            return View(loaiQuanHes);
        }

        //
        // GET: /LoaiQuanHeGiaDinh/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /LoaiQuanHeGiaDinh/Create

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            loaiQuanHe = new tbl_DM_LoaiQuanHeGiaDinh();
            return PartialView("Create", loaiQuanHe);
        }

        //
        // POST: /LoaiQuanHeGiaDinh/Create

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
                    loaiQuanHe = new tbl_DM_LoaiQuanHeGiaDinh();
                    loaiQuanHe.ghiChu = collection["ghiChu"];
                    loaiQuanHe.tenLoaiQuanHe = collection["tenLoaiQuanHe"];
                    loaiQuanHe.nguoiLap = GetUser().manv;
                    loaiQuanHe.ngayLap = DateTime.Now;
                    context.tbl_DM_LoaiQuanHeGiaDinhs.InsertOnSubmit(loaiQuanHe);
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
        // GET: /LoaiQuanHeGiaDinh/Edit/5

        public ActionResult Edit(int id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            using (context = new LinqDanhMucDataContext())
            {
                loaiQuanHe = context.tbl_DM_LoaiQuanHeGiaDinhs.Where(s => s.id == id).FirstOrDefault();
            }
            return PartialView("Edit", loaiQuanHe);
        }

        //
        // POST: /LoaiQuanHeGiaDinh/Edit/5

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
                    loaiQuanHe = context.tbl_DM_LoaiQuanHeGiaDinhs.Where(s => s.id == id).FirstOrDefault();
                    loaiQuanHe.ghiChu = collection["ghiChu"];
                    loaiQuanHe.tenLoaiQuanHe = collection["tenLoaiQuanHe"];
                    loaiQuanHe.ngayCapNhat = DateTime.Now;
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
        // POST: /LoaiQuanHeGiaDinh/Delete/5

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
                    loaiQuanHe = context.tbl_DM_LoaiQuanHeGiaDinhs.Where(s => s.id == id).FirstOrDefault();
                    context.tbl_DM_LoaiQuanHeGiaDinhs.DeleteOnSubmit(loaiQuanHe);
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
