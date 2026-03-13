using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Models.DanhMuc;

namespace BatDongSan.Controllers.DanhMuc
{
    public class QuanLyCapBacController : ApplicationController
    {
        private LinqDanhMucDataContext context;
        private IList<tbl_DM_CapBacChucDanh> chiNhanhs;
        private tbl_DM_CapBacChucDanh DMCapBac;
        private readonly string MCV = "QuanLyCapBac";
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
                chiNhanhs = context.tbl_DM_CapBacChucDanhs.ToList();
            }
            return View(chiNhanhs);
        }

        //
        // GET: /QuanLyCapBac/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /QuanLyCapBac/Create

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            DMCapBac = new tbl_DM_CapBacChucDanh();
            return PartialView("Create", DMCapBac);
        }

        //
        // POST: /QuanLyCapBac/Create

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
                    int id = context.tbl_DM_CapBacChucDanhs.OrderByDescending(d => d.id).FirstOrDefault().id;
                    if (id == null) id = 0;
                    id = id + 1;
                    DMCapBac = new tbl_DM_CapBacChucDanh();
                    DMCapBac.soCapBac =Convert.ToInt32(collection["soCapBac"]);
                    DMCapBac.tenCapBac = collection["tenCapBac"];
                    DMCapBac.moTa = collection["moTa"];
                    DMCapBac.id = id;
                    context.tbl_DM_CapBacChucDanhs.InsertOnSubmit(DMCapBac);
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
        // GET: /QuanLyCapBac/Edit/5

        public ActionResult Edit(int? id)
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
                DMCapBac = context.tbl_DM_CapBacChucDanhs.Where(s => s.id == id).FirstOrDefault();
            }
            return PartialView("Edit", DMCapBac);
        }

        //
        // POST: /QuanLyCapBac/Edit/5

        [HttpPost]
        public ActionResult Edit(int? id, FormCollection collection)
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
                    DMCapBac = context.tbl_DM_CapBacChucDanhs.Where(s => s.id == id).FirstOrDefault();
                    DMCapBac.soCapBac = Convert.ToInt32(collection["soCapBac"]);
                    DMCapBac.tenCapBac = collection["tenCapBac"];
                    DMCapBac.moTa = collection["moTa"];
                   
                   
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
        // POST: /QuanLyCapBac/Delete/5

        [HttpPost]
        public ActionResult Delete(int? id)
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
                    DMCapBac = context.tbl_DM_CapBacChucDanhs.Where(s => s.id == id).FirstOrDefault();
                    context.tbl_DM_CapBacChucDanhs.DeleteOnSubmit(DMCapBac);
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
