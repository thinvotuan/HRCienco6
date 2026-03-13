using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Models.DanhMuc;

namespace BatDongSan.Controllers.DanhMuc
{
    public class BangCapController : ApplicationController
    {
        private LinqDanhMucDataContext context;
        private IList<tbl_DM_BangCap> bangCaps;
        private tbl_DM_BangCap bangCap;
        private readonly string MCV = "BangCap";
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
                bangCaps = context.tbl_DM_BangCaps.ToList();
            }
            return View(bangCaps);
        }

        //
        // GET: /BangCap/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /BangCap/Create

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            bangCap = new tbl_DM_BangCap();
            return PartialView("Create", bangCap);
        }

        //
        // POST: /BangCap/Create

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
                bangCap = new tbl_DM_BangCap();
                using (context = new LinqDanhMucDataContext())
                {
                    bangCap.tenBangCap = collection["tenBangCap"];
                    bangCap.ghiChu = collection["ghiChu"];                    
                    context.tbl_DM_BangCaps.InsertOnSubmit(bangCap);
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
        // GET: /BangCap/Edit/5

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
                bangCap = context.tbl_DM_BangCaps.Where(s => s.id == id).FirstOrDefault();
            }
            return PartialView("Edit", bangCap);
        }

        //
        // POST: /BangCap/Edit/5

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
                    bangCap = context.tbl_DM_BangCaps.Where(s => s.id == id).FirstOrDefault();
                    bangCap.tenBangCap = collection["tenBangCap"];
                    bangCap.ghiChu = collection["ghiChu"];                    
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
        // POST: /BangCap/Delete/5

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
                    bangCap = context.tbl_DM_BangCaps.Where(s => s.id == id).FirstOrDefault();
                    context.tbl_DM_BangCaps.DeleteOnSubmit(bangCap);
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
