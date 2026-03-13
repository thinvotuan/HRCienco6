using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Models.DanhMuc;
using BatDongSan.Models.NhanSu;

namespace BatDongSan.Controllers.DanhMuc
{
    public class LyDoThoiViecController : ApplicationController
    {
        private LinqDanhMucDataContext context;
        private IList<LyDoThoiViecModel> lyDoThoiViecs;
        private tbl_DM_LyDoThoiViec lyDoThoiViec;
        private readonly string MCV = "LyDoThoiViec";
        private bool? permission;
        private string tenNguoiLap(string maNhanVien)
        {
            try
            {
                LinqNhanSuDataContext lq = new LinqNhanSuDataContext();
                string abc= lq.tbl_NS_NhanViens.Where(x => x.maNhanVien == maNhanVien).Select(x => x.ho + " " + x.ten).FirstOrDefault();
                return abc != null ? abc : maNhanVien;
            }
            catch
            {
                return maNhanVien;
            }
        }
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
                lyDoThoiViecs = context.tbl_DM_LyDoThoiViecs.Select(t => new LyDoThoiViecModel
                {
                        tenLyDoThoiViec = t.tenLyDoThoiViec ,
                        ngayLap = t.ngayLap ,
                        nguoiLap = tenNguoiLap(t.nguoiLap),
                        maLyDoThoiViec =t.maLyDoThoiViec 
                }).ToList();
            }
            return View(lyDoThoiViecs);
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
            lyDoThoiViec = new tbl_DM_LyDoThoiViec();
            return PartialView("Create", lyDoThoiViec);
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
                    lyDoThoiViec = new tbl_DM_LyDoThoiViec();
                    lyDoThoiViec.tenLyDoThoiViec = collection["tenLyDoThoiViec"];
                    lyDoThoiViec.nguoiLap = GetUser().manv;
                    lyDoThoiViec.ngayLap = DateTime.Now;
                    context.tbl_DM_LyDoThoiViecs.InsertOnSubmit(lyDoThoiViec);
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
            lyDoThoiViec = context.tbl_DM_LyDoThoiViecs.Where(s => s.maLyDoThoiViec == id).FirstOrDefault();

            return PartialView("Edit", lyDoThoiViec);            
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
                    lyDoThoiViec = context.tbl_DM_LyDoThoiViecs.Where(s => s.maLyDoThoiViec == id).FirstOrDefault();
                    lyDoThoiViec.tenLyDoThoiViec = collection["tenLyDoThoiViec"];
                    lyDoThoiViec.ngayLap = DateTime.Now;
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
                    lyDoThoiViec = context.tbl_DM_LyDoThoiViecs.Where(s => s.maLyDoThoiViec == id).FirstOrDefault();
                    context.tbl_DM_LyDoThoiViecs.DeleteOnSubmit(lyDoThoiViec);
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
