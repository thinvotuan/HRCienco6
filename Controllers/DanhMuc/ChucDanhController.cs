using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Models.DanhMuc;

namespace BatDongSan.Controllers.DanhMuc
{
    public class ChucDanhController : ApplicationController
    {
        private LinqDanhMucDataContext context;
        private IList<Sys_ChucDanh> chucDanhs;
        private Sys_ChucDanh chucDanh;
        private readonly string MCV = "ChucDanh";
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

            context = new LinqDanhMucDataContext();
            chucDanhs = context.Sys_ChucDanhs.OrderBy(o=>o.SoCapBac).ToList();
            return View(chucDanhs);
        }

        //
        // GET: /ChucDanh/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /ChucDanh/Create

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
            ViewBag.CapBacs = context.tbl_DM_CapBacChucDanhs.ToList();
            return PartialView("Create");
        }

        //
        // POST: /ChucDanh/Create

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
                context = new LinqDanhMucDataContext();
                chucDanh = new Sys_ChucDanh();
                chucDanh.GhiChu = collection["GhiChu"];
                chucDanh.MaChucDanh = collection["MaChucDanh"];
                chucDanh.NgayLap = DateTime.Now;
                chucDanh.NhiemVu = collection["NhiemVu"];
                chucDanh.SoCapBac = Convert.ToInt32(collection["SoCapBac"]);
                chucDanh.TenChucDanh = collection["TenChucDanh"];
                var obj = context.Sys_ChucDanhs.Where(s => s.MaChucDanh == chucDanh.MaChucDanh).FirstOrDefault();
                if (obj != null)
                {
                    TempData["TonTai"] = "Lưu không thành công - mã chức danh đã tồn tại";
                    return RedirectToAction("Index");
                }
                context.Sys_ChucDanhs.InsertOnSubmit(chucDanh);
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /ChucDanh/Edit/5

        public ActionResult Edit(string id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            context = new LinqDanhMucDataContext();
            chucDanh = context.Sys_ChucDanhs.Where(s => s.MaChucDanh == id).FirstOrDefault();
            ViewBag.CapBacs = context.tbl_DM_CapBacChucDanhs.ToList();
            return PartialView("Edit", chucDanh);
        }

        //
        // POST: /ChucDanh/Edit/5

        [HttpPost]
        public ActionResult Edit(string id, FormCollection collection)
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
                context = new LinqDanhMucDataContext();
                chucDanh = context.Sys_ChucDanhs.Where(s => s.MaChucDanh == id).FirstOrDefault();
                chucDanh.GhiChu = collection["GhiChu"];                
                chucDanh.NgayLap = DateTime.Now;
                chucDanh.NhiemVu = collection["NhiemVu"];
                chucDanh.SoCapBac = Convert.ToInt32(collection["SoCapBac"]);
                chucDanh.TenChucDanh = collection["TenChucDanh"];                
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

 
        //
        // POST: /ChucDanh/Delete/5

        [HttpPost]
        public ActionResult Delete(string id)
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
                context = new LinqDanhMucDataContext();
                chucDanh = context.Sys_ChucDanhs.Where(s => s.MaChucDanh == id).FirstOrDefault();
                context.Sys_ChucDanhs.DeleteOnSubmit(chucDanh);
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
    }
}
