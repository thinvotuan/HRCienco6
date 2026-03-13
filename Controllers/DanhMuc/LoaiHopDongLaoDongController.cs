using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Helper.Utils;
using BatDongSan.Models.DanhMuc;

namespace BatDongSan.Controllers.DanhMuc
{
    public class LoaiHopDongLaoDongController : ApplicationController
    {
        private LinqDanhMucDataContext context;
        private IList<tbl_DM_ThoiHanHopDongLaoDong> loaiHopDongs;
        private tbl_DM_ThoiHanHopDongLaoDong loaiHopDong;
        private readonly string MCV = "LoaiHopDongLaoDong";
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
                loaiHopDongs = context.tbl_DM_ThoiHanHopDongLaoDongs.ToList();
            }
            return View(loaiHopDongs);
        }

        //
        // GET: /LoaiHopDongLaoDong/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /LoaiHopDongLaoDong/Create

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            return PartialView("Create", loaiHopDong);
        }

        //
        // POST: /LoaiHopDongLaoDong/Create

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
                    loaiHopDong = new tbl_DM_ThoiHanHopDongLaoDong();
                    loaiHopDong.ghiChu = collection["ghiChu"];
                    loaiHopDong.loaiThoiHan = collection["loaiThoiHan"] == "True" ? true : false;
                    loaiHopDong.ngayLap = DateTime.Now;
                    loaiHopDong.nguoiLap = GetUser().manv;
                    loaiHopDong.soThang = Convert.ToInt16(collection["soThang"]);
                    loaiHopDong.tenThoiHanHopDong = collection["tenThoiHanHopDong"];
                    context.tbl_DM_ThoiHanHopDongLaoDongs.InsertOnSubmit(loaiHopDong);
                    context.SubmitChanges();
                    SaveActiveHistory("Sửa loại hợp đồng lao động: " + loaiHopDong.tenThoiHanHopDong);
                }
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /LoaiHopDongLaoDong/Edit/5

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
                loaiHopDong = context.tbl_DM_ThoiHanHopDongLaoDongs.Where(s => s.id == id).FirstOrDefault(); ;
            }
            return PartialView("Edit", loaiHopDong);
        }

        //
        // POST: /LoaiHopDongLaoDong/Edit/5

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
                    loaiHopDong = context.tbl_DM_ThoiHanHopDongLaoDongs.Where(s => s.id == id).FirstOrDefault(); 
                    loaiHopDong.ghiChu = collection["ghiChu"];
                    loaiHopDong.loaiThoiHan = collection["loaiThoiHan"] == "True" ? true : false;
                    loaiHopDong.ngayLap = DateTime.Now;
                    loaiHopDong.nguoiLap = GetUser().manv;
                    loaiHopDong.soThang = Convert.ToInt16(collection["soThang"]);
                    loaiHopDong.tenThoiHanHopDong = collection["tenThoiHanHopDong"];                    
                    context.SubmitChanges();
                     SaveActiveHistory("Sửa loại hợp đồng lao động: " + id);
                }
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // POST: /LoaiHopDongLaoDong/Delete/5

        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
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
                    loaiHopDong = context.tbl_DM_ThoiHanHopDongLaoDongs.Where(s => s.id == id).FirstOrDefault(); ;
                    context.tbl_DM_ThoiHanHopDongLaoDongs.DeleteOnSubmit(loaiHopDong);
                    context.SubmitChanges();
                    SaveActiveHistory("Xóa hợp đồng lao động: " + id);
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
