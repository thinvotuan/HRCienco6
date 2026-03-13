using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Models.DanhMuc;

namespace BatDongSan.Controllers.DanhMuc
{
    public class ChiNhanhNganHangController : ApplicationController
    {
        private LinqDanhMucDataContext context;
        private IList<tbl_DM_ChiNhanhNganHang> chiNhanhs;
        private tbl_DM_ChiNhanhNganHang chiNhanh;
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
                chiNhanhs = context.tbl_DM_ChiNhanhNganHangs.ToList();
            }
            return View(chiNhanhs);
        }

        //
        // GET: /ChiNhanhNganHang/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /ChiNhanhNganHang/Create

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            chiNhanh = new tbl_DM_ChiNhanhNganHang();
            return PartialView("Create", chiNhanh);
        }

        //
        // POST: /ChiNhanhNganHang/Create

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
                    chiNhanh = new tbl_DM_ChiNhanhNganHang();
                    chiNhanh.diaChi = collection["diaChi"];
                    chiNhanh.dienThoai = collection["dienThoai"];
                    chiNhanh.email = collection["email"];
                    chiNhanh.fax = collection["fax"];
                    chiNhanh.ghiChu = collection["ghiChu"];
                    chiNhanh.maChiNhanh = collection["maChiNhanh"];
                    chiNhanh.maSoThue = collection["maSoThue"];
                    chiNhanh.ngayLap = DateTime.Now;
                    chiNhanh.nguoiLap = GetUser().manv; ;
                    chiNhanh.tenChiNhanh = collection["tenChiNhanh"];
                    chiNhanh.website = collection["website"];
                    context.tbl_DM_ChiNhanhNganHangs.InsertOnSubmit(chiNhanh);
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
        // GET: /ChiNhanhNganHang/Edit/5

        public ActionResult Edit(string id)
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
                chiNhanh = context.tbl_DM_ChiNhanhNganHangs.Where(s => s.maChiNhanh == id).FirstOrDefault();
            }
            return PartialView("Edit", chiNhanh);
        }

        //
        // POST: /ChiNhanhNganHang/Edit/5

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
                using (context = new LinqDanhMucDataContext())
                {
                    chiNhanh = context.tbl_DM_ChiNhanhNganHangs.Where(s => s.maChiNhanh == id).FirstOrDefault();
                    chiNhanh.diaChi = collection["diaChi"];
                    chiNhanh.dienThoai = collection["dienThoai"];
                    chiNhanh.email = collection["email"];
                    chiNhanh.fax = collection["fax"];
                    chiNhanh.ghiChu = collection["ghiChu"];
                    chiNhanh.maChiNhanh = collection["maChiNhanh"];
                    chiNhanh.maSoThue = collection["maSoThue"];
                    chiNhanh.tenChiNhanh = collection["tenChiNhanh"];
                    chiNhanh.website = collection["website"];                    
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
        // POST: /ChiNhanhNganHang/Delete/5

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
                using (context = new LinqDanhMucDataContext())
                {
                    chiNhanh = context.tbl_DM_ChiNhanhNganHangs.Where(s => s.maChiNhanh == id).FirstOrDefault();
                    context.tbl_DM_ChiNhanhNganHangs.DeleteOnSubmit(chiNhanh);
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
