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
    public class MucLuongToiThieuController : ApplicationController
    {
        private LinqNhanSuDataContext context;
        private IList<tbl_NS_MucLuongToiThieu> mucLuongs;
        private tbl_NS_MucLuongToiThieu mucLuong;
        private readonly string MCV = "MucLuongToiThieu";
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
            mucLuongs = context.tbl_NS_MucLuongToiThieus.ToList();
            return View(mucLuongs);
        }

        //
        // GET: /MucLuongToiThieu/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /MucLuongToiThieu/Create

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            mucLuong = new tbl_NS_MucLuongToiThieu();
            mucLuong.ngayApDung = DateTime.Now;
            return PartialView("Create", mucLuong);
        }

        //
        // POST: /MucLuongToiThieu/Create

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
                mucLuong = new tbl_NS_MucLuongToiThieu();
                mucLuong.ghiChu = collection["ghiChu"];
                mucLuong.mucLuong = Convert.ToDecimal(collection["mucLuong"]);
                mucLuong.ngayCapNhat = DateTime.Now;
                mucLuong.ngayApDung = DateTime.ParseExact(collection["ngayApDung"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                context.tbl_NS_MucLuongToiThieus.InsertOnSubmit(mucLuong);
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /MucLuongToiThieu/Edit/5

        public ActionResult Edit(int id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            context = new LinqNhanSuDataContext();
            mucLuong = context.tbl_NS_MucLuongToiThieus.Where(s => s.id == id).FirstOrDefault();
            return PartialView("Edit", mucLuong);
        }

        //
        // POST: /MucLuongToiThieu/Edit/5

        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
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
                mucLuong = context.tbl_NS_MucLuongToiThieus.Where(s => s.id == id).FirstOrDefault();
                mucLuong.ghiChu = collection["ghiChu"];
                mucLuong.ngayCapNhat = DateTime.Now;
                mucLuong.mucLuong = Convert.ToDecimal(collection["mucLuong"]);                
                mucLuong.ngayApDung = DateTime.ParseExact(collection["ngayApDung"], "dd/MM/yyyy", CultureInfo.InvariantCulture);                
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }


        //
        // POST: /MucLuongToiThieu/Delete/5

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
                context = new LinqNhanSuDataContext();
                mucLuong = context.tbl_NS_MucLuongToiThieus.Where(s => s.id == id).FirstOrDefault();
                context.tbl_NS_MucLuongToiThieus.DeleteOnSubmit(mucLuong);
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
