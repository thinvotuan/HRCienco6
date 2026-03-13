using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Models.DanhMuc;

namespace BatDongSan.Controllers.DanhMuc
{
    public class LoaiNghiPhepController : ApplicationController
    {
        private LinqDanhMucDataContext context;
        private IList<tbl_DM_LoaiNghiPhep> loaiNgayNghis;
        private tbl_DM_LoaiNghiPhep loaiNgayNghi;
        private readonly string MCV = "LoaiNghiPhep";
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
                loaiNgayNghis = context.tbl_DM_LoaiNghiPheps.ToList();
            }
            return View(loaiNgayNghis);
        }

        //
        // GET: /LoaiNghiPhep/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /LoaiNghiPhep/Create

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            loaiNgayNghi = new tbl_DM_LoaiNghiPhep();
            loaiNgayNghi.soNgayNghiToiDa = 0;
            loaiNgayNghi.phanTramLuongHuong = 0;
            return PartialView("Create", loaiNgayNghi);
        }

        //
        // POST: /LoaiNghiPhep/Create

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
                    loaiNgayNghi = new tbl_DM_LoaiNghiPhep();
                    loaiNgayNghi.maLoaiNghiPhep = collection["maLoaiNghiPhep"];
                    loaiNgayNghi.tenLoaiNghiPhep = collection["tenLoaiNghiPhep"];
                    loaiNgayNghi.soNgayNghiToiDa = Convert.ToDecimal(collection["soNgayNghiToiDa"]);
                    loaiNgayNghi.ghiChu = collection["ghiChu"];
                    loaiNgayNghi.phanTramLuongHuong = Convert.ToDecimal(collection["phanTramLuongHuong"]);
                    loaiNgayNghi.tinhCong = collection["tinhCong"] == "True" ? true : false;
                    loaiNgayNghi.trangThai = collection["trangThai"].Contains("true") ? true : false;
                    loaiNgayNghi.nguoiLap = GetUser().manv;
                    loaiNgayNghi.ngayLap = DateTime.Now;
                    loaiNgayNghi.soNgayRangBuoc = String.IsNullOrEmpty(collection["soNgayRangBuoc"]) ? (int?)null : Convert.ToInt32(collection["soNgayRangBuoc"]);

                    if (context.tbl_DM_LoaiNghiPheps.Where(s => s.maLoaiNghiPhep == loaiNgayNghi.maLoaiNghiPhep).FirstOrDefault() != null)
                    {
                        TempData["TonTai"] = "Lưu không thành công - mã loại nghỉ phép đã tồn tại";
                        return RedirectToAction("Index");
                    }
                    context.tbl_DM_LoaiNghiPheps.InsertOnSubmit(loaiNgayNghi);
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
        // GET: /LoaiNghiPhep/Edit/5

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
                loaiNgayNghi = context.tbl_DM_LoaiNghiPheps.Where(s => s.id == id).FirstOrDefault();
            }
            return PartialView("Edit", loaiNgayNghi);
        }

        //
        // POST: /LoaiNghiPhep/Edit/5

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
                    loaiNgayNghi = context.tbl_DM_LoaiNghiPheps.Where(s => s.id == id).FirstOrDefault();
                    loaiNgayNghi.maLoaiNghiPhep = collection["maLoaiNghiPhep"];
                    loaiNgayNghi.tenLoaiNghiPhep = collection["tenLoaiNghiPhep"];
                    loaiNgayNghi.soNgayNghiToiDa = Convert.ToDecimal(collection["soNgayNghiToiDa"]);
                    loaiNgayNghi.ghiChu = collection["ghiChu"];
                    loaiNgayNghi.tinhCong = collection["tinhCong"] == "True" ? true : false;
                    loaiNgayNghi.trangThai = collection["trangThai"].Contains("true") ? true : false;
                    loaiNgayNghi.phanTramLuongHuong = Convert.ToDecimal(collection["phanTramLuongHuong"]);
                    loaiNgayNghi.ngayCapNhat = DateTime.Now;
                    loaiNgayNghi.soNgayRangBuoc = String.IsNullOrEmpty(collection["soNgayRangBuoc"]) ? (int?)null : Convert.ToInt32(collection["soNgayRangBuoc"]);
                    var obj = context.tbl_DM_LoaiNghiPheps.Where(s => s.maLoaiNghiPhep == loaiNgayNghi.maLoaiNghiPhep).FirstOrDefault();
                    if (obj != null && obj.id != loaiNgayNghi.id)
                    {
                        TempData["TonTai"] = "Mã đã tồn tại";
                        return PartialView("Edit", loaiNgayNghi);
                    }
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
        // POST: /LoaiNghiPhep/Delete/5

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
                    loaiNgayNghi = context.tbl_DM_LoaiNghiPheps.Where(s => s.id == id).FirstOrDefault();
                    context.tbl_DM_LoaiNghiPheps.DeleteOnSubmit(loaiNgayNghi);
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
