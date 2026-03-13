using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Models.DanhMuc;

namespace BatDongSan.Controllers.DanhMuc
{
    public class DMThuNhapKhacController : ApplicationController
    {
        private LinqDanhMucDataContext context = new LinqDanhMucDataContext();
        private IList<tbl_DM_ThuNhapKhac> _DMThuNhapKhacs;
        private tbl_DM_ThuNhapKhac _DMThuNhapKhac;
        private readonly string MCV = "DMThuNhapKhac";
        private bool? permission;
        //
        // GET: /DMThuNhapKhac/

        public ActionResult Index()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            _DMThuNhapKhacs = context.tbl_DM_ThuNhapKhacs.OrderByDescending(o => o.ngayLap).ToList();
            return View(_DMThuNhapKhacs);
        }

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            _DMThuNhapKhac = new tbl_DM_ThuNhapKhac();
            return PartialView("Create", _DMThuNhapKhac);
        }
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
                _DMThuNhapKhac = new tbl_DM_ThuNhapKhac();
                _DMThuNhapKhac.maLoaiThuNhapKhac = collection["maLoaiThuNhapKhac"];
                _DMThuNhapKhac.ngayLap = DateTime.Now;
                _DMThuNhapKhac.nguoiLap = GetUser().manv;
                _DMThuNhapKhac.tenLoaiThuNhapKhac = collection["tenLoaiThuNhapKhac"];
                _DMThuNhapKhac.coTinhBaoHiemKhong = collection["tinhBaoHiem"].Contains("true");
                _DMThuNhapKhac.coTinhThueKhong = collection["tinhThue"].Contains("true");
                _DMThuNhapKhac.ghiChu = collection["ghiChu"];

                var obj = context.tbl_DM_ThuNhapKhacs.Where(s => s.maLoaiThuNhapKhac == _DMThuNhapKhac.maLoaiThuNhapKhac).FirstOrDefault();
                if (obj != null)
                {
                    TempData["TonTai"] = "Lưu không thành công - mã loại thu nhập đã tồn tại.";
                    return RedirectToAction("Index");
                }
                context.tbl_DM_ThuNhapKhacs.InsertOnSubmit(_DMThuNhapKhac);
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
        public ActionResult Edit(string id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            _DMThuNhapKhac = context.tbl_DM_ThuNhapKhacs.Where(s => s.maLoaiThuNhapKhac == id).FirstOrDefault();
            return PartialView("Edit", _DMThuNhapKhac);
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
                _DMThuNhapKhac = context.tbl_DM_ThuNhapKhacs.Where(s => s.maLoaiThuNhapKhac == id).FirstOrDefault();
                _DMThuNhapKhac.tenLoaiThuNhapKhac = collection["tenLoaiThuNhapKhac"];
                _DMThuNhapKhac.coTinhBaoHiemKhong = collection["tinhBaoHiem"].Contains("true");
                _DMThuNhapKhac.coTinhThueKhong = collection["tinhThue"].Contains("true");
                _DMThuNhapKhac.ghiChu = collection["ghiChu"];
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
                _DMThuNhapKhac = context.tbl_DM_ThuNhapKhacs.Where(s => s.maLoaiThuNhapKhac == id).FirstOrDefault();
                context.tbl_DM_ThuNhapKhacs.DeleteOnSubmit(_DMThuNhapKhac);
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
