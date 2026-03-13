using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Models.DanhMuc;

namespace BatDongSan.Controllers.DanhMuc
{
    public class DMXepLoaiDGController : ApplicationController
    {
        private LinqDanhMucDataContext context;
        private IList<tbl_DM_XepLoaiDG> xepLoais;
        private tbl_DM_XepLoaiDG xepLoai;        
        private readonly string MCV = "DMXepLoaiDG";
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
                xepLoais = context.tbl_DM_XepLoaiDGs.OrderBy(d=>d.thuTu).ToList();
            }
            return View(xepLoais);
        }

        //
        // GET: /DMXepLoaiDG/Create

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            xepLoai = new tbl_DM_XepLoaiDG();
            return PartialView("Create", xepLoai);
        }

        //
        // POST: /DMXepLoaiDG/Create

        [HttpPost]
        public ActionResult Create(tbl_DM_XepLoaiDG viewModel)
        {
            try
            {
                using (context = new LinqDanhMucDataContext())
                {
                    xepLoai = viewModel;

                    var check = context.tbl_DM_XepLoaiDGs.Where(s => s.maXepLoai == viewModel.maXepLoai).FirstOrDefault();
                    if (check != null)
                    {
                        TempData["TonTai"] = "Mã này đã tồn tại. Vui lòng nhập mã khác";
                        return RedirectToAction("Index");
                    }

                    context.tbl_DM_XepLoaiDGs.InsertOnSubmit(xepLoai);
                    context.SubmitChanges();
                    return RedirectToAction("Index");
                }
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /DMXepLoaiDG/Edit/5

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
                xepLoai = context.tbl_DM_XepLoaiDGs.Where(s => s.maXepLoai == id).FirstOrDefault();
            }
            return PartialView("Edit", xepLoai);
        }

        //
        // POST: /DMXepLoaiDG/Edit/5

        [HttpPost]
        public ActionResult Edit(string id, tbl_DM_XepLoaiDG viewModel)
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
                    xepLoai = context.tbl_DM_XepLoaiDGs.Where(s => s.maXepLoai == id).FirstOrDefault();
                    xepLoai.tenXepLoai = viewModel.tenXepLoai;
                    xepLoai.ghiChu = viewModel.ghiChu;
                    xepLoai.thuTu = viewModel.thuTu ?? 0;
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
        // POST: /DMXepLoaiDG/Delete/5

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
                    xepLoai = context.tbl_DM_XepLoaiDGs.Where(s => s.maXepLoai == id).FirstOrDefault();
                    context.tbl_DM_XepLoaiDGs.DeleteOnSubmit(xepLoai);
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
