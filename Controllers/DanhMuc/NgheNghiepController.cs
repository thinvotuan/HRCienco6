using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Models.DanhMuc;

namespace BatDongSan.Controllers.DanhMuc
{
    public class NgheNghiepController : ApplicationController
    {
        private LinqDanhMucDataContext context;
        private IList<tbl_DM_NgheNghiep> ngheNghieps;
        private tbl_DM_NgheNghiep ngheNghiep;        
        private readonly string MCV = "NgheNghiep";
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
                ngheNghieps = context.tbl_DM_NgheNghieps.ToList();
            }
            return View(ngheNghieps);
        }

        //
        // GET: /NgheNghiep/Create

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            ngheNghiep = new tbl_DM_NgheNghiep();
            return PartialView("Create", ngheNghiep);
        }

        //
        // POST: /NgheNghiep/Create

        [HttpPost]
        public ActionResult Create(tbl_DM_NgheNghiep viewModel)
        {
            try
            {
                using (context = new LinqDanhMucDataContext())
                {
                    ngheNghiep = viewModel;

                    var check = context.tbl_DM_NgheNghieps.Where(s => s.maNgheNghiep == viewModel.maNgheNghiep).FirstOrDefault();
                    if (check != null)
                    {
                        TempData["TonTai"] = "Mã này đã tồn tại. Vui lòng nhập mã khác";
                        return RedirectToAction("Index");
                    }

                    context.tbl_DM_NgheNghieps.InsertOnSubmit(ngheNghiep);
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
        // GET: /NgheNghiep/Edit/5

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
                ngheNghiep = context.tbl_DM_NgheNghieps.Where(s => s.maNgheNghiep == id).FirstOrDefault();
            }
            return PartialView("Edit", ngheNghiep);
        }

        //
        // POST: /NgheNghiep/Edit/5

        [HttpPost]
        public ActionResult Edit(string id, tbl_DM_NgheNghiep viewModel)
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
                    ngheNghiep = context.tbl_DM_NgheNghieps.Where(s => s.maNgheNghiep == id).FirstOrDefault();
                    ngheNghiep.tenNgheNghiep = viewModel.tenNgheNghiep;
                    ngheNghiep.ghiChu = viewModel.ghiChu;
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
        // POST: /NgheNghiep/Delete/5

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
                    ngheNghiep = context.tbl_DM_NgheNghieps.Where(s => s.maNgheNghiep == id).FirstOrDefault();
                    context.tbl_DM_NgheNghieps.DeleteOnSubmit(ngheNghiep);
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
