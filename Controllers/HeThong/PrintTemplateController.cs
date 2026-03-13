using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Models.HeThong;

namespace BatDongSan.Controllers.HeThong
{
    public class PrintTemplateController : ApplicationController
    {
        private LinqHeThongDataContext context = new LinqHeThongDataContext();
        private IList<Sys_PrintTemplate> templates;
        private Sys_PrintTemplate template;
        private readonly string MCV = "PrintTemplate";
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
            templates = context.Sys_PrintTemplates.ToList();
            return View(templates);
        }

        //
        // GET: /PrintTemplate/Details/5

        public ActionResult Details(int id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            return View();
        }

        //
        // GET: /PrintTemplate/Create

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            template = new Sys_PrintTemplate();
            ViewBag.Bases = context.Sys_BaseTemplates;
            var loaiMauIns = context.Sys_LoaiMauIns.ToList();
            loaiMauIns.Insert(0, new Sys_LoaiMauIn { maLoaiMauIn = "", tenLoaiMauIn = "[ Chọn loại mẫu in ]" });
            ViewBag.LoaiMauIns = new SelectList(loaiMauIns, "maLoaiMauIn", "tenLoaiMauIn");
            return View(template);
        }

        //
        // POST: /PrintTemplate/Create

        [HttpPost]
        [ValidateInput(false)]
        public ActionResult Create(Sys_PrintTemplate viewModel)
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

                // TODO: Add insert logic here
                template = viewModel;
                template.nguoiLap = GetUser().manv;
                template.ngayLap = DateTime.Now;
                template.loaiMauIn = "PT";
                var check = context.Sys_PrintTemplates.Where(s => s.maMauIn == template.maMauIn).FirstOrDefault();
                if (check != null)
                {
                    TempData["Exist"] = "Mã mẫu in đã tồn tại";
                    ViewBag.Bases = context.Sys_BaseTemplates;
                    return View(template);
                }
                context.Sys_PrintTemplates.InsertOnSubmit(template);
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /PrintTemplate/Edit/5

        public ActionResult Edit(int id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            ViewBag.Bases = context.Sys_BaseTemplates;
            template = context.Sys_PrintTemplates.Where(s => s.id == id).FirstOrDefault();
            var loaiMauIns = context.Sys_LoaiMauIns;
            ViewBag.LoaiMauIns = new SelectList(loaiMauIns, "maLoaiMauIn", "tenLoaiMauIn", template.loaiMauIn);
            return View(template);
        }

        //
        // POST: /PrintTemplate/Edit/5

        [HttpPost]
        [ValidateInput(false)]
        public ActionResult Edit(int id, Sys_PrintTemplate viewModel)
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

                template = context.Sys_PrintTemplates.Where(s => s.id == id).FirstOrDefault();
                template.html = viewModel.html;
                template.loaiMauIn = viewModel.loaiMauIn;
                template.tenMauIn = viewModel.tenMauIn;
                template.noiDung = viewModel.noiDung;
                context.SubmitChanges();
                return RedirectToAction("Edit", new { id = id });
            }
            catch
            {
                return View();
            }
        }

        //
        // POST: /PrintTemplate/Delete/5

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
                // TODO: Add delete logic here
                var deleteTemplate = context.Sys_PrintTemplates.Where(d => d.id == id);
                context.Sys_PrintTemplates.DeleteAllOnSubmit(deleteTemplate);
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
