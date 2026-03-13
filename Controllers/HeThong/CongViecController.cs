using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Models.HeThong;
using BatDongSan.Helper.Common;

namespace BatDongSan.Controllers.HeThong
{
    public class CongViecController : ApplicationController
    {
        private LinqHeThongDataContext context = new LinqHeThongDataContext();
        private IList<Sys_CongViec> congViecs;
        private Sys_CongViec congViec;
        private Sys_VuViecCuaCongViec vuViecCuaCongViec;
        private Sys_CongViecVaVuViec congViecVuViec;
        private bool? permission;
        private readonly string MCV = "DMCongViec";

        public ActionResult Index(string congViecCha, string congViecCon)
        {
            try
            {
                #region Role user
                permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion

                if (congViecCha == null) congViecCha = "HETHONG";

                congViecs = context.Sys_CongViecs.OrderBy(o => o.doUuTien).ToList();
                ViewBag.VuViecs = context.Sys_VuViecs.OrderBy(s => s.maVuViec).ToList();

                //Danh sách công việc cha
                IList<Sys_CongViec> congViecChas = congViecs.Where(s => s.maCha == null).ToList();
                ViewBag.CongViecChas = new SelectList(congViecChas, "maCongViec", "tenCongViec", congViecCha);
                //Danh sách công việc con
                IList<Sys_CongViec> congViecCons = congViecs.Where(s => s.maCha == congViecCha).ToList();
                congViecCons.Insert(0, new Sys_CongViec { maCongViec = "", tenCongViec = "[ Tất cả ]" });
                ViewBag.CongViecCons = new SelectList(congViecCons, "maCongViec", "tenCongViec", congViecCon);

                //Danh sách vụ việc của công việc con
                ViewBag.VuViecOfCV = context.sp_Sys_VuViecCuaCongViec(null, null).ToList();

                if (String.IsNullOrWhiteSpace(congViecCon))
                {
                    if (context.sp_Sys_CongViecVuViec().Where(s => s.maCha == congViecCha).ToList().Count < 1)
                        ViewBag.CVVV = context.sp_Sys_CongViecVuViec().Where(s => s.maCongViec == congViecCha).ToList();
                    else
                        ViewBag.CVVV = context.sp_Sys_CongViecVuViec().Where(s => s.maCha == congViecCha).ToList();
                }
                else
                {
                    if (context.sp_Sys_CongViecVuViec().Where(s => s.maCha == congViecCon).ToList().Count <= 1)
                    {
                        ViewBag.CVVV = context.sp_Sys_CongViecVuViec().Where(s => s.maCongViec == congViecCon).ToList();
                    }
                    else
                        ViewBag.CVVV = context.sp_Sys_CongViecVuViec().Where(s => s.maCha == congViecCon).ToList();
                }

                return View(congViecs);
            }
            catch
            {
                return RedirectToAction("Index");
            }
        }

        //
        // GET: /CongViec/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /CongViec/Create

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            congViecs = context.Sys_CongViecs.ToList();
            congViecs.Insert(0, new Sys_CongViec { maCongViec = "", tenCongViec = "" });
            ViewBag.CongViecs = new SelectList(congViecs, "maCongViec", "tenCongViec");
            ViewBag.VuViecs = context.Sys_VuViecs.ToList();
            return View();

        }

        //
        // POST: /CongViec/Create

        [HttpPost]
        public ActionResult Create(FormCollection collection, string[] maVuViecs)
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
                congViec = new Sys_CongViec();
                congViec.maCongViec = collection["maCongViec"];
                congViec.doUuTien = String.IsNullOrEmpty(collection["doUuTien"]) ? (int?)null : Convert.ToInt32(collection["doUuTien"].ToString());
                congViec.ghiChu = collection["ghiChu"];
                congViec.maCha = String.IsNullOrEmpty(collection["maCha"]) ? null : collection["maCha"];
                congViec.onMenu = collection["onmenu"].Contains("true") ? true : false;
                congViec.tenAction = collection["tenAction"];
                congViec.tenCongViec = collection["tenCongViec"];
                congViec.tenController = collection["tenController"];
                congViec.maIcon = String.IsNullOrEmpty(collection["maIcon"]) ? null : collection["maIcon"];
                if (context.Sys_CongViecs.Where(a => a.maCongViec == congViec.maCongViec).FirstOrDefault() != null)
                {
                    TempData["Mssg"] = "Mã công việc đã tồn tại";
                    congViecs = context.Sys_CongViecs.ToList();
                    congViecs.Insert(0, new Sys_CongViec { maCongViec = "", tenCongViec = "" });
                    ViewBag.CongViecs = new SelectList(congViecs, "maCongViec", "tenCongViec");
                    ViewBag.VuViecs = context.Sys_VuViecs.ToList();
                    return View(congViec);
                }
                context.Sys_CongViecs.InsertOnSubmit(congViec);
                if (maVuViecs != null)
                {
                    InsertVuViecCuaCongViec(maVuViecs);
                }
                context.SubmitChanges();
                SaveActiveHistory("Thêm mới công việc. MãCV: "+ congViec.maCongViec + " Tên CV: "+ congViec.tenCongViec +" Controller: " + congViec.tenController);
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /CongViec/Edit/5

        public ActionResult Edit(string id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            congViec = context.Sys_CongViecs.Where(s => s.maCongViec == id).FirstOrDefault();
            congViecs = context.Sys_CongViecs.ToList();
            congViecs.Insert(0, new Sys_CongViec { maCongViec = "", tenCongViec = "" });
            ViewBag.CongViecs = new SelectList(congViecs, "maCongViec", "tenCongViec", congViec.maCha);
            ViewBag.VuViecs = context.Sys_VuViecs.ToList();
            ViewBag.VuViecCuaCVs = context.sp_Sys_VuViecCuaCongViec(null, null).Where(s => s.maCongViec == id).Select(s => s.maVuViec).ToList();
            return View(congViec);
        }

        //
        // POST: /CongViec/Edit/5

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
                congViec = context.Sys_CongViecs.Where(s => s.maCongViec == id).FirstOrDefault();
                congViec.doUuTien = String.IsNullOrEmpty(collection["doUuTien"]) ? (int?)null : Convert.ToInt32(collection["douutien"].ToString());
                congViec.ghiChu = collection["ghiChu"];
                congViec.maCha = String.IsNullOrWhiteSpace(collection["maCha"]) ? null : collection["maCha"];
                congViec.onMenu = collection["onMenu"].Contains("true") ? true : false;
                congViec.tenAction = collection["tenAction"];
                congViec.tenCongViec = collection["tenCongViec"];
                congViec.tenController = collection["tenController"];
                congViec.maIcon = String.IsNullOrWhiteSpace(collection["maIcon"]) ? null : collection["maIcon"];
                context.SubmitChanges();
                SaveActiveHistory("Sửa công việc. MãCV: " + congViec.maCongViec + " Tên CV: " + congViec.tenCongViec + " Controller: " + congViec.tenController);
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

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
                DeleteFromCongViecVV(id);
                DeleteFromVuViecCV(id);
                congViec = context.Sys_CongViecs.Where(s => s.maCongViec == id).FirstOrDefault();
                context.Sys_CongViecs.DeleteOnSubmit(congViec);
                context.SubmitChanges();
                SaveActiveHistory("Xóa công việc. MãCV: " + congViec.maCongViec + " Tên CV: " + congViec.tenCongViec + " Controller: " + congViec.tenController);
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }


        [HttpPost]
        public ActionResult UpdateVuViec(string maVuViec, string maCongViec)
        {
            try
            {
                vuViecCuaCongViec = context.Sys_VuViecCuaCongViecs.Where(s => s.maVuViec == maVuViec && s.maCongViec == maCongViec).FirstOrDefault();
                if (vuViecCuaCongViec != null)
                {
                    context.Sys_VuViecCuaCongViecs.DeleteOnSubmit(vuViecCuaCongViec);
                }
                else
                {
                    vuViecCuaCongViec = new Sys_VuViecCuaCongViec();
                    vuViecCuaCongViec.maVuViec = maVuViec;
                    vuViecCuaCongViec.maCongViec = maCongViec;
                    context.Sys_VuViecCuaCongViecs.InsertOnSubmit(vuViecCuaCongViec);
                }

                congViecVuViec = context.Sys_CongViecVaVuViecs.Where(s => s.maCongViec == maCongViec && s.maVuViec == maVuViec).FirstOrDefault();
                if (congViecVuViec != null)
                {
                    context.Sys_CongViecVaVuViecs.DeleteOnSubmit(congViecVuViec);
                }
                context.SubmitChanges();
                return Json(new { success = true });
            }
            catch
            {
                return View();
            }
        }
        public ActionResult ListIcons()
        {
            return PartialView("PartialListIcon");
        }
        private void InsertVuViecCuaCongViec(string[] maVuViecs)
        {
            try
            {
                List<Sys_VuViecCuaCongViec> vuViecCuaCongViecs = new List<Sys_VuViecCuaCongViec>();
                string maCongViec = congViec.maCongViec;
                if (congViec.maCongViec != null && maVuViecs != null)
                {
                    for (int i = 0; i < maVuViecs.Length; i++)
                    {
                        Sys_VuViecCuaCongViec vuViecCuaCongViec = new Sys_VuViecCuaCongViec();
                        vuViecCuaCongViec.maCongViec = congViec.maCongViec;
                        vuViecCuaCongViec.maVuViec = maVuViecs[i];
                        vuViecCuaCongViecs.Add(vuViecCuaCongViec);
                    }
                    context.Sys_VuViecCuaCongViecs.InsertAllOnSubmit(vuViecCuaCongViecs);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void DeleteFromCongViecVV(string id)
        {
            IList<Sys_CongViecVaVuViec> CVVVs = new List<Sys_CongViecVaVuViec>();
            CVVVs = context.Sys_CongViecVaVuViecs.Where(s => s.maCongViec == id).ToList();
            context.Sys_CongViecVaVuViecs.DeleteAllOnSubmit(CVVVs);
        }

        private void DeleteFromVuViecCV(string id)
        {
            IList<Sys_VuViecCuaCongViec> VVOfCVs = new List<Sys_VuViecCuaCongViec>();
            VVOfCVs = context.Sys_VuViecCuaCongViecs.Where(s => s.maCongViec == id).ToList();
            context.Sys_VuViecCuaCongViecs.DeleteAllOnSubmit(VVOfCVs);
        }
    }
}
