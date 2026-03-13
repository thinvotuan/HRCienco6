using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Models.HeThong;
using BatDongSan.Models.NhanSu;

namespace BatDongSan.Controllers.NhanSu
{
    public class QuanLyGiupDoController : ApplicationController
    {
        private LinqHeThongDataContext contentHT = new LinqHeThongDataContext();

        private readonly string MCV = "QuanLyGiupDo";
        private bool? permission;
        //
        // GET: /GiupDo/

        public ActionResult Index(int? id, bool? msg)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            try
            {

                ViewBag.lsGiupDo = contentHT.Sys_GiupDos.ToList();
              
                return View();
            }
            catch (Exception ex)
            {

                ViewBag.Message = ex.Message;
                return View("Error");
            }
        }
        public ActionResult IndexView(int? id, bool? msg)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            try
            {
                ViewBag.lsGiupDo = contentHT.Sys_GiupDos.ToList();
                
                return View();
            }
            catch (Exception ex)
            {

                ViewBag.Message = ex.Message;
                return View("Error");
            }
        }

        //
        // GET: /GiupDo/Details/5

        
        public ActionResult Details(int id, bool? msg)
        {

            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXemChiTiet);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            try
            {
                if (msg == true)
                {
                    ViewBag.Message = "Cập nhật dữ liệu thành công";
                }
                else
                {
                    ViewBag.Message = string.Empty;
                }
                Sys_GiupDo giupDo = new Sys_GiupDo();

                giupDo = contentHT.Sys_GiupDos.Where(d => d.id.Equals(id)).FirstOrDefault();
                if (giupDo != null)
                {
                    ViewBag.data = giupDo;

                    return View();
                }
                else
                {
                    ViewData["Message"] = "Error";
                    return View("Error");
                }
            }
            catch (Exception ex)
            {

                ViewBag.Message = ex.Message;
                return View("Error");
            }
        }
        public ActionResult DetailsView(int id, bool? msg)
        {

            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXemChiTiet);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            try
            {
                if (msg == true)
                {
                    ViewBag.Message = "Cập nhật dữ liệu thành công";
                }
                else
                {
                    ViewBag.Message = string.Empty;
                }
                Sys_GiupDo giupDo = new Sys_GiupDo();

                giupDo = contentHT.Sys_GiupDos.Where(d => d.id.Equals(id)).FirstOrDefault();
                if (giupDo != null)
                {
                    ViewBag.data = giupDo;

                    return View();
                }
                else
                {
                    ViewData["Message"] = "Error";
                    return View("Error");
                }
            }
            catch (Exception ex)
            {

                ViewBag.Message = ex.Message;
                return View("Error");
            }
        }
        //
        // GET: /GiupDo/Create

        public ActionResult Create()
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

                Sys_GiupDo data = new Sys_GiupDo();
               
                return View("");
            }
            catch (Exception ex)
            {
               
                ViewBag.Message = ex.Message;
                return View("Error");
            }
        }

        //
        // POST: /GiupDo/Create

        [AcceptVerbs(HttpVerbs.Post)][ValidateInput(false)]
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
                // TODO: Add insert logic here
                
                Sys_GiupDo giupDo = new Sys_GiupDo();
                giupDo.ten = collection.Get("ten");
                giupDo.noiDung = collection.Get("noiDung");
                giupDo.ngayLap = DateTime.Now;
                giupDo.nguoiLap = GetUser().manv;
              
                contentHT.Sys_GiupDos.InsertOnSubmit(giupDo);
                contentHT.SubmitChanges();
                int id = contentHT.Sys_GiupDos.OrderByDescending(d => d.id).FirstOrDefault().id;
                return RedirectToAction("Index/", new { msg = false });
            }



            catch (Exception ex)
            {

                ViewBag.Message = ex.Message;
                return View("Error");
            }
        }

       

        //
        // GET: /GiupDo/Edit/5

        public ActionResult Edit(int id, bool? msg)
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
                if (msg == true)
                {
                    ViewBag.Message = "Cập nhật dữ liệu thành công";
                }
                else
                {
                    ViewBag.Message = string.Empty;
                }
                Sys_GiupDo giupDo = new Sys_GiupDo();
               
                giupDo = contentHT.Sys_GiupDos.Where(d => d.id.Equals(id)).FirstOrDefault();
                if (giupDo != null)
                {
                    ViewBag.data = giupDo;

                    return View();
                }
                else{
                    ViewData["Message"] = "Error";
                    return View("Error");
                }
            }
            catch (Exception ex)
            {
                
                ViewBag.Message = ex.Message;
                return View("Error");
            }
        }

        //
        // POST: /GiupDo/Edit/5

        [AcceptVerbs(HttpVerbs.Post)]
        [ValidateInput(false)]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here
               
                Sys_GiupDo giupDo = new Sys_GiupDo();
                giupDo = contentHT.Sys_GiupDos.Where(d => d.id.Equals(id)).FirstOrDefault();
                giupDo.ten = collection.Get("ten");
                giupDo.noiDung = collection.Get("noiDung");
               
                contentHT.SubmitChanges();
                return RedirectToAction("Edit/" + id, new { msg = true });
            }
            catch (Exception ex)
            {
               
                ViewData["Message"] = ex.Message;
                return View("Error");
            }
        }

      

        public IQueryable<Sys_GiupDo> FindAll()
        {
          
            var sql = from p in contentHT.Sys_GiupDos
                      select p;
            return sql;
        }


        //
        // GET: /GiupDo/Delete/5

        //public ActionResult Delete(int id)
        //{
        //    return View();
        //}

        //
        // POST: /GiupDo/Delete/5

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Delete(int id)
        {
            try
            {
                // TODO: Add delete logic here
              
                Sys_GiupDo giupDo = new Sys_GiupDo();
                giupDo = contentHT.Sys_GiupDos.Where(d => d.id.Equals(id)).FirstOrDefault();
                contentHT.Sys_GiupDos.DeleteOnSubmit(giupDo);
                contentHT.SubmitChanges();
                
              
            }
            catch (Exception ex)
            {
               
                return Json(ex.Message);
            }
            return Json(String.Empty);
        }
        public ActionResult GetThongTin(int id)
        {
            
            var thongTin = contentHT.Sys_GiupDos.Where(s => s.id == id).FirstOrDefault();
          
            return PartialView("_PartialView", thongTin);
        }
    }
}
