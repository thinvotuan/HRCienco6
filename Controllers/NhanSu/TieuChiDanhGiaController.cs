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
    public class TieuChiDanhGiaController : ApplicationController
    {
        private LinqNhanSuDataContext hr = new LinqNhanSuDataContext();

        private readonly string MCV = "TieuChiDanhGia";
        private bool? permission;
        //
        // GET: /TieuChiDanhGia/

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
               
                ViewBag.lsTieuChi = hr.tbl_NS_TieuChis.ToList();
                return View();
            }
            catch (Exception ex)
            {

                ViewBag.Message = ex.Message;
                return View("Error");
            }
        }

        //
        // GET: /TieuChiDanhGia/Details/5

        
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
                tbl_NS_TieuChi tieuChi = new tbl_NS_TieuChi();

                tieuChi = hr.tbl_NS_TieuChis.Where(d => d.id.Equals(id)).FirstOrDefault();
                if (tieuChi != null)
                {
                    ViewBag.data = tieuChi;

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
        // GET: /TieuChiDanhGia/Create

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

                tbl_NS_TieuChi data = new tbl_NS_TieuChi();
               
                return View("");
            }
            catch (Exception ex)
            {
               
                ViewBag.Message = ex.Message;
                return View("Error");
            }
        }

        //
        // POST: /TieuChiDanhGia/Create

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Create(FormCollection collection)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            //#region Role user
            //permisson = GetPermission(maCongViec, new string[] { BangPhanQuyen.QuyenThem });
            //if (!permisson.HasValue)
            //    return RedirectToAction("LogOn", "Account");
            //if (!permisson.Value)
            //    return View("AccessDenied");
            //#endregion
            try
            {
                // TODO: Add insert logic here
                
                tbl_NS_TieuChi tieuChi = new tbl_NS_TieuChi();
                tieuChi.tenTieuChi = collection.Get("tenNhomTieuChi");
                tieuChi.maNhom = collection.Get("maNhomTieuChi");
                tieuChi.heSoDiem =Convert.ToDouble(collection.Get("heSoDiem"));
                tieuChi.ghiChu = collection.Get("note");
                //kiểm tra tên nhóm
                if (!TestTenNhom(tieuChi.maNhom.ToString()))
                {
                    ViewBag.Message = "Mã nhóm tiêu chí : " + tieuChi.tenTieuChi + " đã tồn tại rồi. Vui lòng nhập lại.";
                    ViewBag.data = tieuChi;
                    return View();
                }
                hr.tbl_NS_TieuChis.InsertOnSubmit(tieuChi);
                hr.SubmitChanges();
                int id = hr.tbl_NS_TieuChis.OrderByDescending(d => d.id).FirstOrDefault().id;
                return RedirectToAction("Edit/" + id, new { msg = false });
            }
            catch (Exception ex)
            {

                ViewBag.Message = ex.Message;
                return View("Error");
            }
        }

        public bool TestTenNhom(string maNhom)
        {
            
            var ds = hr.tbl_NS_TieuChis.ToList();
            foreach (var item in ds)
            {
                if (item.maNhom.Equals(maNhom))
                {
                    return false;
                }
            }
            return true;
        }

        //
        // GET: /TieuChiDanhGia/Edit/5

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
                tbl_NS_TieuChi tieuChi = new tbl_NS_TieuChi();
               
                tieuChi = hr.tbl_NS_TieuChis.Where(d => d.id.Equals(id)).FirstOrDefault();
                if (tieuChi != null)
                {
                    ViewBag.data = tieuChi;

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
        // POST: /TieuChiDanhGia/Edit/5

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here
               
                tbl_NS_TieuChi tieuChi = new tbl_NS_TieuChi();
                tieuChi = hr.tbl_NS_TieuChis.Where(d => d.id.Equals(id)).FirstOrDefault();
                tieuChi.tenTieuChi = collection.Get("tenNhomTieuChi");
                tieuChi.heSoDiem = Convert.ToDouble(collection.Get("heSoDiem"));
                tieuChi.ghiChu = collection.Get("note");
                if (TestTenNhomEdit(tieuChi))
                {
                    ViewData["Message"] = "Tên nhóm tiêu chí : " + tieuChi.tenTieuChi + " đã tồn tại rồi. Vui lòng nhập lại.";
                    ViewData["data"] = tieuChi;
                    return View();
                }
                hr.SubmitChanges();
                return RedirectToAction("Edit/" + id, new { msg = true });
            }
            catch (Exception ex)
            {
               
                ViewData["Message"] = ex.Message;
                return View("Error");
            }
        }

        public bool TestTenNhomEdit(tbl_NS_TieuChi tieuChi)
        {
            var data = FindAll().Where(c => !c.tenTieuChi.Equals(tieuChi.tenTieuChi));
            if (data.Where(p => p.tenTieuChi.Equals(tieuChi.tenTieuChi)).Count() > 0)
                return true;
            return false;
        }

        public IQueryable<tbl_NS_TieuChi> FindAll()
        {
          
            var sql = from p in hr.tbl_NS_TieuChis
                      select p;
            return sql;
        }


        //
        // GET: /TieuChiDanhGia/Delete/5

        //public ActionResult Delete(int id)
        //{
        //    return View();
        //}

        //
        // POST: /TieuChiDanhGia/Delete/5

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Delete(int id)
        {
            try
            {
                // TODO: Add delete logic here
              
                tbl_NS_TieuChi tieuChi = new tbl_NS_TieuChi();
                tieuChi = hr.tbl_NS_TieuChis.Where(d => d.id.Equals(id)).FirstOrDefault();
                hr.tbl_NS_TieuChis.DeleteOnSubmit(tieuChi);
                hr.SubmitChanges();
                //return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
               
                return Json(ex.Message);
            }
            return Json(String.Empty);
        }
    }
}
