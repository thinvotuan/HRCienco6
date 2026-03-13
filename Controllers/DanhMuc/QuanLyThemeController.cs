using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Models.HeThong;

namespace BatDongSan.Controllers.DanhMuc
{
    public class QuanLyThemeController : ApplicationController
    {

        private LinqHeThongDataContext context = new LinqHeThongDataContext();
        private IList<tbl_DM_TienIchTheme> chiNhanhs;
        private tbl_DM_TienIchTheme DMCapBac;
        private readonly string MCV = "QuanLyTheme";
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
            //Theme Color

            var listTheme = context.tbl_DM_TienIchThemes.OrderByDescending(d => d.id).ToList();
            Session["listTheme"] = listTheme;
             //Check idColor in table user
                var mauMacDinhUser = context.Sys_Users.Where(d => d.manv == GetUser().manv).Select(d => new { maMau = d.idColor}).ToArray();
                if (mauMacDinhUser != null) {
                    if (mauMacDinhUser.FirstOrDefault() != null)
                    {
                        Session["maMau"] = mauMacDinhUser.FirstOrDefault().maMau;
                    }
                }
                if (Session["maMau"] == null)
                {
                    var mauMacDinh = context.tbl_DM_TienIchThemes.Where(d => d.trangThai == 1).Select(d => new { maMau = d.maMau, trangThai = d.trangThai }).ToArray();
                    if (mauMacDinh != null)
                    {
                        if (mauMacDinh.FirstOrDefault() != null)
                        {
                            Session["maMau"] = mauMacDinh.FirstOrDefault().maMau;
                        }
                        else
                        {
                            Session["maMau"] = "#1A9DCC";
                        }
                    }
                }
            // End Theme
           chiNhanhs = context.tbl_DM_TienIchThemes.ToList();
           
            return View(chiNhanhs);
        }

        //
        // GET: /QuanLyTheme/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /QuanLyTheme/Create

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            DMCapBac = new tbl_DM_TienIchTheme();
            return PartialView("Create", DMCapBac);
        }

        //
        // POST: /QuanLyTheme/Create

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
                
                    var list = context.tbl_DM_TienIchThemes.OrderByDescending(d => d.id).ToList();
                     int id = 0;
                     if (list.Count >0)
                     {
                         id = list.FirstOrDefault().id;
                     }
                     else
                     {
                         id = 0;
                     }
                     id = id + 1;
                    DMCapBac = new tbl_DM_TienIchTheme();
                    DMCapBac.maMau =collection["maMau"];
                    DMCapBac.ten = collection["ten"];
                    
                    DMCapBac.id = id;
                    context.tbl_DM_TienIchThemes.InsertOnSubmit(DMCapBac);
                    context.SubmitChanges();
                
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /QuanLyTheme/Edit/5

        public ActionResult Edit(int? id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

           
                DMCapBac = context.tbl_DM_TienIchThemes.Where(s => s.id == id).FirstOrDefault();
           
            return PartialView("Edit", DMCapBac);
        }

        //
        // POST: /QuanLyTheme/Edit/5

        [HttpPost]
        public ActionResult Edit(int? id, FormCollection collection)
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
               
                    DMCapBac = context.tbl_DM_TienIchThemes.Where(s => s.id == id).FirstOrDefault();
                   
                    DMCapBac.maMau = collection["maMau"];
                    DMCapBac.ten = collection["ten"];
                   
                   
                   
                    context.SubmitChanges();
             
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // POST: /QuanLyTheme/Delete/5

        [HttpPost]
        public ActionResult Delete(int? id)
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
                
                    DMCapBac = context.tbl_DM_TienIchThemes.Where(s => s.id == id).FirstOrDefault();
                    context.tbl_DM_TienIchThemes.DeleteOnSubmit(DMCapBac);
                    context.SubmitChanges();
               
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
        public ActionResult ChangeTheme(int idTheme)
        {
            try
            {

                var thongTin = context.tbl_DM_TienIchThemes.Where(d => d.trangThai == 1).ToList();
                var some = context.tbl_DM_TienIchThemes.Where(d => d.trangThai == 1).ToList();
                some.ForEach(a =>
                {
                    a.trangThai = 0;

                }
                            );
                context.SubmitChanges();


                var thongTin2 = context.tbl_DM_TienIchThemes.Where(d => d.id == idTheme).FirstOrDefault();
                if (thongTin2 != null)
                {


                    thongTin2.trangThai = 1;
                    context.SubmitChanges();


                }
                var thongTinUser = context.Sys_Users.Where(d => d.manv == GetUser().manv).FirstOrDefault();
                if (thongTinUser != null)
                {


                    thongTinUser.idColor = thongTin2.maMau;
                    context.SubmitChanges();


                }
                return Json(new
                {
                    giaTri = idTheme

                });
            }
            catch (Exception e)
            {
                return Json("Lỗi: " + e.Message);
            }
        }
        public ActionResult changeThemeUser(string maMau, string tenMau, string url)
        {
            try
            {
                maMau = "#"+maMau;
                Session["maMau"] = maMau;
                Session["tenMau"] = tenMau;
            
                //Get thong tin user
             
                //var thongTin = context.tbl_DM_TienIchThemes.Where(d => d.trangThai ==1).ToList();
                //var some = context.tbl_DM_TienIchThemes.Where(d=>d.trangThai == 1).ToList();
                //some.ForEach(a =>
                //{
                //    a.trangThai = 0;
                    
                //}
                //            );
                //context.SubmitChanges();
                
               
                //var thongTin2 = context.tbl_DM_TienIchThemes.Where(d => d.maMau == maMau).FirstOrDefault();
                //if (thongTin2 != null)
                //{


                //    thongTin2.trangThai = 1;
                //    context.SubmitChanges();


                //}
                // update idColor for user.
                 var thongTinUser = context.Sys_Users.Where(d => d.manv == GetUser().manv ).FirstOrDefault();
                if (thongTinUser != null)
                {


                    thongTinUser.idColor = maMau;
                    context.SubmitChanges();


                }
                
            }
            catch
            {

               
            }
            Response.Redirect(url);
            return RedirectToAction("index");
        }
    }
}
