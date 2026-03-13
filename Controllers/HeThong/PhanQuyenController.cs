using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Models.HeThong;
using BatDongSan.Helper.Common;
using BatDongSan.Utils.Paging;

namespace BatDongSan.Controllers.HeThong
{
    public class PhanQuyenController : ApplicationController
    {
        private LinqHeThongDataContext context = new LinqHeThongDataContext();
        private IList<Sys_NhomUser> userGroups;
        private IList<sp_Sys_CongViecCuaUsersResult> quyenCuaNhomUsers;
        private IList<sp_Sys_User_IndexResult> userOfGroups;
        private IList<sp_Sys_VuViecCuaCongViecResult> vuViecCuaCongViecs;
        private Sys_NhomUser nhomUser;
        private Sys_CongViecVaVuViec congViecVaVuViec;
        private bool? permission;
        //private static int defaultPageSize = 20;
        private readonly string MCV = "PhanQuyenNhom";

        public ActionResult Index()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            userGroups = context.Sys_NhomUsers.ToList();
            if (Request.IsAjaxRequest())
            {
                return PartialView("PartialContent", userGroups);
            }
            return View(userGroups);
        }

        //
        // GET: /PhanQuyen/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /PhanQuyen/Create

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            return View();
        }

        //
        // POST: /PhanQuyen/Create

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
                nhomUser = new Sys_NhomUser();
                nhomUser.maNhomUser = collection["maNhomUser"];
                nhomUser.tenNhomUser = collection["tenNhomUser"];
                nhomUser.ghiChu = collection["ghiChu"];
                var nhomUsers = context.Sys_NhomUsers.Where(s => s.maNhomUser == nhomUser.maNhomUser || s.tenNhomUser == nhomUser.tenNhomUser).ToList();
                if (nhomUsers.Count() <= 0)
                {
                    context.Sys_NhomUsers.InsertOnSubmit(nhomUser);
                    context.SubmitChanges();
                    return RedirectToAction("Index");
                }
                else
                {
                    if (nhomUsers.Where(s => s.maNhomUser.ToLower() == nhomUser.maNhomUser.ToLower()).Count() > 0) TempData["MessgId"] = "Mã nhóm user đã tồn tại";
                    if (nhomUsers.Where(s => s.tenNhomUser.ToLower() == nhomUser.tenNhomUser.ToLower()).Count() > 0) TempData["MessgTen"] = "Tên nhóm user đã tồn tại";
                    return View(nhomUser);
                }
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /PhanQuyen/Edit/5

        public ActionResult Edit(string id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            ViewBag.NhanVienThuocNhoms = context.sp_Sys_User_Index(id, null, null, null, null, null).ToList();
            nhomUser = context.Sys_NhomUsers.Where(s => s.maNhomUser == id).FirstOrDefault();
            return View(nhomUser);
        }

        //
        // POST: /PhanQuyen/Edit/5

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
                var list = context.Sys_UserThuocNhoms.Where(s => s.maNhomUser == id);
                context.Sys_UserThuocNhoms.DeleteAllOnSubmit(list);

                nhomUser = context.Sys_NhomUsers.Where(s => s.maNhomUser == id).FirstOrDefault();
                nhomUser.tenNhomUser = collection["tenNhomUser"];
                nhomUser.ghiChu = collection["ghiChu"];
                //Cập nhật lại danh sách user thuộc nhóm
                IList<Sys_UserThuocNhom> userThuocNhoms = new List<Sys_UserThuocNhom>();
                string[] userIds = collection.GetValues("userId");
                if (userIds != null)
                {
                    for (int i = 0; i < userIds.Length; i++)
                    {
                        Sys_UserThuocNhom userThuocNhom = new Sys_UserThuocNhom();
                        userThuocNhom.maNhomUser = nhomUser.maNhomUser;
                        userThuocNhom.userId = Convert.ToInt32(userIds[i]);
                        userThuocNhoms.Add(userThuocNhom);
                    }
                    context.Sys_UserThuocNhoms.InsertAllOnSubmit(userThuocNhoms);
                }

                var nhomUsers = context.Sys_NhomUsers.Where(s => s.maNhomUser != nhomUser.maNhomUser && s.tenNhomUser == nhomUser.tenNhomUser).ToList();
                if (nhomUsers.Count() <= 0)
                {
                    context.SubmitChanges();
                    return RedirectToAction("Index");
                }
                else
                {
                    if (nhomUsers.Where(s => s.maNhomUser.ToLower() == nhomUser.maNhomUser.ToLower()).Count() > 0) TempData["MessgId"] = "Mã nhóm user đã tồn tại";
                    if (nhomUsers.Where(s => s.tenNhomUser.ToLower() == nhomUser.tenNhomUser.ToLower()).Count() > 0) TempData["MessgTen"] = "Tên nhóm user đã tồn tại";
                    return View(nhomUser);
                }
            }
            catch
            {
                return View();
            }
        }

        //
        // POST: /PhanQuyen/Delete/5

        [HttpPost]
        public ActionResult Delete(string[] nhomUser)
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
                var nhomUsers = context.Sys_NhomUsers.Where(s => nhomUser.Contains(s.maNhomUser));
                context.Sys_NhomUsers.DeleteAllOnSubmit(nhomUsers);
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        public ActionResult GetNhanVien(string id)
        {
            ViewBag.Get = true;
            userOfGroups = context.sp_Sys_User_Index(id, null, null,null,null, null).ToList();
            return PartialView("PartialNhanVien", userOfGroups);
        }

        public ActionResult QuyenTruyCap(string id, bool? viewall)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogOn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            try
            {
                ViewBag.NhomUser = context.Sys_NhomUsers.Where(s => s.maNhomUser == id).FirstOrDefault();
                if (viewall == true)
                {
                    var groups = context.Sys_NhomUsers;
                    ViewBag.NhomUsers = new SelectList(groups, "maNhomUser", "tenNhomUser", id);
                    id = null;
                }
                quyenCuaNhomUsers = context.sp_Sys_CongViecCuaUsers(id, null).ToList();
                return View(quyenCuaNhomUsers);
            }
            catch (Exception e)
            {
                return View("Error", e.Message);
            }
        }

        public ActionResult XemChiTietQuyenHan(string congViec, string nhomUser)
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
                ViewBag.XemChiTiet = true;
                vuViecCuaCongViecs = context.sp_Sys_VuViecCuaCongViec(congViec, nhomUser).ToList();
                return PartialView("PartialChiTietQuyenHan", vuViecCuaCongViecs);
            }
            catch
            {
                return View();
            }
        }

        [HttpPost]
        public ActionResult UpdateVuViec(string maVuViec, string maCongViec, string maNhomUser)
        {
            try
            {
                var maCha = context.Sys_CongViecs.Where(s => s.maCongViec == maCongViec)
                                                    .Select(s => s.maCha)
                                                    .FirstOrDefault();
                if (maCha != null && context.Sys_CongViecVaVuViecs.Where(s => s.maCongViec == maCha && s.maNhomUser == maNhomUser).FirstOrDefault() == null)
                {
                    congViecVaVuViec = new Sys_CongViecVaVuViec();
                    congViecVaVuViec.maVuViec = "001";
                    congViecVaVuViec.maNhomUser = maNhomUser;
                    congViecVaVuViec.maCongViec = maCha;
                    context.Sys_CongViecVaVuViecs.InsertOnSubmit(congViecVaVuViec);
                }

                congViecVaVuViec = context.Sys_CongViecVaVuViecs
                                           .Where(s => s.maVuViec == maVuViec
                                               && s.maCongViec == maCongViec
                                               && s.maNhomUser == maNhomUser)
                                           .FirstOrDefault();
                if (congViecVaVuViec != null)
                {
                    context.Sys_CongViecVaVuViecs.DeleteOnSubmit(congViecVaVuViec);
                }
                else
                {
                    congViecVaVuViec = new Sys_CongViecVaVuViec();
                    congViecVaVuViec.maVuViec = maVuViec;
                    congViecVaVuViec.maCongViec = maCongViec;
                    congViecVaVuViec.maNhomUser = maNhomUser;
                    context.Sys_CongViecVaVuViecs.InsertOnSubmit(congViecVaVuViec);
                }
                context.SubmitChanges();
                return Json(new { success = true });
            }
            catch
            {
                return View();
            }
        }

        public ActionResult GetMoreUsers(int? page, string nhomUser, string searchString)
        {
            int currentPageIndex = page.HasValue ? page.Value : 1;
            int? tongSoDong = 0;
            var users = context.sp_Sys_User_Index(null, null, searchString, null,currentPageIndex, 20).ToList();
            try
            {
                ViewBag.Count = users[0].tongSoDong;
                tongSoDong = users[0].tongSoDong;
            }
            catch
            {
                ViewBag.Count = 0;
            }
            TempData["Params"] = nhomUser + "," + searchString;
            return PartialView("PartialUsers", users.ToPagedList(currentPageIndex, 20, true, tongSoDong));
        }
    }
}
