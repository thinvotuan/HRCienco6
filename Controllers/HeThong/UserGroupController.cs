using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Models.HeThong;
using BatDongSan.Helper.Common;

namespace BatDongSan.Controllers.HeThong
{
    public class UserGroupController : ApplicationController
    {
        private LinqHeThongDataContext context = new LinqHeThongDataContext();
        private IQueryable<Sys_CongViec> congViecs;
        private Sys_NhomUser nhomUser;
        private Sys_CongViecVaVuViec congViecVaVV;
        private static string maNhomUser;              
        private bool? permission;
        private readonly string MCV = "UserGroup";

        public ActionResult Index(string congViecCha, string congViecCon)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            if (String.IsNullOrWhiteSpace(congViecCha))
            {
                congViecCha = "HETHONG";
            }                  

            //Nếu công việc cha chỉ có 1 row (không có công việc con) -> set List với điều kiện macv = macv con (for filter by macongvieccon)
            if (context.sp_Sys_NhomUserCongViecVuViec_Index(congViecCha, congViecCon).ToList().Count <= 1)
            {
                ViewBag.UserCVVVs = context.sp_Sys_NhomUserCongViecVuViec_Index(null, null).Where(s => s.maCongViec == congViecCon).ToList();
            }
            else
                ViewBag.UserCVVVs = context.sp_Sys_NhomUserCongViecVuViec_Index(congViecCha, congViecCon).ToList();

            ViewBag.NhomUsers = context.Sys_NhomUsers;
            congViecs = context.Sys_CongViecs.OrderBy(o => o.tenCongViec);
            ViewBag.CongViecs = congViecs;

            //Số lượng vụ viêc
            ViewBag.VuViecCount = context.Sys_VuViecs.Count();

            //Danh sách công việc cha
            var congViecChas = congViecs.Where(s => s.maCha == null);
            ViewBag.CongViecChas = new SelectList(congViecChas, "maCongViec", "tenCongViec", congViecCha);

            //Danh sách công việc con
            IList<Sys_CongViec> congViecCons = congViecs.Where(s => s.maCha == congViecCha).ToList();
            congViecCons.Insert(0, new Sys_CongViec{maCongViec="", tenCongViec=" Tất cả "});
            ViewBag.CongViecCons = new SelectList(congViecCons, "maCongViec", "tenCongViec", congViecCon);

            //Danh sách vụ việc vủa công việc
            ViewBag.VuViecOfCV = context.sp_Sys_VuViecCuaCongViec(null, null).ToList();

            // Get toàn bộ List Nhóm user CV va VV
            IList<sp_Sys_NhomUserCongViecVuViec_IndexResult> NhomUserCVVVs = context.sp_Sys_NhomUserCongViecVuViec_Index(null, null).ToList();            
            return View(NhomUserCVVVs);
        }       

        //
        // GET: /UserGroup/Create

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            return View();
        }

        //
        // POST: /UserGroup/Create

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

            bool flag1 = false;
            bool flag2 = false;
            string tenNhomUser = collection["tennhomuser"];
            maNhomUser = collection["manhomuser"];
            nhomUser = new Sys_NhomUser();
            nhomUser.ghiChu = collection["ghichu"];
            if (context.Sys_NhomUsers.Where(s => s.maNhomUser == maNhomUser).Select(s => s.maNhomUser).FirstOrDefault() == null)
            {
                nhomUser.maNhomUser = maNhomUser;
                flag1 = true;
            }
            else
            {
                flag1 = false;
                TempData["MessgId"] = "Mã nhóm user đã tồn tại";
            }
            if (context.Sys_NhomUsers.Where(s => s.tenNhomUser == tenNhomUser).Select(s => s.tenNhomUser).FirstOrDefault() == null)
            {
                nhomUser.tenNhomUser = tenNhomUser;
                flag2 = true;
            }
            else
            {
                flag2 = false;
                TempData["MessgTen"] = "Tên nhóm user đã tồn tại";
            }
            if (flag1 == true && flag2 == true)
            {
                context.Sys_NhomUsers.InsertOnSubmit(nhomUser);
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            else
                return RedirectToAction("Create");
        }

        //
        // GET: /UserGroup/Edit/5

        public ActionResult Edit(string id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            nhomUser = context.Sys_NhomUsers.Where(s => s.maNhomUser == id).FirstOrDefault();
            return View(nhomUser);
        }

        //
        // POST: /UserGroup/Edit/5

        [HttpPost]
        public ActionResult Edit(string id, FormCollection collection)
        {
            try
            {
                nhomUser = context.Sys_NhomUsers.Where(s => s.maNhomUser == id).FirstOrDefault();
                nhomUser.tenNhomUser = collection["tenNhomUser"];
                nhomUser.ghiChu = collection["ghiChu"];
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }        

        //
        // POST: /UserGroup/Delete/5

        [HttpPost]
        public ActionResult Delete(string id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXoa);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return RedirectToAction("AccessDenied");
            #endregion

            try
            {
                nhomUser = context.Sys_NhomUsers.Where(s => s.maNhomUser == id).FirstOrDefault();
                context.Sys_NhomUsers.DeleteOnSubmit(nhomUser);
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return RedirectToAction("Index");
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
                if (context.Sys_CongViecVaVuViecs.Where(s => s.maCongViec == maCha && s.maNhomUser == maNhomUser).FirstOrDefault() == null)
                {
                    congViecVaVV = new Sys_CongViecVaVuViec();
                    congViecVaVV.maVuViec = "001";
                    congViecVaVV.maNhomUser = maNhomUser;
                    congViecVaVV.maCongViec = maCha;
                    context.Sys_CongViecVaVuViecs.InsertOnSubmit(congViecVaVV);
                }

                congViecVaVV = context.Sys_CongViecVaVuViecs
                                           .Where(s => s.maVuViec == maVuViec
                                               && s.maCongViec == maCongViec
                                               && s.maNhomUser == maNhomUser)
                                           .FirstOrDefault();
                if (congViecVaVV != null)
                {
                    context.Sys_CongViecVaVuViecs.DeleteOnSubmit(congViecVaVV);
                }
                else
                {
                    congViecVaVV = new Sys_CongViecVaVuViec();
                    congViecVaVV.maVuViec = maVuViec;
                    congViecVaVV.maCongViec = maCongViec;
                    congViecVaVV.maNhomUser = maNhomUser;
                    context.Sys_CongViecVaVuViecs.InsertOnSubmit(congViecVaVV);
                }
                context.SubmitChanges();
                return Json(new { success = true });
            }
            catch
            {
                return View();
            }
        }
    }
}
