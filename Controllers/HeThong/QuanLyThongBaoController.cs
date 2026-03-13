using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Helper.Utils;
using BatDongSan.Models.DanhMuc;
using BatDongSan.Models.HeThong;
using BatDongSan.Models.NhanSu;
using BatDongSan.Utils.Paging;

namespace BatDongSan.Controllers.NhanSu
{
    public class QuanLyThongBaoController : ApplicationController
    {
        private LinqHeThongDataContext contentHT = new LinqHeThongDataContext();
        private LinqDanhMucDataContext contextDM = new LinqDanhMucDataContext();
        private LinqNhanSuDataContext contextNS = new LinqNhanSuDataContext();
        private IList<sp_ThongBao_IndexResult> thongBaos;
        private IList<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan> phongBans;
     
        private StringBuilder buildTree;
        private int defaultPageSize = 20;
        private readonly string MCV = "QuanLyThongBao";
        private bool? permission;
        //
        // GET: /ThongBao/

        public ActionResult Index(int? page, int? pageSize, string searchString)
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
                buildTree = new StringBuilder();
                phongBans = contextDM.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan>().ToList();
                buildTree = TreePhongBanStyle.BuildTreeDepartment(phongBans);
                ViewBag.PhongBans = buildTree.ToString();

                var getPhongBan = contextNS.tbl_NS_NhanVienPhongBans.Where(d => d.maNhanVien == GetUser().manv).OrderByDescending(d => d.ngayLap).Select(d => d.maPhongBan).FirstOrDefault();

                var listThongBao = contentHT.sp_Sys_User_ThongBao(GetUser().manv, getPhongBan).Where(d=>d.trangThaiXem == 0).ToList().Count;
                Session["listThongBao"] = listThongBao;

                int currentPageIndex = page.HasValue ? page.Value : 1;
                defaultPageSize = pageSize ?? defaultPageSize;
                int? tongSoDong = 0;
                TempData["Params"] = pageSize + "," + searchString;
                thongBaos = contentHT.sp_ThongBao_Index(searchString, currentPageIndex, defaultPageSize).ToList();

                try
                {
                    ViewBag.Count = thongBaos[0].tongSoDong;
                    tongSoDong = thongBaos[0].tongSoDong;
                }
                catch
                {
                    ViewBag.Count = 0;
                }

                if (Request.IsAjaxRequest())
                {
                    return PartialView("PartialIndex", thongBaos.ToPagedList(currentPageIndex, defaultPageSize, true, tongSoDong));
                }
                return View(thongBaos.ToPagedList(currentPageIndex, defaultPageSize, true, tongSoDong));
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
                var getPhongBan = contextNS.tbl_NS_NhanVienPhongBans.Where(d => d.maNhanVien == GetUser().manv).OrderByDescending(d => d.ngayLap).Select(d => d.maPhongBan).FirstOrDefault();

                var listThongBao = contentHT.sp_Sys_User_ThongBao(GetUser().manv, getPhongBan).ToList();
                ViewBag.lsThongBao = listThongBao;
                return View();
            }
            catch (Exception ex)
            {

                ViewBag.Message = ex.Message;
                return View("Error");
            }
        }

        //
        // GET: /ThongBao/Details/5


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
                Sys_ThongBao thongBao = new Sys_ThongBao();

                thongBao = contentHT.Sys_ThongBaos.Where(d => d.id.Equals(id)).FirstOrDefault();
                if (thongBao != null)
                {
                    ViewBag.data = thongBao;

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
                Sys_ThongBao thongBao = new Sys_ThongBao();

                thongBao = contentHT.Sys_ThongBaos.Where(d => d.id.Equals(id)).FirstOrDefault();
                if (thongBao != null)
                {
                    ViewBag.data = thongBao;

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
        // GET: /ThongBao/Create

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
                buildTree = new StringBuilder();
                phongBans = contextDM.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan>().ToList();
                buildTree = TreePhongBanStyle.BuildTreeDepartment(phongBans);
                ViewBag.PhongBans = buildTree.ToString();

                Sys_ThongBao data = new Sys_ThongBao();

                return View("");
            }
            catch (Exception ex)
            {

                ViewBag.Message = ex.Message;
                return View("Error");
            }
        }

        //
        // POST: /ThongBao/Create

        [AcceptVerbs(HttpVerbs.Post)]
        [ValidateInput(false)]
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

                Sys_ThongBao thongBao = new Sys_ThongBao();
                thongBao.tenThongBao = collection.Get("tenThongBao");
                thongBao.noiDungThongBao = collection.Get("noiDungThongBao");
                thongBao.maPhongBan = collection["maPhongBan"];
                thongBao.ngayLap = DateTime.Now;
                thongBao.nguoiLap = GetUser().manv;

                contentHT.Sys_ThongBaos.InsertOnSubmit(thongBao);
                contentHT.SubmitChanges();
                int id = contentHT.Sys_ThongBaos.OrderByDescending(d => d.id).FirstOrDefault().id;
                return RedirectToAction("Index/", new { msg = false });
            }



            catch (Exception ex)
            {

                ViewBag.Message = ex.Message;
                return View("Error");
            }
        }



        //
        // GET: /ThongBao/Edit/5

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
                buildTree = new StringBuilder();
                phongBans = contextDM.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan>().ToList();
                buildTree = TreePhongBanStyle.BuildTreeDepartment(phongBans);
                ViewBag.PhongBans = buildTree.ToString();
                if (msg == true)
                {
                    ViewBag.Message = "Cập nhật dữ liệu thành công";
                }
                else
                {
                    ViewBag.Message = string.Empty;
                }
                Sys_ThongBao thongBao = new Sys_ThongBao();

                thongBao = contentHT.Sys_ThongBaos.Where(d => d.id.Equals(id)).FirstOrDefault();
                if (thongBao != null)
                {
                    ViewBag.data = thongBao;

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
        // POST: /ThongBao/Edit/5

        [AcceptVerbs(HttpVerbs.Post)]
        [ValidateInput(false)]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                Sys_ThongBao thongBao = new Sys_ThongBao();
                thongBao = contentHT.Sys_ThongBaos.Where(d => d.id.Equals(id)).FirstOrDefault();
                thongBao.tenThongBao = collection.Get("tenThongBao");
                thongBao.noiDungThongBao = collection.Get("noiDungThongBao");
                thongBao.maPhongBan = collection["maPhongBan"];

                contentHT.SubmitChanges();
                return RedirectToAction("Edit/" + id, new { msg = true });
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("Error");
            }
        }



        public IQueryable<Sys_ThongBao> FindAll()
        {

            var sql = from p in contentHT.Sys_ThongBaos
                      select p;
            return sql;
        }


        //
        // GET: /ThongBao/Delete/5

        //public ActionResult Delete(int id)
        //{
        //    return View();
        //}

        //
        // POST: /ThongBao/Delete/5

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Delete(int id)
        {
            try
            {
                // TODO: Add delete logic here

                Sys_ThongBao thongBao = new Sys_ThongBao();
                thongBao = contentHT.Sys_ThongBaos.Where(d => d.id.Equals(id)).FirstOrDefault();
                contentHT.Sys_ThongBaos.DeleteOnSubmit(thongBao);
                contentHT.SubmitChanges();

                var User_ThongBao = contentHT.Sys_User_ThongBaos.Where(d => d.idThongBao == id);
                contentHT.Sys_User_ThongBaos.DeleteAllOnSubmit(User_ThongBao);
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
            //Check with manv and idthongbao exist
            var exist = contentHT.Sys_User_ThongBaos.Where(d => d.idThongBao.Equals(id) && d.manv == GetUser().manv).FirstOrDefault();
            if (exist == null)
            {
                //
                Sys_User_ThongBao thongBao = new Sys_User_ThongBao();
                thongBao.idThongBao = id;
                thongBao.manv = GetUser().manv;
                contentHT.Sys_User_ThongBaos.InsertOnSubmit(thongBao);
                contentHT.SubmitChanges();
            }
            var thongTin = contentHT.Sys_ThongBaos.Where(s => s.id == id).FirstOrDefault();
            var getPhongBan = contextNS.tbl_NS_NhanVienPhongBans.Where(d => d.maNhanVien == GetUser().manv).OrderByDescending(d => d.ngayLap).Select(d => d.maPhongBan).FirstOrDefault();

            var listThongBao = contentHT.sp_Sys_User_ThongBao(GetUser().manv, getPhongBan).Where(d=>d.trangThaiXem == 0).ToList().Count;
            Session["listThongBao"] = listThongBao;
            return PartialView("_PartialView", thongTin);
        }

        #region Gửi email thông báo
        [ValidateInput(false)]
        public ActionResult SendEmailThongBao(int id)
        {
            try
            {
                StringBuilder content = new StringBuilder();
                Sys_ThongBao thongBao = contentHT.Sys_ThongBaos.Where(d => d.id == id).FirstOrDefault();

                List<Sys_EmailGuiThongBao> emailThongBaos = new List<Sys_EmailGuiThongBao>();
                var nhanVienList = contextNS.sp_NS_NhanVien_Index2(null, null, thongBao.maPhongBan, 0, 0, null).ToList();

                // var users = contentHT.Sys_Users.Where(s => s.status == true).ToList();
                MailHelper mailInit = new MailHelper();// lay cac tham so trong webconfig                
                foreach (var item in nhanVienList)
                {
                    content.Clear();
                    content.Append("<h3>Email từ hệ thống nhân sự</h3>");
                    content.Append(thongBao.noiDungThongBao);
                    content.Append("<p style='font-style: italic;'>Thanks and Regards!</p>");
                    Sys_EmailGuiThongBao emailThongBao = new Sys_EmailGuiThongBao();
                    emailThongBao.idThongBao = id;
                    emailThongBao.ngayGui = DateTime.Now;
                    emailThongBao.nguoiGui = GetUser().manv;

                    emailThongBao.noiDung = content.ToString();
                    emailThongBao.emailGuiDen = item.email;
                    emailThongBao.maNhanVien = item.maNhanVien;
                    if (SendMailGeneral(thongBao.tenThongBao, item.email, emailThongBao.noiDung))
                    {
                        emailThongBao.trangThaiGui = true;
                    }
                    else
                    {
                        emailThongBao.trangThaiGui = false;
                    }
                    emailThongBaos.Add(emailThongBao);
                }
                thongBao.daGuiEmail = true;
                contentHT.Sys_EmailGuiThongBaos.InsertAllOnSubmit(emailThongBaos);
                contentHT.SubmitChanges();
                return Json(string.Empty);
            }
            catch
            {
                return View();
            }
        }
        #endregion
    }
}
