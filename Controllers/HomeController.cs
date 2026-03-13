using System;
using System.Collections.Generic;
using System.Linq;
using System.Transactions;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using DotNetOpenAuth.AspNet;
using Microsoft.Web.WebPages.OAuth;
using WebMatrix.WebData;
using BatDongSan.Utils;
using BatDongSan.Models.HeThong;
using BatDongSan.Filters;
using BatDongSan.Models.NhanSu;
using BatDongSan.Helper.Utils;
namespace BatDongSan.Controllers
{
    public class HomeController : Controller
    {
        private LinqHeThongDataContext context = new LinqHeThongDataContext();
        private LinqNhanSuDataContext hr = new LinqNhanSuDataContext();
        private Sys_User user;
        public ActionResult Index()
        {
            if (System.Configuration.ConfigurationManager.AppSettings["CongTy"].ToString() == "Cienco6")
            {
                Session["CongTy"] = "Cienco6";
            }
            else
            {
                Session["CongTy"] = "";
            }

            if (Session["CongViecUser"] == null)
            {
                return View("Login");
            }
            else
            {
                var ds = (Sys_User)Session["User"];
                Session["countNhacViec"] = context.sp_HT_NhacViec(ds == null ? string.Empty : ds.manv, null, null).Count();
                // Bat dau thong bao
                var NgayBatDauThongBao1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayBatDauThongBao1").Select(d => d.giaTriNgayThang).FirstOrDefault();
                var NgayKetThucThongBao1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayKetThucThongBao1").Select(d => d.giaTriNgayThang).FirstOrDefault();

                //var NgayBatDauThongBao2 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayBatDauThongBao2").Select(d => d.giaTriNgayThang).FirstOrDefault();
                //var NgayKetThucThongBao2 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayKetThucThongBao2").Select(d => d.giaTriNgayThang).FirstOrDefault();
                if ((DateTime.Now >= NgayBatDauThongBao1 && DateTime.Now <= NgayKetThucThongBao1))
                {
                    return RedirectToAction("ThongBao", "DanhGiaTinNhiem");
                }
                else
                {
                    return View("Index");
                }
            }

        }

        [AllowAnonymous]
        public ActionResult Login(string returnUrl)
        {
            if (System.Configuration.ConfigurationManager.AppSettings["CongTy"].ToString() == "Cienco6")
            {
                Session["CongTy"] = "Cienco6";
            }
            else
            {
                Session["CongTy"] = "";
            }
            if (returnUrl == "undifined") returnUrl = null;
            ViewBag.ReturnUlr = returnUrl;
            return View();
        }
        public ActionResult loadFormLogin(string returnUrl)
        {
            ViewBag.ReturnUlr = returnUrl;
            return PartialView("_FormLoginPartial");
        }
        public int SaveStatusMenu(string statusMenu)
        {
            try
            {

                Session["statusMenu"] = statusMenu;
            }
            catch
            {


            }
            return 1;
        }

        [HttpPost]
        [ValidateInput(false)]
        public int SaveMaCVMenuActive(string MaCVActive)
        {
            try
            {

                Session["MaCVActive"] = MaCVActive;
            }
            catch
            {


            }
            return 1;
        }
        //
        // POST: /Account/Login
        [AllowAnonymous]
        [HttpPost]
        public ActionResult Login(FormCollection collection, string returnUrl)
        {

            string username = collection["username"];
            if (ValidateLogin(username.Trim().ToLower(), collection["password"]) == true)
            {
                //Lưu lịch sử hoạt động

                Sys_LichSuHoatDong lichSu = new Sys_LichSuHoatDong();
                lichSu.controller = GenerateUtil.GetRouteData().Controller;
                lichSu.action = GenerateUtil.GetRouteData().Action;
                lichSu.ngayLap = DateTime.Now;
                lichSu.nguoiLap = user.manv;
                lichSu.noiDung = "Đăng nhập hệ thống";
                context.Sys_LichSuHoatDongs.InsertOnSubmit(lichSu);
                context.SubmitChanges();
                /*END*/
                IList<sp_Sys_CongViecCuaUsersResult> congViecs = context.sp_Sys_CongViecCuaUsers(null, username).ToList();
                Session["CongViecUser"] = congViecs;
                Session["User"] = user;
                Session["manv"] = user.manv;
                Session["countNhacViec"] = context.sp_HT_NhacViec(user == null ? string.Empty : user.manv, null, null).Count();
                var NhanVien = context.GetTable<BatDongSan.Models.NhanSu.tbl_NS_NhanVien>()
                                      .Where(d => d.maNhanVien == user.manv).FirstOrDefault();
                //Session["SanGiaoDich"] = NhanVien.Sys_SanGiaoDich;
                if (NhanVien != null)
                {
                    Session["TenNhanVien"] = NhanVien.ho + " " + NhanVien.ten;
                    Session["HinhAnhNhanVien"] = NhanVien.anhDaiDienURL;
                    Session["avatarMaNV"] = NhanVien.maNhanVien;
                }
                // Show nhung thong bao moi
                if (System.Configuration.ConfigurationManager.AppSettings["CongTy"].ToString() == "Cienco6")
                {
                    Session["CongTy"] = "Cienco6";
                }
                else {
                    Session["CongTy"] = "";
                }
                //var listThongBao = context.sp_Sys_User_ThongBao(user.manv, 2).ToList().Count;
                var getPhongBan = hr.tbl_NS_NhanVienPhongBans.Where(d => d.maNhanVien == user.manv).OrderByDescending(d => d.ngayLap).Select(d => d.maPhongBan).FirstOrDefault();

                var listThongBao = context.sp_Sys_User_ThongBao(user.manv, getPhongBan).Where(d => d.trangThaiXem == 0).ToList().Count;
                Session["listThongBao"] = listThongBao;
                // End show thong bao moi
                //Theme Color
                var listTheme = context.tbl_DM_TienIchThemes.OrderByDescending(d => d.id).ToList();
                Session["listTheme"] = listTheme;

                //Check idColor in table user
                var mauMacDinhUser = context.Sys_Users.Where(d => d.manv == user.manv).Select(d => new { maMau = d.idColor }).ToArray();
                if (mauMacDinhUser != null)
                {
                    if (mauMacDinhUser.FirstOrDefault() != null)
                    {
                        Session["maMau"] = mauMacDinhUser.FirstOrDefault().maMau;
                    }
                }
                if (Session["maMau"] == null)
                {
                    //End check
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
                // Bat dau thong bao
                var NgayBatDauThongBao1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayBatDauThongBao1").Select(d => d.giaTriNgayThang).FirstOrDefault();
                var NgayKetThucThongBao1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayKetThucThongBao1").Select(d => d.giaTriNgayThang).FirstOrDefault();

                //var NgayBatDauThongBao2 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayBatDauThongBao2").Select(d => d.giaTriNgayThang).FirstOrDefault();
                //var NgayKetThucThongBao2 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayKetThucThongBao2").Select(d => d.giaTriNgayThang).FirstOrDefault();

                if ((DateTime.Now >= NgayBatDauThongBao1 && DateTime.Now <= NgayKetThucThongBao1))
                {
                    return RedirectToAction("ThongBao", "DanhGiaTinNhiem");
                }
                if (!String.IsNullOrEmpty(returnUrl))
                {
                    return Redirect(returnUrl);
                }
                else
                {


                    return RedirectToAction("Index");
                }
            }
            else
            {
                Session["CongViecUser"] = null;
                return RedirectToAction("Login");
            }
        }

        //
        // POST: /Account/LogOff

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult LogOut(string returnUrl)
        {
            //Lưu lịch sử hoạt động
            
            Sys_LichSuHoatDong lichSu = new Sys_LichSuHoatDong();
            lichSu.controller = GenerateUtil.GetRouteData().Controller;
            lichSu.action = GenerateUtil.GetRouteData().Action;
            lichSu.ngayLap = DateTime.Now;
            lichSu.nguoiLap = Convert.ToString(Session["manv"]);
            lichSu.noiDung = "Đăng xuất hệ thống";
            context.Sys_LichSuHoatDongs.InsertOnSubmit(lichSu);
            context.SubmitChanges();
            
            /*END*/
            returnUrl = Request.UrlReferrer.PathAndQuery;
            Session.Clear();
            Session.Abandon();
            return RedirectToAction("Login", new { returnUrl = returnUrl });
        }

        //
        // GET: /Account/Register

        [AllowAnonymous]
        public ActionResult Register()
        {
            return View();
        }


        private bool ValidateLogin(string username, string password)
        {
            user = new Sys_User();
            string enryptedPass = HashingHelper.Encrypt(password, true);
            string passde = HashingHelper.Decrypt("z3ilUULG3MBsiRx4JQdiqg==", true);
            user = context.Sys_Users.Where(s => s.userName == username && s.password == enryptedPass && s.status == true).FirstOrDefault();
            if (user != null)
            {
                //check đã nghĩ việc hay chưa
                var checkStatusNV = hr.tbl_NS_NhanViens.Where(d => d.maNhanVien == user.manv).Select(d => d.trangThai).FirstOrDefault() ?? 0;

                if (checkStatusNV == 1)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            else
            {
                TempData["ErrorMessg"] = "Tên đăng nhập hoặc mật khẩu không hợp lệ";
            }
            return false;
        }
        public ActionResult CheckSession()
        {


            if (Session["User"] == null)
            {
                return Json(false);
            }
            return Json(true);


        }
        public ActionResult AccessDenied()
        {
            return View("AccessDenied");
        }

        public JsonResult GetCongViecCanXuLy()
        {
            try
            {
                var ds = (Sys_User)Session["User"];
                int countNhacViec = context.sp_HT_NhacViec(ds.manv, null, null).Count();
                return Json(countNhacViec);
            }
            catch
            {
                return Json(0);
            }
        }

    }
}

