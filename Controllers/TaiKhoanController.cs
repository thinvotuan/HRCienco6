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
using BatDongSan.Helper.Utils;
using System.Configuration;
using System.Net.Mail;
using BatDongSan.Models.NhanSu;
using BatDongSan.Models.DanhMuc;

namespace BatDongSan.Controllers
{
    public class TaiKhoanController : Controller
    {
        private LinqHeThongDataContext context = new LinqHeThongDataContext();
        private Sys_User user;
        public ActionResult Index()
        {
            if (Session["CongViecUser"] == null)
            {
                return View("Login");
            }
            else
                return View("Index");
        }
        public ActionResult QuenMatKhau()
        {
            if (Session["User"] != null)
            {
                return RedirectToAction("Index", "Home");
            }
            ViewBag.Tieude = "Quên mật khẩu";
            return View();
        }
        public ActionResult ChangePassword()
        {

            return View();

        }
        public ActionResult CheckEmail(string email)
        {
            var hasValue = string.Empty;
            var thongTin = context.Sys_Users.Where(d => d.email == email).FirstOrDefault();
            if (thongTin != null)
            {
                // Email Ton tai trong database
                return Json("1");

            }
            else
            {
                return Json("0");
            }

        }
        public ActionResult CheckEmailForget(string email, string username)
        {
            var hasValue = string.Empty;
            var thongTin = context.Sys_Users.Where(d => d.email == email && d.userName == username).FirstOrDefault();
            if (thongTin != null)
            {
                // Email Ton tai trong database
                return Json("success");

            }
            else
            {
                return View();
            }

        }
        private string GeneratePassword()
        {
            string strPwdchar = "abcdefghijklmnopqrstuvwxyz0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string strPwd = "";
            Random rnd = new Random();
            for (int i = 0; i <= 15; i++)
            {
                int iRandom = rnd.Next(0, strPwdchar.Length - 1);
                strPwd += strPwdchar.Substring(iRandom, 1);
            }
            return strPwd;
        }
        public JsonResult ResetPassword(string email)
        {
            try
            {
                //Get thong tin user
                var thongTin = context.Sys_Users.Where(d => d.email == email).FirstOrDefault();
                if (thongTin != null)
                {
                    string generteString = GeneratePassword();
                    // Insert string to database
                    string appPath = string.Format("{0}://{1}{2}{3}",
                 Request.Url.Scheme,
                 Request.Url.Host,
                  (Request.Url.Port == 80) ? string.Empty : ":" + Request.Url.Port,
                 "/TaiKhoan/ConfirmResetPW?code=" + generteString);
                    string linkReset = appPath;
                    thongTin.codeResetPW = generteString;
                    thongTin.statusResetPW = 0;
                    context.SubmitChanges();
                    bool kq = SendMailReset(thongTin.userName, thongTin.email, linkReset);
                    if (kq == true)
                    {
                        //Lưu lịch sử hoạt động

                        Sys_LichSuHoatDong lichSu = new Sys_LichSuHoatDong();
                        lichSu.controller = GenerateUtil.GetRouteData().Controller;
                        lichSu.action = GenerateUtil.GetRouteData().Action;
                        lichSu.ngayLap = DateTime.Now;
                        lichSu.nguoiLap = Convert.ToString(Session["manv"]);
                        lichSu.noiDung = "Reset mật khẩu cho username: " + thongTin.userName;
                        context.Sys_LichSuHoatDongs.InsertOnSubmit(lichSu);
                        context.SubmitChanges();

                        /*END*/
                        return Json("ok");

                    }
                    return Json("false");
                }


                return Json("false");
            }
            catch (Exception e)
            {
                return Json(e.Message);
            }
        }
        public JsonResult ResetPasswordNew(string email, string username)
        {
            try
            {
                //Get thong tin user
                var thongTin = context.Sys_Users.Where(d => d.email == email && d.userName == username).FirstOrDefault();
                if (thongTin != null)
                {
                    string generteString = GeneratePassword();
                    // Insert string to database
                    string appPath = string.Format("{0}://{1}{2}{3}",
                 Request.Url.Scheme,
                 Request.Url.Host,
                  (Request.Url.Port == 80) ? string.Empty : ":" + Request.Url.Port,
                 "/TaiKhoan/ConfirmResetPW?code=" + generteString);
                    string linkReset = appPath;
                    thongTin.codeResetPW = generteString;
                    thongTin.statusResetPW = 0;
                    context.SubmitChanges();
                    bool kq = SendMailReset(thongTin.userName, thongTin.email, linkReset);
                    var result = new { Result = "Successed" };
                    if (kq == true)
                    {
                        //Lưu lịch sử hoạt động

                        Sys_LichSuHoatDong lichSu = new Sys_LichSuHoatDong();
                        lichSu.controller = GenerateUtil.GetRouteData().Controller;
                        lichSu.action = GenerateUtil.GetRouteData().Action;
                        lichSu.ngayLap = DateTime.Now;
                        lichSu.nguoiLap = Convert.ToString(Session["manv"]);
                        lichSu.noiDung = "Reset mật khẩu cho username: " + thongTin.userName;
                        context.Sys_LichSuHoatDongs.InsertOnSubmit(lichSu);
                        context.SubmitChanges();

                        /*END*/
                        return Json(result, JsonRequestBehavior.AllowGet);
                    }
                    result = new { Result = "Failed" };
                    return Json(result, JsonRequestBehavior.AllowGet);
                }


                return Json("false");
            }
            catch (Exception e)
            {
                return Json(e.Message);
            }
        }

        public bool SendMailReset(string hotenNV, string emailNV, string Code)
        {
            MailHelper mailInit = new MailHelper(); // lay cac tham so trong webconfig
            System.Text.StringBuilder content = new System.Text.StringBuilder();
            content.Append("<h3>Email từ hệ thống nhân sự</h3>");
            content.Append("<p>Xin chào: " + hotenNV + " !</p>");
            content.Append("<p>Bạn vừa yêu cầu reset mật khẩu.</p>");
            content.Append("<p>Vì lý do bảo mật, chúng tôi không gửi mật khẩu của bạn qua E-mail. Để tạo mật khẩu mới, bạn hãy thực hiện theo hướng dẫn bên dưới:</p>");
            content.Append("<p>Vui lòng bấm vào link sau để đặt mật khẩu mới: <strong>" + Code + "</strong>");
            content.Append("<p>Nếu như bạn nhớ mật khẩu và không hề yêu cầu tạo mật khẩu mới, xin hãy bỏ qua E-mail này. </p>");

            //content.Append("<p>Thông tin này được yêu cầu từ IP: " + Request.UserHostAddress + "</p>");
            content.Append("<p style='font-style:italic'>Thanks and Regards!</p>");
            //Send only email is @thuanviet.com.vn
            //string[] array01 = emailNV.ToLower().Split('@');
            // string string2 = ConfigurationManager.AppSettings["OnlySend"]; //get domain from config files
            //string[] array1 = string2.Split(',');
            // bool EmailofThuanViet;
            //EmailofThuanViet = array1.Contains(array01[1]);
            // if (emailNV == "" || emailNV == null || EmailofThuanViet == false)
            // {
            //    return false;
            // }
            MailAddress toMail = new MailAddress(emailNV, hotenNV); // goi den mail
            mailInit.ToMail = toMail;
            mailInit.Body = content.ToString();
            return mailInit.SendMail();
        }
        public ActionResult ConfirmResetPW(string code)
        {
            ViewBag.statusReset = 1;
            ViewBag.MaCode = code;
            var thongTin = context.Sys_Users.Where(d => d.codeResetPW == code && d.statusResetPW == 0).FirstOrDefault();

            if (thongTin != null)
            {

                // Insert string to database



                context.SubmitChanges();
                ViewBag.statusReset = 0;

            }

            return View();
        }

        public JsonResult DatLaiPassword(string maCode, string passwd)
        {
            try
            {
                //Get thong tin user
                var thongTin = context.Sys_Users.Where(d => d.codeResetPW == maCode).FirstOrDefault();
                if (thongTin != null)
                {

                    if (!string.IsNullOrEmpty(passwd))
                    {

                        thongTin.password = HashingHelper.Encrypt(passwd, true);
                    }

                    thongTin.statusResetPW = 1;
                    context.SubmitChanges();
                    //Lưu lịch sử hoạt động

                    Sys_LichSuHoatDong lichSu = new Sys_LichSuHoatDong();
                    lichSu.controller = GenerateUtil.GetRouteData().Controller;
                    lichSu.action = GenerateUtil.GetRouteData().Action;
                    lichSu.ngayLap = DateTime.Now;
                    lichSu.nguoiLap = Convert.ToString(Session["manv"]);
                    lichSu.noiDung = "Đặt lại mật khẩu cho username: " + thongTin.userName;
                    context.Sys_LichSuHoatDongs.InsertOnSubmit(lichSu);
                    context.SubmitChanges();

                    /*END*/
                    if ((string)Session["CongTy"] == "Cienco6")
                    {

                        BatDongSan.Models.ERP.linqERPCienco6DataContext linqERP = new BatDongSan.Models.ERP.linqERPCienco6DataContext();
                        var thongtinNV = linqERP.tbl_NhanVienHRs.Where(d => d.email == thongTin.email).FirstOrDefault();
                        if (thongtinNV != null)
                        {
                            thongtinNV.matKhau = thongTin.password;
                            linqERP.SubmitChanges();
                        }
                        // Cap nhat sang DMS
                        LinqNhanSuDataContext lqNS = new LinqNhanSuDataContext();
                        BatDongSan.Models.DMS.LinqDMSDataContextDataContext DMSContext = new BatDongSan.Models.DMS.LinqDMSDataContextDataContext();
                        var NhanVienHR = (from p in lqNS.vw_NS_DanhSachNhanVienTheoPhongBans
                                          join nv in lqNS.tbl_NS_NhanViens on p.maNhanVien equals nv.maNhanVien
                                          join cp in lqNS.GetTable<tbl_DM_CapBacChucDanh>() on p.SoCapBac equals cp.soCapBac
                                          join q in lqNS.GetTable<Sys_User>() on p.maNhanVien equals q.manv
                                          where p.maNhanVien == thongTin.manv
                                          select new NhanVienModel
                                          {
                                              maNhanVien = p.maNhanVien,
                                              userName = q.userName,
                                              matKhau = q.password,
                                              email = p.email,
                                              maPhongBan = p.maPhongBan,
                                              trangThai = p.trangThai,
                                              tenChucDanh = p.TenChucDanh,
                                              maChucDanh = p.maChucDanh,
                                              soDienThoai = q.telephone,
                                              soCapBac = p.SoCapBac,
                                              tenCapBacQL = cp.tenCapBac,
                                              hoVaTen = nv.ho + " " + nv.ten,
                                              choPhepVaoERPKhac = true
                                          }

                                         ).FirstOrDefault();
                        if (NhanVienHR != null)
                        {
                            var NhanVienUpdate = DMSContext.tbl_DMS_NhanVienHRs.Where(d => d.maNhanVien == thongTin.manv).FirstOrDefault();
                            if (NhanVienUpdate != null)
                            {
                                DMSContext.tbl_DMS_NhanVienHRs.DeleteOnSubmit(NhanVienUpdate);
                            }
                            // Insert new
                            BatDongSan.Models.DMS.tbl_DMS_NhanVienHR tblSysMaC = new BatDongSan.Models.DMS.tbl_DMS_NhanVienHR();
                            tblSysMaC.maNhanVien = NhanVienHR.maNhanVien;
                            tblSysMaC.maCongTrinh = "TP00";
                            tblSysMaC.userName = NhanVienHR.userName;
                            tblSysMaC.matKhau = NhanVienHR.matKhau;
                            tblSysMaC.email = NhanVienHR.email;
                            tblSysMaC.trangThai = NhanVienHR.trangThai;
                            tblSysMaC.tenChucDanh = NhanVienHR.tenChucDanh;
                            tblSysMaC.maChucDanh = NhanVienHR.maChucDanh;
                            tblSysMaC.soDienThoai = NhanVienHR.soDienThoai;
                            tblSysMaC.soCapBac = NhanVienHR.soCapBac;
                            tblSysMaC.tenCapBac = NhanVienHR.tenCapBacQL;
                            tblSysMaC.maPhongBan = NhanVienHR.maPhongBan;
                            tblSysMaC.hoTen = NhanVienHR.hoVaTen;
                            tblSysMaC.dungLuong = 10;
                            tblSysMaC.choPhepVaoERPKhac = NhanVienHR.choPhepVaoERPKhac;
                            DMSContext.tbl_DMS_NhanVienHRs.InsertOnSubmit(tblSysMaC);
                            DMSContext.SubmitChanges();
                            // 
                        }
                        // End
                        BatDongSan.Models.DMS.LinqDMSDataContextDataContext DMSContext2 = new BatDongSan.Models.DMS.LinqDMSDataContextDataContext();
                        var NhanVienHR2 = (from nv in lqNS.tbl_NS_NhanViens
                                           join q in lqNS.GetTable<Sys_User>() on nv.maNhanVien equals q.manv
                                           where nv.maNhanVien == NhanVienHR.maNhanVien
                                           select new NhanVienModel
                                           {
                                               maNhanVien = nv.maNhanVien,
                                               userName = q.userName,
                                               matKhau = q.password


                                           }

                                         ).FirstOrDefault();
                        if (NhanVienHR != null)
                        {
                            var NhanVienUpdate = DMSContext2.tbl_DMS_NhanVienHRs.Where(d => d.maNhanVien == NhanVienHR2.maNhanVien).FirstOrDefault();
                            if (NhanVienUpdate != null)
                            {
                                NhanVienUpdate.matKhau = NhanVienHR.matKhau;
                                var tblUser = DMSContext2.Sys_Users.Where(d => d.manv == NhanVienHR2.maNhanVien).FirstOrDefault();
                                if (tblUser != null)
                                {
                                    tblUser.password = NhanVienHR.matKhau;
                                }
                                DMSContext2.SubmitChanges();
                            }

                        }
                        return Json("ok");

                    }
                }


                return Json("false");
            }
            catch (Exception e)
            {
                return Json(e.Message);
            }
        }


    }
}

