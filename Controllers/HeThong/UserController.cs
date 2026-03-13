using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Models.HeThong;
using BatDongSan.Utils;
using BatDongSan.Helper.Common;
using BatDongSan.Utils.Paging;
using BatDongSan.Helper.Utils;
using System.Configuration;
using System.Net.Mail;
using BatDongSan.Models.NhanSu;
using System.Text;
using BatDongSan.Models.DanhMuc;

namespace BatDongSan.Controllers.HeThong
{
    public class UserController : ApplicationController
    {
        private LinqHeThongDataContext context = new LinqHeThongDataContext();
        private LinqNhanSuDataContext contextNhanSu = new LinqNhanSuDataContext();
        LinqDanhMucDataContext linqDM = new LinqDanhMucDataContext();
        private Sys_UserThuocNhom userThuocNhom;
        private IList<Sys_UserThuocNhom> userThuocNhoms;
        private IList<sp_Sys_User_IndexResult> users;
        private Sys_User user;
        private UserModel model;
        private readonly string MCV = "NguoiSuDung";
        private bool? permission;

        public ActionResult Index(int? page, int? pageSize, string searchString)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion


            ViewBag.UserGroups = context.Sys_NhomUsers.ToList();
            ViewBag.UserOfNhoms = context.sp_Sys_UserThuocNhom_Index().ToList();

            int currentPageIndex = page.HasValue ? page.Value : 1;
            pageSize = pageSize ?? 30;
            int? tongSoDong = 0;
            users = context.sp_Sys_User_Index(null, null, searchString, null, currentPageIndex, pageSize).ToList();
            try
            {
                ViewBag.Count = users[0].tongSoDong;
                tongSoDong = users[0].tongSoDong;
            }
            catch
            {
                ViewBag.Count = 0;
            }
            //Check user administator moi duoc phep cap nhat

            int countNhanSu = context.sp_Sys_User_Index("Admin", null, null, null, null, null).Where(d => d.manv == GetUser().manv).Count();
            ViewBag.Admin = countNhanSu;
            if (Request.IsAjaxRequest())
            {
                return PartialView("PartialContent", users.ToPagedList(currentPageIndex, 30, true, tongSoDong));
            }
            else
                return View(users.ToPagedList(currentPageIndex, 30, true, tongSoDong));
        }

        //
        // GET: /User/Details/5

        public ActionResult Details(int id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXemChiTiet);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            model = (from u in context.Sys_Users
                     join nv in context.GetTable<BatDongSan.Models.NhanSu.tbl_NS_NhanVien>() on u.manv equals nv.maNhanVien into g
                     from a in g.DefaultIfEmpty()
                     where u.userId == id
                     select new UserModel
                     {
                         Email = u.email,
                         HoTenNV = a.ho + " " + a.ten,
                         MaNV = a.maNhanVien,
                         Note = u.note,
                         Password = u.password,
                         Status = u.status ?? false,
                         Telephone = u.telephone,
                         UserId = u.userId,
                         UserName = u.userName
                     }).FirstOrDefault();
            return View(model);
        }

        //
        // GET: /User/Create

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
        // POST: /User/Create

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

                string pw = GeneratePassword();
                user = new Sys_User();
                user.userName = collection["userName"].ToLower();

                user.password = HashingHelper.Encrypt(pw, true);
                user.email = collection["email"];
                user.note = collection["note"];
                user.manv = collection["manv"];


                user.status = collection["status"].Contains("true") ? true : false;


                var checkList = context.Sys_Users.Where(s => s.userName == user.userName || s.manv == user.manv).ToList();
                if (checkList.Count() > 0)
                {
                    if (checkList.Where(s => s.userId == user.userId).FirstOrDefault() != null)
                        TempData["TenDangNhap"] = "Tên đăng nhập đã tồn tại";

                    if (checkList.Where(s => s.manv == user.manv).FirstOrDefault() != null)
                        TempData["MaNhanVien"] = "Nhân viên này đã được đăng ký";

                    return View(user);
                }
                context.Sys_Users.InsertOnSubmit(user);

                // Send email

                var getNhanVien = contextNhanSu.tbl_NS_NhanViens.Where(d => d.maNhanVien == collection["manv"]).FirstOrDefault();
                if (getNhanVien != null)
                {

                    string userName = collection["userName"];
                    try
                    {
                        bool kq = SendMailCreateNewUser(getNhanVien.ho + ' ' + getNhanVien.ten, getNhanVien.email, pw, userName);
                        //End send email
                        if (kq == true)
                        {
                            context.SubmitChanges();
                            var chkNV = context.Sys_Users.Where(d => d.manv == user.manv).FirstOrDefault();
                            if (chkNV != null)
                            {
                                UpdateNhomUser(chkNV.userId, "NhanVien");
                                SaveActiveHistory("Thêm mới user: " + userName);
                                if ((string)Session["CongTy"] == "Cienco6")
                                {
                                    ImportHRToERPMotNhanVien(chkNV.email);
                                }
                            }
                        }
                        else
                        {
                            return View("error");
                        }
                    }
                    catch (Exception ex)
                    {
                        return View(ex);
                    }
                }
                return RedirectToAction("Index");
            }
            catch
            {
                return View(user);
            }
        }

        //
        // GET: /User/Edit/5

        public ActionResult Edit(int id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            model = (from u in context.Sys_Users
                     join nv in context.GetTable<BatDongSan.Models.NhanSu.tbl_NS_NhanVien>() on u.manv equals nv.maNhanVien into g
                     from a in g.DefaultIfEmpty()
                     where u.userId == id
                     select new UserModel
                     {
                         Email = u.email,
                         HoTenNV = a.ho + " " + a.ten,
                         MaNV = a.maNhanVien,
                         Note = u.note,
                         Password = u.password,
                         Status = u.status ?? false,
                         Telephone = u.telephone,
                         UserId = u.userId,
                         UserName = u.userName
                     }).FirstOrDefault();
            return View(model);
        }

        //
        // POST: /User/Edit/5

        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
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
                user = context.Sys_Users.Where(s => s.userId == id).FirstOrDefault();
                user.userName = collection["UserName"];
                user.email = collection["Email"];
                user.note = collection["Note"];
                user.status = collection["Status"].Contains("true") ? true : false;
                // user.telephone = collection["Telephone"];
                context.SubmitChanges();
                SaveActiveHistory("Sửa user: " + user.userName);
                return RedirectToAction("Index");
            }

            catch
            {
                return RedirectToAction("Edit");
            }
        }
        //
        // POST: /User/Delete/5

        [HttpPost]
        public ActionResult Delete(int id)
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
                user = context.Sys_Users.Where(s => s.userId == id).FirstOrDefault();
                if (user.userName == "admin")
                {
                    TempData["Denied"] = "Không được xóa user admin";
                    return RedirectToAction("Index");
                }
                userThuocNhoms = user.Sys_UserThuocNhoms.ToList();
                if (userThuocNhoms.Count() > 0)
                {
                    context.Sys_UserThuocNhoms.DeleteAllOnSubmit(userThuocNhoms);
                    context.SubmitChanges();
                }
                context.Sys_Users.DeleteOnSubmit(user);
                context.SubmitChanges();
                SaveActiveHistory("Xóa userID: " + id + ", username: " + user.userName);
                return RedirectToAction("Index");
            }
            catch
            {
                return RedirectToAction("Index");
            }
        }

        [HttpPost]
        public ActionResult UpdateNhomUser(int userId, string maNhomUser)
        {
            try
            {
                if (maNhomUser == "Admin")
                {
                    int countNhanSu = context.sp_Sys_User_Index("Admin", null, null, null, null, null).Where(d => d.manv == GetUser().manv).Count();
                    if (countNhanSu <= 0)
                    {
                        return Json(new { success = false });
                    }
                }
                userThuocNhom = context.Sys_UserThuocNhoms
                                       .Where(s => s.maNhomUser == maNhomUser && s.userId == userId)
                                       .FirstOrDefault();
                if (userThuocNhom != null)
                    context.Sys_UserThuocNhoms.DeleteOnSubmit(userThuocNhom);
                else
                {
                    userThuocNhom = new Sys_UserThuocNhom();
                    userThuocNhom.userId = userId;
                    userThuocNhom.maNhomUser = maNhomUser;
                    context.Sys_UserThuocNhoms.InsertOnSubmit(userThuocNhom);
                }
                context.SubmitChanges();
                SaveActiveHistory("Cập nhật nhóm user: userID " + userId + " nhóm: " + maNhomUser);
                return Json(new { success = true });
            }
            catch
            {
                return View();
            }
        }

        public ActionResult GetNhanVienInfo(int? page, string searchString)
        {
            ViewBag.isGet = "True";
            int currentPageIndex = page.HasValue ? page.Value : 1;
            ViewBag.SearchString = searchString;
            var nhanViens = context.sp_NS_NhanVien_Index(searchString, null, null, currentPageIndex, 25).ToList();
            ViewBag.Count = nhanViens[0].tongSoDong;
            return PartialView("PartialNhanViens", nhanViens.ToPagedList(currentPageIndex, 25, true, nhanViens[0].tongSoDong));
        }
        /// <summary>
        /// Danh sách phòng ban
        /// </summary>
        /// <returns></returns>
        public ActionResult ChonNhanVien()
        {
            StringBuilder buildTree = new StringBuilder();
            var phongBans = linqDM.tbl_DM_PhongBans.ToList();
            buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
            ViewBag.NVPB = buildTree.ToString();
            return PartialView("_ChonNhanVien");
        }

        /// <summary>
        /// Danh sách nhân viên phòng ban
        /// </summary>
        /// <param name="page"></param>
        /// <param name="searchString"></param>
        /// <param name="maPhongBan"></param>
        /// <returns></returns>
        public ActionResult LoadNhanVien(int? page, string searchString, string maPhongBan)
        {
            IList<sp_PB_DanhSachNhanVienResult> phongBan1s;
            phongBan1s = linqDM.sp_PB_DanhSachNhanVien(searchString, maPhongBan).Where(d => d.userName == null || d.userName == "").ToList();
            ViewBag.isGet = "True";
            int currentPageIndex = page.HasValue ? page.Value : 1;
            ViewBag.Count = phongBan1s.Count();
            ViewBag.Search = searchString;
            ViewBag.MaPhongBan = maPhongBan;
            return PartialView("_LoadNhanVien", phongBan1s.ToPagedList(currentPageIndex, 10));
        }

        /// <summary>
        /// 
        /// </summary>
        private static readonly Random Random = new Random();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="passwordLength"></param>
        /// <param name="strongPassword"></param>
        /// <returns></returns>
        private static string PasswordGenerator(int passwordLength, bool strongPassword)
        {
            int seed = Random.Next(1, int.MaxValue);
            const string allowedChars = "abcdefghijkmnopqrstuvwxyzABCDEFGHJKLMNOPQRSTUVWXYZ0123456789";
            const string specialCharacters = @"!#$%&'()*+,-./:;<=>?@[\]_";

            var chars = new char[passwordLength];
            var rd = new Random(seed);

            for (var i = 0; i < passwordLength; i++)
            {
                // If we are to use special characters
                if (strongPassword && i % Random.Next(3, passwordLength) == 0)
                {
                    chars[i] = specialCharacters[rd.Next(0, specialCharacters.Length)];
                }
                else
                {
                    chars[i] = allowedChars[rd.Next(0, allowedChars.Length)];
                }
            }
            return new string(chars);
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
        public bool SendMailCreateNewUser(string hotenNV, string emailNV, string password, string userName)
        {
            string appPath = string.Format("{0}://{1}{2}{3}",
                 Request.Url.Scheme,
                 Request.Url.Host,
                  (Request.Url.Port == 80) ? string.Empty : ":" + Request.Url.Port, "");

            MailHelper mailInit = new MailHelper(); // lay cac tham so trong webconfig
            System.Text.StringBuilder content = new System.Text.StringBuilder();
            content.Append("<h3>Email từ hệ thống nhân sự</h3>");
            content.Append("<p><span style='color:rgb(83, 151, 51)'>Xin chào: </span>" + hotenNV + " !</p>");
            content.Append("<p>Tài khoản để đăng nhập vào hệ thống nhân sự.</p>");
            content.Append("<p>Tên đăng nhập: " + userName + "</p>");
            content.Append("<p>Mật khẩu: " + password + "");
            content.Append("<p><span style='color:rgb(83, 151, 51)'>Website truy cập: </span><a href='" + appPath + "' target='_blank'>" + appPath + "</a></p>");
            content.Append("<p><span style='color:rgb(83, 151, 51)'>Vui lòng đổi mật khẩu sau khi truy cập!</span></p>");
            content.Append("<p style='font-style:italic; color:rgb(196, 22, 28)'>Thanks and Regards!</p>");
            //Send only email is @thuanviet.com.vn

            //string[] array01 = emailNV.ToLower().Split('@');
            //string string2 = ConfigurationManager.AppSettings["OnlySend"]; //get domain from config files
            //string[] array1 = string2.Split(',');
            MailAddress toMail = new MailAddress(emailNV, hotenNV); // goi den mail
            mailInit.ToMail = toMail;
            mailInit.Body = content.ToString();
            return mailInit.SendMail();
        }

        public bool SendMailResetPass(string hotenNV, string emailNV, string password)
        {
            MailHelper mailInit = new MailHelper(); // lay cac tham so trong webconfig
            System.Text.StringBuilder content = new System.Text.StringBuilder();
            content.Append("<h3>Email từ hệ thống nhân sự</h3>");
            content.Append("<p>Xin chào: " + hotenNV + " !</p>");
            content.Append("<p>Mật khẩu đăng nhập hệ thống là: <b>" + password + "</b>");
            //content.Append("<p>Link truy cập: " + mailInit.Url + " !</p>");
            content.Append("<p>Thông tin này được yêu cầu từ IP: " + Request.UserHostAddress + "</p>");
            content.Append("<p style='font-style:italic'>Thanks and Regards!</p>");
            //Send only email is @thuanviet.com.vn
            //string[] array01 = emailNV.ToLower().Split('@');
            //string string2 = ConfigurationManager.AppSettings["OnlySend"]; //get domain from config files
            //string[] array1 = string2.Split(',');
            //bool EmailofThuanViet;
            //EmailofThuanViet = array1.Contains(array01[1]);
            //if (emailNV == "" || emailNV == null || EmailofThuanViet == false)
            //{
            //    return false;
            //}
            MailAddress toMail = new MailAddress(emailNV, hotenNV); // goi den mail
            mailInit.ToMail = toMail;
            mailInit.Body = content.ToString();
            return mailInit.SendMail();
        }

        [HttpPost]
        public ActionResult ResetPassword(int id)
        {
            try
            {
                //string newPassword = PasswordGenerator(10, true);
                string pw = GeneratePassword();
                user = context.Sys_Users.Where(s => s.userId == id).FirstOrDefault();
                user.password = HashingHelper.Encrypt(pw, true);

                context.SubmitChanges();
                string tenNhanVien = user.userName;
                SendMailResetPass(tenNhanVien, user.email, pw);
                if ((string)Session["CongTy"] == "Cienco6")
                {
                    CapNhatERPCienco6(user.email, user.password);
                    UpdatePasswordToDMS(user.manv);
                }
                return Json(new { Success = true });
            }
            catch
            {
                return View();
            }
        }

        [HttpPost]
        public ActionResult LoadMore(int offset, int fetchNext, string searchString)
        {
            try
            {
                users = context.sp_Sys_User_Index(null, null, searchString, null, offset, fetchNext).ToList();
                ViewBag.UserGroups = context.Sys_NhomUsers.ToList();
                ViewBag.UserOfNhoms = context.sp_Sys_UserThuocNhom_Index().ToList();
                ViewBag.OffSet = offset;
                return PartialView("PartialContent", users);
            }
            catch
            {
                return View();
            }
        }
        public JsonResult ProcesschangePassword(string passwd, string oldpasswd)
        {
            try
            {

                string username = GetUser().userName;
                if (username != "")
                {
                    string oldpasswd2 = HashingHelper.Encrypt(oldpasswd, true);
                    //Get thong tin user
                    var thongTin = context.Sys_Users.Where(d => d.userName == username && d.password == oldpasswd2).FirstOrDefault();
                    if (thongTin != null)
                    {

                        if (!string.IsNullOrEmpty(passwd))
                        {

                            thongTin.password = HashingHelper.Encrypt(passwd, true);
                        }



                        context.SubmitChanges();
                        SaveActiveHistory("Thay đổi mật khẩu: " + GetUser().userName);
                        if ((string)Session["CongTy"] == "Cienco6")
                        {
                            CapNhatERPCienco6(thongTin.email, thongTin.password);
                            UpdatePasswordToDMS(GetUser().manv);
                            SaveActiveHistory("Cập nhật mật khẩu đến DMS: " + GetUser().userName);
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

        #region Danh sách nhắc việc duyệt phiếu
        public ActionResult ListCongViecProcess(int? page, int? pageSize, string searchString, string maCongViec)
        {
            //#region Role user
            //permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            //if (!permission.HasValue)
            //    return View("LogIn");
            //if (!permission.Value)
            //    return View("AccessDenied");
            //#endregion
            try
            {
                BindDataCongViec(maCongViec);
                searchString = searchString ?? "";
                pageSize = pageSize ?? 30;
                int currentPageIndex = page.HasValue ? page.Value : 1;
                int offset = page * (page - 1) ?? 0;
                int fetchNext = pageSize ?? 30;

                var nhacViecs = context.sp_HT_NhacViec(GetUser().manv, maCongViec, searchString).Skip(offset).Take(fetchNext).ToList();
                ViewBag.Count = context.sp_HT_NhacViec(GetUser().manv, maCongViec, searchString).Count();
                TempData["Params"] = searchString + "," + (maCongViec ?? string.Empty);
                if (Request.IsAjaxRequest())
                {
                    return PartialView("PartialIndex", nhacViecs.ToPagedList(currentPageIndex, fetchNext, true, (int)ViewBag.Count));
                }
                return View(nhacViecs.ToPagedList(currentPageIndex, fetchNext, true, (int)ViewBag.Count));
            }
            catch
            {
                return View("Error");
            }
        }
        #endregion

        #region Binddata

        private void BindDataCongViec(string maCongViec)
        {
            var dsCongViec = (from cd in context.tbl_HT_CongViecCanDuyets
                              join cv in context.Sys_CongViecs on cd.maCongViec equals cv.maCongViec
                              select new
                              {
                                  maCongViec = cv.maCongViec,
                                  tenCongViec = cv.tenCongViec,
                              }).Distinct().ToList();

            Dictionary<string, string> dict = new Dictionary<string, string>();
            dict.Add("", "[Chọn]");

            foreach (var item in dsCongViec)
            {
                dict.Add(item.maCongViec, item.tenCongViec);
            }
            ViewBag.CongViec = new SelectList(dict, "Key", "Value", maCongViec);

        }

        #endregion

    }
}
