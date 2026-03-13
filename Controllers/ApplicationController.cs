using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Mail;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;
using BatDongSan.Helper.Utils;
using BatDongSan.Models.HeThong;
using BatDongSan.Models.NhanSu;
using BatDongSan.Models.PhieuDeNghi;
using BatDongSan.Models.ERP;
using BatDongSan.Models.DanhMuc;
using BatDongSan.Models.QLSuatAn;
using System.Drawing;
using BatDongSan.Models.BarcodeLib;

namespace BatDongSan.Controllers
{
    public class ApplicationController : Controller
    {
        LinqHeThongDataContext lqHeThong = new LinqHeThongDataContext();
        LinqNhanSuDataContext lqNS = new LinqNhanSuDataContext();
        LinqPhieuDeNghiDataContext lqPhieuDN = new LinqPhieuDeNghiDataContext();
        public int start;
        public int offset;

        private Sys_User user;
        private VuViecRepository vuViecServices = new VuViecRepository();
        protected override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            if (Session["CongViecUser"] == null || Session["User"] == null)
            {
                filterContext.Result = new RedirectToRouteResult(new RouteValueDictionary(new
                {
                    controller = "Home",
                    action = "LogIn",
                    returnUrl = Request.CurrentExecutionFilePath
                }));
            }
            else
            {
                return;
            }
        }
        public Boolean CheckQuyetInPhieuSA(int UserID, string maNhomUser)
        {
            QLSuatAnDataContext linqCD = new QLSuatAnDataContext();
            Boolean flash = false;
            var checkQuyen = linqCD.GetTable<Sys_UserThuocNhom>().Where(d => d.maNhomUser == maNhomUser && d.userId == UserID).FirstOrDefault();
            if (checkQuyen != null)
            {
                flash = true;
            }
            return flash;
        }
        #region Tạo mã vạch tự động
        public void GenerateBarcodeImage(string maPhieu, string maNhanVien)
        {
            try
            {
                QLSuatAnDataContext linqCuDan = new QLSuatAnDataContext();
                string Forecolor = "000000";
                string Backcolor = "FFFFFF";
                Image barcodeImage = null;
                Barcode b = new Barcode();
                TYPE type = new TYPE();
                b.BarWidth = 1;
                b.AspectRatio = 1.5f;
                type = BatDongSan.Models.BarcodeLib.TYPE.CODE128;
                barcodeImage = b.Encode(type, maPhieu.Trim(), System.Drawing.ColorTranslator.FromHtml("#" + Forecolor), System.Drawing.ColorTranslator.FromHtml("#" + Backcolor), 520, 300);
                string tenFileDinhKem = maPhieu + "_" + HoVaTen(maNhanVien);
                string tenFileDinhKemLuu = string.Empty;
                if (maPhieu.Contains(" "))
                {

                    tenFileDinhKemLuu = Guid.NewGuid().ToString() + ".png";
                }
                else
                {
                    tenFileDinhKemLuu = maPhieu + ".png";
                }

                string path = "~/FileUploads/BarcodeImage";
                string endCode = b.EncodedValue;
                tbl_HT_MaVach maVach = new tbl_HT_MaVach();
                maVach.maPhieu = maPhieu;
                maVach.tenFileDinhKem = tenFileDinhKem;
                maVach.tenFileDinhKemLuu = tenFileDinhKemLuu;
                maVach.enCoded = endCode;
                linqCuDan.tbl_HT_MaVaches.InsertOnSubmit(maVach);
                b.SaveImage(HttpContext.Server.MapPath(path + "/" + tenFileDinhKemLuu), BatDongSan.Models.BarcodeLib.SaveTypes.PNG);
                linqCuDan.SubmitChanges();
            }
            catch
            {

            }
            //(HttpContext.Current.Server.MapPath(path + "/" + filename), BarcodeLib.SaveTypes.PNG);
        }
        #endregion
        public int CapNhatERPCienco6(string email, string password)
        {
            try
            {
                if ((string)Session["CongTy"] == "Cienco6")
                {
                    BatDongSan.Models.ERP.linqERPCienco6DataContext linqERP = new BatDongSan.Models.ERP.linqERPCienco6DataContext();
                    var thongtinNV = linqERP.tbl_NhanVienHRs.Where(d => d.email == email).FirstOrDefault();
                    if (thongtinNV != null)
                    {
                        thongtinNV.matKhau = password;
                        linqERP.SubmitChanges();

                    }
                }
            }
            catch
            {

            }
            return 1;
        }
        public int ImportHRToERPMotNhanVien(string email)
        {
            if ((string)Session["CongTy"] == "Cienco6")
            {

                linqERPCienco6DataContext linqERP = new linqERPCienco6DataContext();

                // update nvCienco6
                LinqNhanSuDataContext linqNS = new LinqNhanSuDataContext();
                var listNVNew = (from p in linqNS.vw_NS_DanhSachNhanVienTheoPhongBans
                                 join nv in linqNS.tbl_NS_NhanViens on p.maNhanVien equals nv.maNhanVien
                                 join cp in linqNS.GetTable<tbl_DM_CapBacChucDanh>() on p.SoCapBac equals cp.soCapBac
                                 join q in linqNS.GetTable<Sys_User>() on p.maNhanVien equals q.manv
                                 where nv.trangThai == 0 && q.userName != "admin" && nv.email == email
                                 select new
                                 {
                                     maNhanVien = p.maNhanVien + "C6",
                                     userName = p.email,
                                     matKhau = q.password,
                                     email = p.email,
                                     maPhongBan = p.maPhongBan,
                                     tenPhongBan = p.tenPhongBan + "C6",
                                     trangThai = nv.trangThai,
                                     tenChucDanh = p.TenChucDanh + "C6",
                                     maChucDanh = p.maChucDanh,
                                     soDienThoai = q.telephone,
                                     soCapBac = p.SoCapBac,
                                     tenCapBacQL = cp.tenCapBac,
                                     hoVaTen = nv.ho + " " + nv.ten,
                                     ten = nv.ten,
                                     ho = nv.ho,
                                     gioiTinh = nv.gioiTinh,
                                     ngaySinh = nv.ngaySinh,
                                     noiSinh = nv.noiSinh,
                                     tinhTrangHonNhan = nv.tinhTrangHonNhan,
                                     phoneNumber1 = nv.phoneNumber1,
                                 }

                                ).ToList();
                // Delete nhan vien cienco6
                var tblDelet_NhanVien = linqERP.tbl_NhanVienHRs.Where(d => d.email == email).ToList();
                if (tblDelet_NhanVien.Count > 0)
                {
                    linqERP.tbl_NhanVienHRs.DeleteAllOnSubmit(tblDelet_NhanVien);
                }
                var tblDelet_TBL_NhanViens = linqERP.NS_TBL_NhanViens.Where(d => d.Email == email).ToList();
                if (tblDelet_TBL_NhanViens.Count > 0)
                {
                    linqERP.NS_TBL_NhanViens.DeleteAllOnSubmit(tblDelet_TBL_NhanViens);
                }
                var checkHTUsers = linqERP.HT_Users.Where(d => d.email == email).ToList();
                if (checkHTUsers.Count > 0)
                {
                    linqERP.HT_Users.DeleteAllOnSubmit(checkHTUsers);

                }

                linqERP.SubmitChanges();

                foreach (var item in listNVNew)
                {



                    // insert new

                    tbl_NhanVienHR tblNhanVienERP = new tbl_NhanVienHR();
                    tblNhanVienERP.maNhanVien = item.maNhanVien;
                    tblNhanVienERP.maCongTrinh = "TP00";
                    tblNhanVienERP.userName = item.userName;
                    tblNhanVienERP.matKhau = item.matKhau;
                    tblNhanVienERP.email = item.email;
                    tblNhanVienERP.trangThai = item.trangThai;
                    tblNhanVienERP.tenChucDanh = item.tenChucDanh;
                    tblNhanVienERP.maChucDanh = item.maChucDanh;
                    tblNhanVienERP.soDienThoai = item.soDienThoai;
                    tblNhanVienERP.soCapBac = item.soCapBac;
                    tblNhanVienERP.tenCapBac = item.tenCapBacQL;
                    tblNhanVienERP.maPhongBan = item.maPhongBan;
                    tblNhanVienERP.hoTen = item.hoVaTen;
                    tblNhanVienERP.ngaySinh = item.ngaySinh;
                    tblNhanVienERP.trangThai = 0;
                    tblNhanVienERP.nvCienco6 = 1;
                    linqERP.tbl_NhanVienHRs.InsertOnSubmit(tblNhanVienERP);
                    linqERP.SubmitChanges();
                    // Insert table NS_TBL_NhanVien
                    NS_TBL_NhanVien tblNhanVien = new NS_TBL_NhanVien();
                    tblNhanVien.Manv = item.maNhanVien;
                    tblNhanVien.Ho = item.ho;
                    tblNhanVien.Ten = item.ten;
                    tblNhanVien.DienThoai = item.soDienThoai;
                    tblNhanVien.Email = item.email;
                    tblNhanVien.nvCienco6 = 1;
                    linqERP.NS_TBL_NhanViens.InsertOnSubmit(tblNhanVien);
                    linqERP.SubmitChanges();
                    // Check HT_us
                    var checkHTUser = linqERP.HT_Users.Where(d => d.email == item.email).FirstOrDefault();
                    if (checkHTUser == null)
                    {
                        // Insert table HT_Users
                        HT_User htUser = new HT_User();
                        htUser.userName = item.userName;
                        htUser.password = item.matKhau;
                        htUser.email = item.email;
                        htUser.telephone = item.soDienThoai;
                        htUser.note = "Cap nhat tu nhan su Cienco6";
                        htUser.status = true;
                        htUser.manv = item.maNhanVien;
                        linqERP.HT_Users.InsertOnSubmit(htUser);
                        linqERP.SubmitChanges();
                    }

                    // Check NS_TBL_PhongBanNhanVien
                    var checkNV_PhongBan = linqERP.NS_TBL_PhongBanNhanViens.Where(d => d.Manv == item.maNhanVien).FirstOrDefault();
                    if (checkNV_PhongBan == null)
                    {
                        // Insert table NS_TBL_PhongBanNhanVien
                        NS_TBL_PhongBanNhanVien tblPhongBanNV = new NS_TBL_PhongBanNhanVien();
                        tblPhongBanNV.Manv = item.maNhanVien;
                        tblPhongBanNV.NgayCapNhat = DateTime.Now;
                        tblPhongBanNV.MaPhongBan = 60;
                        linqERP.NS_TBL_PhongBanNhanViens.InsertOnSubmit(tblPhongBanNV);
                        linqERP.SubmitChanges();
                    }

                    var checkNV_ChucDanh = linqERP.NS_TBL_ChucDanhNhanViens.Where(d => d.Manv == item.maNhanVien).FirstOrDefault();
                    if (checkNV_ChucDanh == null)
                    {
                        // Insert table NS_TBL_ChucDanhNhanVien
                        NS_TBL_ChucDanhNhanVien tblChucDanh = new NS_TBL_ChucDanhNhanVien();
                        tblChucDanh.Manv = item.maNhanVien;
                        tblChucDanh.MaChucDanh = "NV";
                        tblChucDanh.NgayCapNhat = DateTime.Now;
                        linqERP.NS_TBL_ChucDanhNhanViens.InsertOnSubmit(tblChucDanh);
                        linqERP.SubmitChanges();
                    }

                }
            }
            return 1;
        }

        #region
        // Cap nhat password to DMS
        public int UpdatePasswordToDMS(string maNhanVien)
        {
            try
            {
                if ((string)Session["CongTy"] == "Cienco6")
                {

                    BatDongSan.Models.DMS.LinqDMSDataContextDataContext DMSContext = new BatDongSan.Models.DMS.LinqDMSDataContextDataContext();
                    var NhanVienHR = (from nv in lqNS.tbl_NS_NhanViens
                                      join q in lqNS.GetTable<Sys_User>() on nv.maNhanVien equals q.manv
                                      where nv.maNhanVien == maNhanVien
                                      select new NhanVienModel
                                      {
                                          maNhanVien = nv.maNhanVien,
                                          userName = q.userName,
                                          matKhau = q.password


                                      }

                                     ).FirstOrDefault();
                    if (NhanVienHR != null)
                    {
                        var NhanVienUpdate = DMSContext.tbl_DMS_NhanVienHRs.Where(d => d.maNhanVien == maNhanVien).FirstOrDefault();
                        if (NhanVienUpdate != null)
                        {
                            NhanVienUpdate.matKhau = NhanVienHR.matKhau;
                            var tblUser = DMSContext.Sys_Users.Where(d => d.manv == maNhanVien).FirstOrDefault();
                            if (tblUser != null)
                            {
                                tblUser.password = NhanVienHR.matKhau;
                            }
                            DMSContext.SubmitChanges();
                        }

                    }
                }
                return 1;
            }
            catch
            {
                return 0;
            }
        }
        //Cap nhat erp to DMS
        public int UpdateNhanVienToDMS(string maNhanVien)
        {
            BatDongSan.Models.DMS.LinqDMSDataContextDataContext DMSContext = new BatDongSan.Models.DMS.LinqDMSDataContextDataContext();
            var NhanVienHR = (from p in lqNS.vw_NS_DanhSachNhanVienTheoPhongBans
                              join nv in lqNS.tbl_NS_NhanViens on p.maNhanVien equals nv.maNhanVien
                              join cp in lqNS.GetTable<tbl_DM_CapBacChucDanh>() on p.SoCapBac equals cp.soCapBac
                              join q in lqNS.GetTable<Sys_User>() on p.maNhanVien equals q.manv
                              where p.maNhanVien == maNhanVien
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
                var NhanVienUpdate = DMSContext.tbl_DMS_NhanVienHRs.Where(d => d.maNhanVien == maNhanVien).FirstOrDefault();
                if (NhanVienUpdate != null)
                {
                    DMSContext.tbl_DMS_NhanVienHRs.DeleteOnSubmit(NhanVienUpdate);
                    DMSContext.SubmitChanges();
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
                tblSysMaC.choPhepVaoERPKhac = NhanVienHR.choPhepVaoERPKhac;
                tblSysMaC.dungLuong = 10;
                DMSContext.tbl_DMS_NhanVienHRs.InsertOnSubmit(tblSysMaC);
                DMSContext.SubmitChanges();
                // 
            }
            //SaveActiveHistory("Cập nhật nhân viên DMS: " + maNhanVien);
            return 1;
        }
        #endregion
        public bool? GetPermission(string maCongViec, string maVuViec)
        {
            user = GetUser();
            if (user == null)
                return null;
            else
            {
                IList<Sys_VuViec> vuViecs = vuViecServices.GetVuViecByCongViecVaUserName(user.userName, maCongViec).ToList();
                ViewBag.VuViecAccess = vuViecs;
                if (vuViecs.Where(d => d.maVuViec == maVuViec).FirstOrDefault() == null)
                    return false;
            }
            return true;
        }

        //GET: User session state
        public Sys_User GetUser()
        {
            Sys_User user = new Sys_User();
            if (Session["User"] != null)
            {
                user = (Sys_User)Session["User"];
                return user;
            }
            return null;
        }

        /// <summary>
        /// Get quyền user theo công việc & vụ việc
        /// </summary>
        /// <param name="maCongViec"></param>
        /// <param name="maVuViec"></param>
        /// <returns>null:hết session(View("LogOn")), false: ko có quyền(View("AccessDenied")), true: có quyền(tiếp tục)</returns>        
        public bool? GetQuyen(string maCongViec, string maVuViec, string UserName)
        {
            IList<Sys_VuViec> vuViecs = vuViecServices.GetVuViecByCongViecVaUserName(UserName, maCongViec).ToList();
            ViewBag.VuViecAccess = vuViecs;
            if (vuViecs.Where(d => d.maVuViec.Equals(maVuViec)) == null)
                return false;
            return true;
        }

        public ActionResult AccessDenied()
        {
            return View();
        }

        public void PagingLoaderController(string url, int total, int page, string param)
        {
            var TotalRow = total;
            var currentPage = page;
            var PerPage = 100;
            var Params = param;
            var totalPage = (TotalRow / PerPage) + (TotalRow % PerPage > 0 ? 1 : 0);
            start = (page - 1) * PerPage;
            offset = PerPage;
            ViewData["page"] = page;
            ViewData["total"] = totalPage;
            ViewData["totalRow"] = TotalRow;
            ViewData["startIndex"] = start;
        }
        public void PagingLoaderFullController(string url, int total, int page, string param)
        {
            var TotalRow = total;
            var currentPage = page;
            var PerPage = 100;
            var Params = param;
            var totalPage = (TotalRow / PerPage) + (TotalRow % PerPage > 0 ? 1 : 0);
            start = (page - 1) * PerPage;
            offset = PerPage;
            ViewData["page"] = page;
            ViewData["total"] = totalPage;
            ViewData["totalRow"] = TotalRow;
            ViewData["startIndex"] = start;
        }

        public void BindDataTrangThai(string maCongViec)
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();
            dict.Add("", "[Chọn]");
            dict.Add("0", "Tạo mới");
            foreach (var item in lqHeThong.sp_QuiTrinhDuyet_ListBuocDuyet(maCongViec).ToList())
            {
                dict.Add(item.Id.ToString(), item.TenBuocDuyet);
            }
            ViewBag.TrangThaiDuyet = new SelectList(dict, "Key", "Value", string.Empty);
        }


        public bool SendMailGeneral(string tieuDe, string emailNV, string content)
        {
            MailHelper mailInit = new MailHelper(); // lay cac tham so trong webconfig            
            ////Send only email is @thuanviet.com.vn
            //string[] array01 = emailNV.ToLower().Split('@');
            //string string2 = ConfigurationManager.AppSettings["OnlySend"]; //get domain from config files
            //string[] array1 = string2.Split(',');
            //bool EmailofThuanViet;
            //EmailofThuanViet = array1.Contains(array01[1]);
            //if (emailNV == "" || emailNV == null || EmailofThuanViet == false)
            //{
            //    return false;
            //}
            MailAddress toMail = new MailAddress(emailNV, string.Empty); // goi den mail
            mailInit.ToMail = toMail;
            mailInit.Body = content;
            mailInit.Subject = tieuDe;
            return mailInit.SendMail();
        }


        public void NhanVienQLNSDuyet()
        {
            ViewBag.NhanVienThuocNhoms = lqHeThong.sp_Sys_User_Index("QLChamCong", null,null,null, null, null).Select(d => d.manv).ToArray();
        }
        #region Lưu lịch sử hoạt động
        public void SaveActiveHistory(string noiDung)
        {
            Sys_LichSuHoatDong lichSu = new Sys_LichSuHoatDong();
            lichSu.controller = GenerateUtil.GetRouteData().Controller;
            lichSu.action = GenerateUtil.GetRouteData().Action;
            lichSu.ngayLap = DateTime.Now;
            lichSu.nguoiLap = GetUser().manv;
            lichSu.noiDung = noiDung;
            lqHeThong.GetTable<Sys_LichSuHoatDong>().InsertOnSubmit(lichSu);
            lqHeThong.SubmitChanges();
        }
        #endregion

        #region Họ và tên nhân viên
        public string HoVaTen(string maNhanVien)
        {
            try
            {
                return ((lqHeThong.GetTable<BatDongSan.Models.NhanSu.tbl_NS_NhanVien>().Where(d => d.maNhanVien == maNhanVien).Select(d => d.ho).FirstOrDefault() ?? string.Empty) + " " + (lqHeThong.GetTable<BatDongSan.Models.NhanSu.tbl_NS_NhanVien>().Where(d => d.maNhanVien == maNhanVien).Select(d => d.ten).FirstOrDefault() ?? string.Empty));
            }
            catch
            {
                return string.Empty;
            }
        }
        #endregion

        #region Những user thuộc nhóm Admin nhân sự thì sẽ được quyền tạo trực tiếp phiếu công tác, phiếu nghỉ phép, phiếu tăng ca

        public string AdminNhanSu(string maNhanVien)
        {
            try
            {
                int countNhanSu = lqHeThong.sp_Sys_User_Index("AdminNhanSu", null,null,null, null, null).Where(d => d.manv == maNhanVien).Count();
                if (countNhanSu > 0)
                {
                    ViewBag.AdminNhanSu = "true";
                    return "true";
                }
                else
                {
                    ViewBag.AdminNhanSu = "false";
                    return "false";
                }
            }
            catch
            {
                ViewBag.AdminNhanSu = "false";
                return "true";
            }
        }

        public string Administrator(string maNhanVien)
        {
            try
            {
                int countNhanSu = lqHeThong.sp_Sys_User_Index("Admin", null, null, null, null,null).Where(d => d.manv == maNhanVien).Count();
                if (countNhanSu > 0)
                {
                    ViewBag.Administrator = "true";
                    return "true";
                }
                else
                {
                    ViewBag.Administrator = "false";
                    return "false";
                }
            }
            catch
            {
                ViewBag.Administrator = "false";
                return "true";
            }
        }
        #endregion

        #region Lưu lịch sử cập nhật phiếu và delete của phiếu công tác, phiếu nghỉ phép, phiếu tăng ca
        public void LuuLichSuCapNhatPhieu(string maPhieu, string maCongViec, int trangThai)
        {
            try
            {
                tbl_NS_LichSuCapNhatPhieu lichSu = new tbl_NS_LichSuCapNhatPhieu();
                string noiDung = string.Empty;
                if (maCongViec == "PhieuCongTac")
                {
                    var dsPhieuCongTac = lqPhieuDN.tbl_NS_PhieuCongTacs.Where(d => d.maPhieu == maPhieu).FirstOrDefault();
                    lichSu.maPhieu = maPhieu;
                    lichSu.maCongViec = maCongViec;
                    lichSu.ngayCapNhat = DateTime.Now;
                    lichSu.maNVCapNhat = GetUser().manv;
                    lichSu.trangThai = trangThai;
                    lichSu.maNhanVienLapPhieu = dsPhieuCongTac.maNhanVien;
                    noiDung = "Phòng ban công tac: " + dsPhieuCongTac.phongBanCongTac + ", Ngày bắt đầu: " + dsPhieuCongTac.ngayBatDau.Value.ToString("dd/MM/yyyy") + ", Ngày kết thúc: " + dsPhieuCongTac.ngayKetThuc.Value.ToString("dd/MM/yyyy") + " , Giờ bắt đầu: " + dsPhieuCongTac.gioBatDau.Value.ToString("HH\\:mm") + " , Giờ kết thúc: " + dsPhieuCongTac.gioKetThuc.Value.ToString("HH\\:mm") + ", Phụ cấp công tác: " + dsPhieuCongTac.phuCap + ", Nội dung: " + dsPhieuCongTac.ghiChu + "";
                    lichSu.noiDung = noiDung;
                    lqNS.tbl_NS_LichSuCapNhatPhieus.InsertOnSubmit(lichSu);
                    lqNS.SubmitChanges();
                }
                else if (maCongViec == "PhieuNghiPhep")
                {
                    var dsPhieuNghiPhep = lqPhieuDN.tbl_NS_PhieuNghiPheps.Where(d => d.maPhieu == maPhieu).FirstOrDefault();
                    string tenLoaiNghiPhep = lqHeThong.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_LoaiNghiPhep>().Where(d => d.maLoaiNghiPhep == dsPhieuNghiPhep.loaiNghiPhep).Select(d => d.tenLoaiNghiPhep).FirstOrDefault() ?? string.Empty;
                    string loaiThoiGianNghi = dsPhieuNghiPhep.loaiThoiGianNghi == true ? "Nhiều ngày" : "Trong ngày";
                    if (dsPhieuNghiPhep.loaiThoiGianNghi == false)
                    {
                        string loaiNghi = string.Empty;
                        if (dsPhieuNghiPhep.loaiNgayBatDau == 1)
                        {
                            loaiNghi = "Cả ngày";
                        }
                        else if (dsPhieuNghiPhep.loaiNgayBatDau == 2)
                        {
                            loaiNghi = "Nửa ngày (buối sáng)";
                        }
                        else
                        {
                            loaiNghi = "Nửa ngày (buối chiều)";
                        }
                        noiDung = "Loại nghỉ phép: " + tenLoaiNghiPhep + ", Thời gian nghỉ: " + loaiThoiGianNghi + ", Ngày nghỉ: " + dsPhieuNghiPhep.ngayBatDau.Value.ToString("dd/MM/yyyy") + ", Loại nghỉ: " + loaiNghi + ", Số ngày nghỉ phép: " + dsPhieuNghiPhep.soNgayNghiPhepThucTe + ", Nội dung: " + dsPhieuNghiPhep.lyDo + "";

                    }
                    else
                    {
                        string loaiNghiBatDau = string.Empty;
                        if (dsPhieuNghiPhep.loaiNgayBatDau == 1)
                        {
                            loaiNghiBatDau = "Cả ngày";
                        }
                        else if (dsPhieuNghiPhep.loaiNgayBatDau == 2)
                        {
                            loaiNghiBatDau = "Nửa ngày (buối sáng)";
                        }
                        else
                        {
                            loaiNghiBatDau = "Nửa ngày (buối chiều)";
                        }

                        string loaiNghiKetThuc = string.Empty;
                        if (dsPhieuNghiPhep.loaiNgayKetThuc == 1)
                        {
                            loaiNghiKetThuc = "Cả ngày";
                        }
                        else if (dsPhieuNghiPhep.loaiNgayBatDau == 1)
                        {
                            loaiNghiKetThuc = "Nửa ngày (buối sáng)";
                        }
                        else
                        {
                            loaiNghiKetThuc = "Nửa ngày (buối chiều)";
                        }

                        noiDung = "Loại nghỉ phép: " + tenLoaiNghiPhep + ", Thời gian nghỉ: " + loaiThoiGianNghi + ", Ngày nghỉ bắt đầu: " + dsPhieuNghiPhep.ngayBatDau.Value.ToString("dd/MM/yyyy") + ", Ngày nghỉ kết thúc: " + dsPhieuNghiPhep.ngayKetThuc.Value.ToString("dd/MM/yyyy") + " , Loại nghỉ bắt đầu: " + loaiNghiBatDau + ", Loại nghỉ kết thúc: " + loaiNghiKetThuc + " , Số ngày nghỉ phép: " + dsPhieuNghiPhep.soNgayNghiPhepThucTe + ", Nội dung: " + dsPhieuNghiPhep.lyDo + "";
                    }
                    lichSu.maPhieu = maPhieu;
                    lichSu.maCongViec = maCongViec;
                    lichSu.ngayCapNhat = DateTime.Now;
                    lichSu.maNVCapNhat = GetUser().manv;
                    lichSu.trangThai = trangThai;
                    lichSu.maNhanVienLapPhieu = dsPhieuNghiPhep.maNhanVien;
                    lichSu.noiDung = noiDung;
                    lqNS.tbl_NS_LichSuCapNhatPhieus.InsertOnSubmit(lichSu);
                    lqNS.SubmitChanges();
                }
                else if (maCongViec == "PhieuTangCa")
                {
                    var dsPhieuTangCa = lqPhieuDN.tbl_NS_PhieuTangCas.Where(d => d.soPhieu == maPhieu).FirstOrDefault();
                    string hinhThucTangCa = dsPhieuTangCa.hinhThucTangCa == "tctl" ? "Tăng ca tính lương" : "Tăng ca nghỉ bù";
                    string tenLoaiTangCa = lqNS.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_LoaiTangCa>().Where(d => d.id == dsPhieuTangCa.loaiTangCa).Select(d => d.loaiTangCa).FirstOrDefault() ?? string.Empty;
                    var tangCa = lqPhieuDN.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_LoaiTangCa>().Where(d => d.id == dsPhieuTangCa.loaiTangCa).FirstOrDefault();
                    dsPhieuTangCa.heSoTangCa = (double?)tangCa.heSoTangCa ?? 0;
                    string maNhanVien = string.Empty;
                    var dsNhanVien = lqPhieuDN.tbl_NS_PhieuTangCa_DSNhanViens.Where(d => d.soPhieu == maPhieu).ToList();
                    if (dsNhanVien != null && dsNhanVien.Count > 0)
                    {
                        foreach (var item in dsNhanVien)
                        {
                            maNhanVien = maNhanVien + "," + item.maNhanVien;
                        }
                    }
                    noiDung = "Ngày tăng ca: " + dsPhieuTangCa.ngayTangCa.Value.ToString("dd/MM/yyyy") + ", Giờ bắt đầu: " + dsPhieuTangCa.gioBatDau.Value.ToString("HH\\:mm") + ", Giờ kết thúc " + dsPhieuTangCa.gioKetThuc.Value.ToString("HH\\:mm") + ", Số giờ tăng ca: " + dsPhieuTangCa.soGioTangCa.Value.ToString("HH\\:mm") + ", Loại tăng ca: " + tenLoaiTangCa + ", Hình thức tăng ca: " + hinhThucTangCa + ", Bắt đầu nghỉ giữa ca: " + dsPhieuTangCa.batDauNghiGiuaCa.Value.ToString("HH\\:mm") + ", Kết thúc nghỉ giữa ca: " + dsPhieuTangCa.ketThucNghiGiuaCa.Value.ToString("HH\\:mm") + ", Hế số tăng ca: " + dsPhieuTangCa.heSoTangCa + ", Thời gian nghỉ giữa ca: " + dsPhieuTangCa.thoiGianNghiGiuaCa + ", Nội dung: " + dsPhieuTangCa.noiDungTangCa + ", Danh sách nhân viên: " + maNhanVien + "";
                    lichSu.maPhieu = maPhieu;
                    lichSu.maCongViec = maCongViec;
                    lichSu.ngayCapNhat = DateTime.Now;
                    lichSu.maNVCapNhat = GetUser().manv;
                    lichSu.trangThai = trangThai;
                    lichSu.maNhanVienLapPhieu = dsPhieuTangCa.nguoiLap;
                    lichSu.noiDung = noiDung;
                    lqNS.tbl_NS_LichSuCapNhatPhieus.InsertOnSubmit(lichSu);
                    lqNS.SubmitChanges();
                }
            }
            catch
            {
            }
        }
        #endregion
    }
}