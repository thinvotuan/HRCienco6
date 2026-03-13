using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Models.DanhMuc;
using BatDongSan.Controllers;
using BatDongSan.Utils.Paging;
using BatDongSan.Helper.Common;
using BatDongSan.Models.HeThong;
using BatDongSan.Models.NhanSu;
using System.Text;
using BatDongSan.Helper.Utils;
using BatDongSan.Models.ERP;

namespace BatDongSan.Controllers.DanhMuc
{
    public class PhongBanController : ApplicationController
    {
        private LinqDanhMucDataContext linqDanhMuc = new LinqDanhMucDataContext();

        private LinqHeThongDataContext context = new LinqHeThongDataContext();
        private LinqNhanSuDataContext linqNS = new LinqNhanSuDataContext();
        private IList<tbl_DM_PhongBan> phongBans;
        private tbl_DM_PhongBan phongBan;
        private bool? permission;
        public const string taskIDSystem = "PhongBan";
        private StringBuilder buildTree = null;

        public ActionResult Index()
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return RedirectToAction("AccessDenied");
            #endregion


            try
            {
                buildTree = new StringBuilder();
                phongBans = linqDanhMuc.tbl_DM_PhongBans.ToList();
                buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
                ViewBag.departments = buildTree.ToString();
                string flag = AdminNhanSu(GetUser().manv);
                ViewBag.flagAdmin = flag;
                ViewBag.shiftType = linqDanhMuc.tbl_NS_PhanCas.Where(t => t.tenPhanCa != "").ToList();
                ViewBag.page = 0;
                ViewBag.total = 0;
                return View(phongBans);
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("Error");
            }
        }

        public ActionResult LoadNhanVien(int? id, string qSearch, int _page, string parrentId)
        {
            try
            {
                ViewData["parrentId"] = parrentId;
                string parentID = linqDanhMuc.tbl_DM_PhongBans.Where(d => d.maPhongBan == parrentId).Select(d => d.maCha).FirstOrDefault() ?? string.Empty;
                if (String.IsNullOrEmpty(parentID))
                {
                    parrentId = string.Empty;
                }
                //int page = _page == 0 ? 1 : _page;
                //int pIndex = page;
                //int total = IsNullOrEmpty.sp_PB_DanhSachNhanVien(qSearch, parrentId).Count();
                //PagingLoaderController("/Department.mvc/Index/", total, page, "?qsearch=" + qSearch + "&parrentId=" + parrentId);
                //ViewData["nhanVien"] = hr.sp_PB_DanhSachNhanVien(qSearch, parrentId).Skip(start).Take(offset).ToList();

                ViewData["qSearch"] = qSearch ?? string.Empty;
                return PartialView("LoadNhanVien");

            }
            catch (Exception e)
            {
                ViewData["Message"] = e.Message;
                return View("Error");
            }
        }

        public JsonResult GetDepartmentByID(string departmentID)
        {
            try
            {
                phongBan = new tbl_DM_PhongBan();
                var record = linqDanhMuc.tbl_DM_PhongBans.Where(d => d.maPhongBan == departmentID).FirstOrDefault();

                phongBan.maPhongBan = record.maPhongBan;
                phongBan.tenPhongBan = record.tenPhongBan;
                phongBan.maCha = record.maCha != null ? record.maCha : String.Empty;
                phongBan.ghiChu = record.ghiChu;
                phongBan.maPhanCa = record.maPhanCa;
                phongBan.maNhanVienDuyet = record.maNhanVienDuyet;

                PhongBanModel Data = new PhongBanModel();
                Data.MaPhongBan = phongBan.maPhongBan;
                Data.Ten = phongBan.tenPhongBan;
                Data.maCha = phongBan.maCha != null ? record.maCha : String.Empty;
                Data.GhiChu = phongBan.ghiChu;
                Data.maPhanCa = phongBan.maPhanCa;
                Data.maNhanVienDuyet = record.maNhanVienDuyet;
                Data.tenNhanVienDuyet = HoVaTen(record.maNhanVienDuyet);
                return Json(Data);
            }
            catch (Exception ex)
            {
                return Json(null);
            }
        }
        public string HoVaTen(string MaNV)
        {

            return linqNS.tbl_NS_NhanViens.Where(d => d.maNhanVien == MaNV).Select(d => d.ho + " " + d.ten).FirstOrDefault();
        }
        //public int capNhatCienco6() {
        //    linqERPCienco6DataContext linqERP = new linqERPCienco6DataContext();

        //    // update nvCienco6
        //    LinqNhanSuDataContext linqNS = new LinqNhanSuDataContext();
        //    var listNVNew = (from p in linqNS.vw_NS_DanhSachNhanVienTheoPhongBans
        //                     join nv in linqNS.tbl_NS_NhanViens on p.maNhanVien equals nv.maNhanVien
        //                     join cp in linqNS.GetTable<tbl_DM_CapBacChucDanh>() on p.SoCapBac equals cp.soCapBac
        //                     join q in linqNS.GetTable<Sys_User>() on p.maNhanVien equals q.manv
        //                     where nv.trangThai == 0 && q.userName != "admin"
        //                     select new
        //                     {
        //                         maNhanVien = p.maNhanVien + "C6",
        //                         userName = p.email,
        //                         matKhau = q.password,
        //                         email = p.email,
        //                         maPhongBan = p.maPhongBan,
        //                         tenPhongBan = p.tenPhongBan + "C6",
        //                         trangThai = nv.trangThai,
        //                         tenChucDanh = p.TenChucDanh + "C6",
        //                         maChucDanh = p.maChucDanh,
        //                         soDienThoai = q.telephone,
        //                         soCapBac = p.SoCapBac,
        //                         tenCapBacQL = cp.tenCapBac,
        //                         hoVaTen = nv.ho + " " + nv.ten,
        //                         ten = nv.ten,
        //                         ho = nv.ho,
        //                         gioiTinh = nv.gioiTinh,
        //                         ngaySinh = nv.ngaySinh,
        //                         noiSinh = nv.noiSinh,
        //                         tinhTrangHonNhan = nv.tinhTrangHonNhan,
        //                         phoneNumber1 = nv.phoneNumber1,
        //                     }

        //                    ).ToList();
        //    foreach (var item in listNVNew) {
        //        var checkNEw = linqERP.tbl_NhanVienHRs.Where(d => d.email == item.email).FirstOrDefault();
        //        if (checkNEw != null) {
        //            checkNEw.nvCienco6 = 1;
        //            linqERP.SubmitChanges();
        //        }
        //        var checkNEwNhanVien = linqERP.NS_TBL_NhanViens.Where(d => d.Email == item.email).FirstOrDefault();
        //        if (checkNEwNhanVien != null)
        //        {
        //            checkNEwNhanVien.nvCienco6 = 1;
        //            linqERP.SubmitChanges();
        //        }
        //        var checkUser = linqERP.HT_Users.Where(d => d.email == item.email).FirstOrDefault();
        //        if (checkNEwNhanVien != null)
        //        {
                    
        //            linqERP.HT_Users.DeleteOnSubmit(checkUser);
        //            linqERP.SubmitChanges();
        //        }
        //    }
        //    return 1;
        //}

        public int capNhatPasswodHRtoERP() {
            linqERPCienco6DataContext linqERP = new linqERPCienco6DataContext();
            LinqNhanSuDataContext linqNS = new LinqNhanSuDataContext();
            var listNVNew = (from p in linqNS.vw_NS_DanhSachNhanVienTheoPhongBans
                             join nv in linqNS.tbl_NS_NhanViens on p.maNhanVien equals nv.maNhanVien
                             join cp in linqNS.GetTable<tbl_DM_CapBacChucDanh>() on p.SoCapBac equals cp.soCapBac
                             join q in linqNS.GetTable<Sys_User>() on p.maNhanVien equals q.manv
                             where nv.trangThai == 0 && q.userName != "admin"
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
            foreach (var item in listNVNew)
            {



                // insert new

              
                var tblNhanVienERP = linqERP.tbl_NhanVienHRs.Where(d => d.email == item.email).FirstOrDefault();
                if (tblNhanVienERP != null)
                {
                    tblNhanVienERP.matKhau = item.matKhau;
                    linqERP.SubmitChanges();
                }
                var sysUser = linqERP.HT_Users.Where(d => d.email == item.email).FirstOrDefault();
                if (sysUser != null)
                {
                    sysUser.password = item.matKhau;
                    linqERP.SubmitChanges();
                }

            }
            return 1;
        }
        public int ImportHRToDMS() {
            if ((string)Session["CongTy"] == "Cienco6")
            {

                BatDongSan.Models.DMS.LinqDMSDataContextDataContext DMSContext = new BatDongSan.Models.DMS.LinqDMSDataContextDataContext();
                var listHR = DMSContext.tbl_DMS_NhanVienHRs.Where(d => d.nvDMS == null || d.nvDMS == 0).ToList();
                DMSContext.tbl_DMS_NhanVienHRs.DeleteAllOnSubmit(listHR);
                // Insert new
                var listNVNew = (from p in linqNS.vw_NS_DanhSachNhanVienTheoPhongBans
                                 join nv in linqNS.tbl_NS_NhanViens on p.maNhanVien equals nv.maNhanVien
                                 join cp in linqNS.GetTable<tbl_DM_CapBacChucDanh>() on p.SoCapBac equals cp.soCapBac
                                 join q in linqNS.GetTable<Sys_User>() on p.maNhanVien equals q.manv
                                 where nv.trangThai == 0
                                 select new NhanVienModel
                                 {
                                     maNhanVien = p.maNhanVien,
                                     userName = q.userName,
                                     matKhau = q.password,
                                     email = p.email,
                                     maPhongBan = p.maPhongBan,
                                     tenPhongBan = p.tenPhongBan,
                                     trangThai = nv.trangThai,
                                     tenChucDanh = p.TenChucDanh,
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
                                     choPhepVaoERPKhac = true
                                 }

                                 ).ToList();

                // Add list

                BatDongSan.Models.DMS.tbl_DMS_NhanVienHR tblSysMaC = null;
                List<BatDongSan.Models.DMS.tbl_DMS_NhanVienHR> lst_tblNhanVienBangCap = new List<BatDongSan.Models.DMS.tbl_DMS_NhanVienHR>();




                if (listNVNew != null && listNVNew.Count > 0)
                {
                    foreach (var item in listNVNew)
                    {

                        tblSysMaC = new BatDongSan.Models.DMS.tbl_DMS_NhanVienHR();
                        tblSysMaC.maNhanVien = item.maNhanVien;
                        tblSysMaC.maCongTrinh = "TP00";
                        tblSysMaC.userName = item.userName;
                        tblSysMaC.matKhau = item.matKhau;
                        tblSysMaC.email = item.email;
                        tblSysMaC.trangThai = item.trangThai;
                        tblSysMaC.tenChucDanh = item.tenChucDanh;
                        tblSysMaC.maChucDanh = item.maChucDanh;
                        tblSysMaC.soDienThoai = item.soDienThoai;
                        tblSysMaC.soCapBac = item.soCapBac;
                        tblSysMaC.tenCapBac = item.tenCapBacQL;
                        tblSysMaC.maPhongBan = item.maPhongBan;
                        tblSysMaC.hoTen = item.hoVaTen;
                        tblSysMaC.dungLuong = 10;
                        tblSysMaC.choPhepVaoERPKhac = item.choPhepVaoERPKhac;
                        lst_tblNhanVienBangCap.Add(tblSysMaC);


                    }
                    DMSContext.tbl_DMS_NhanVienHRs.InsertAllOnSubmit(lst_tblNhanVienBangCap);
                    DMSContext.SubmitChanges();

                    SaveActiveHistory("Import dữ liệu nhân sự đến DMS");
                }
                // Add phong ban den DMS
                var listPhongBanThuanViet = linqDanhMuc.tbl_DM_PhongBans.ToList();
                foreach (var item in listPhongBanThuanViet)
                {
                    var checkPhongBanDMS = DMSContext.tbl_DMS_PhongBans.Where(d => d.maPhongBan == item.maPhongBan).FirstOrDefault();
                    if (checkPhongBanDMS == null)
                    {
                        BatDongSan.Models.DMS.tbl_DMS_PhongBan tblDMSPhongBan = new BatDongSan.Models.DMS.tbl_DMS_PhongBan();
                        tblDMSPhongBan.maPhongBan = item.maPhongBan;
                        tblDMSPhongBan.tenPhongBan = item.tenPhongBan;
                        tblDMSPhongBan.maCha = item.maCha;
                        DMSContext.tbl_DMS_PhongBans.InsertOnSubmit(tblDMSPhongBan);
                        DMSContext.SubmitChanges();
                    }
                }
            }
            // End add phong ban
            // End add list
            return 1;
        }
        public int ImportHRToERP() {

            if ((string)Session["CongTy"] == "Cienco6")
            {

                linqERPCienco6DataContext linqERP = new linqERPCienco6DataContext();

                // update nvCienco6
                LinqNhanSuDataContext linqNS = new LinqNhanSuDataContext();
                var listNVNew = (from p in linqNS.vw_NS_DanhSachNhanVienTheoPhongBans
                                 join nv in linqNS.tbl_NS_NhanViens on p.maNhanVien equals nv.maNhanVien
                                 join cp in linqNS.GetTable<tbl_DM_CapBacChucDanh>() on p.SoCapBac equals cp.soCapBac
                                 join q in linqNS.GetTable<Sys_User>() on p.maNhanVien equals q.manv
                                 where nv.trangThai == 0 && q.userName != "admin"
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
                var tblDelet_NhanVien = linqERP.tbl_NhanVienHRs.Where(d => d.nvCienco6 == 1).ToList();
                if (tblDelet_NhanVien != null)
                {
                    linqERP.tbl_NhanVienHRs.DeleteAllOnSubmit(tblDelet_NhanVien);
                }
                var tblDelet_TBL_NhanViens = linqERP.NS_TBL_NhanViens.Where(d => d.nvCienco6 == 1).ToList();
                if (tblDelet_TBL_NhanViens != null)
                {
                    linqERP.NS_TBL_NhanViens.DeleteAllOnSubmit(tblDelet_TBL_NhanViens);
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
                    else
                    {
                        checkHTUser.password = item.matKhau;
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
        public JsonResult GetDepartmentByIDAnHien(string departmentID)
        {
            string hasValue = string.Empty;
            try
            {

                var check = linqDanhMuc.tbl_DM_PhongBans.Where(d => d.maCha != null && d.maCha != string.Empty && d.maPhongBan == departmentID).FirstOrDefault();
                if (check != null)
                {
                    hasValue = "1";
                }
                else
                {
                    hasValue = "0";
                }
                return Json(hasValue);
            }
            catch (Exception ex)
            {
                return Json(hasValue);
            }
        }
        public JsonResult LuuPhongBan(string departmentID, string departmentName, string parentId, string note, string shiftId, bool flag, string maNguoiDuyet)
        {

            try
            {
                phongBan = new tbl_DM_PhongBan();
                phongBan.maPhongBan = departmentID;
                phongBan.maCha = parentId == "" ? null : parentId;
                phongBan.tenPhongBan = departmentName;
                phongBan.nguoiLap = GetUser().userName;
                phongBan.ngayLap = DateTime.Now;
                phongBan.ghiChu = note;
                phongBan.maNhanVienDuyet = maNguoiDuyet;
                try
                {
                    phongBan.maPhanCa = (int)Convert.ToInt64(shiftId);
                }
                catch { }
                int data = 0;
                if (flag == true)
                {
                    if (CheckExistDepartment(departmentID))
                    {
                        data = 0;
                    }
                    else if (CheckExistDepartmentName(departmentName))
                    {
                        data = 1;
                    }
                    else
                    {
                        Insert(phongBan);
                        data = 2;
                    }
                }
                else
                {
                    phongBan.ngayCapNhat = DateTime.Now;
                    Update(phongBan);
                    data = 2;
                }

                return Json(data);
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return Json(null);
            }
        }

        public void Insert(tbl_DM_PhongBan phongBan)
        {
            linqDanhMuc.tbl_DM_PhongBans.InsertOnSubmit(phongBan);
            linqDanhMuc.SubmitChanges();
        }

        public void Update(tbl_DM_PhongBan department)
        {
            var departmentLinq = from t in linqDanhMuc.tbl_DM_PhongBans
                                 where t.maPhongBan.Equals(department.maPhongBan)
                                 select t;
            tbl_DM_PhongBan phongban = departmentLinq.SingleOrDefault();
            phongban.tenPhongBan = department.tenPhongBan;
            phongban.nguoiLap = department.nguoiLap;
            phongban.ngayCapNhat = DateTime.Now;
            phongban.ghiChu = department.ghiChu;
            phongban.maPhanCa = department.maPhanCa;
            phongban.maCha = department.maCha;
            phongban.maNhanVienDuyet = department.maNhanVienDuyet;
            linqDanhMuc.SubmitChanges();
        }

        public bool CheckExistDepartment(string maPhongBan)
        {
            return linqDanhMuc.tbl_DM_PhongBans.Where(d => d.maPhongBan == maPhongBan).Count() > 0 ? true : false;
        }

        public bool CheckExistDepartmentName(string tenPhong)
        {
            return linqDanhMuc.tbl_DM_PhongBans.Where(d => d.tenPhongBan == tenPhong).Count() > 0 ? true : false;
        }

        public ActionResult CapNhatPhongBanCha(int? page, string searchString)
        {
            ViewBag.isGet = "True";
            IList<sp_DM_DanhSachBoPhanResult> phongBans;
            phongBans = linqDanhMuc.sp_DM_DanhSachBoPhan(searchString).ToList();
            int currentPageIndex = page.HasValue ? page.Value : 1;
            ViewBag.Count = phongBans.Count();
            return PartialView("_DanhSachPhongBan", phongBans.ToPagedList(currentPageIndex, 20));
        }

        public ActionResult XoaPhongBan(string departmentID)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXoa);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            try
            {
                RemovePhongBan(departmentID);
            }
            catch (Exception e)
            {
                
                ViewData["Message"] = e.Message;
                return View("Error");
            }
            return Json(String.Empty);
        }

        public void RemovePhongBan(string departmentID)
        {            
            DeletePhongBanCon(departmentID);
            var departmentLinq = from t in linqDanhMuc.tbl_DM_PhongBans
                                 where t.maPhongBan.Equals(departmentID)
                                 select t;
            linqDanhMuc.tbl_DM_PhongBans.DeleteOnSubmit(departmentLinq.SingleOrDefault());
            linqDanhMuc.SubmitChanges();
        }

        private void DeletePhongBanCon(string departmentID)
        {
            var childNode = from t in linqDanhMuc.tbl_DM_PhongBans
                            where t.maCha.Equals(departmentID)
                            select t.maPhongBan;
            if (childNode.Count() > 0)
            {
                foreach (var item in childNode)
                {
                    var sqlDeleteChild = linqDanhMuc.tbl_DM_PhongBans
                        .Where(d => d.maPhongBan.Equals(item));
                    if (sqlDeleteChild.Count() > 0)
                    {
                        linqDanhMuc.tbl_DM_PhongBans.DeleteAllOnSubmit(sqlDeleteChild);
                        //Recursive
                        DeletePhongBanCon(item);
                    }
                }
            }
        }

        public ActionResult NhanVienDuyet()
        {
            buildTree = new StringBuilder();
            phongBans = linqDanhMuc.tbl_DM_PhongBans.ToList();
            buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
            ViewBag.NVPB = buildTree.ToString();
            return PartialView("_NhanVienPhongBan");
        }
        public ActionResult DanhSachPhongBan()
        {
            buildTree = new StringBuilder();
            phongBans = linqDanhMuc.tbl_DM_PhongBans.ToList();
            buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
            ViewBag.NVPB = buildTree.ToString();
            return PartialView("_PartDanhSachPhongBan");
        }

        public ActionResult LoadNhanVienDuyet(int? page, string searchString, string maPhongBan)
        {
            IList<sp_PB_DanhSachNhanVienResult> phongBan1s;
            phongBan1s = linqDanhMuc.sp_PB_DanhSachNhanVien(searchString, maPhongBan).ToList();
            ViewBag.isGet = "True";
            int currentPageIndex = page.HasValue ? page.Value : 1;
            ViewBag.Count = phongBan1s.Count();
            ViewBag.Search = searchString;
            ViewBag.MaPhongBan = maPhongBan;
            return PartialView("LoadNhanVienDuyet", phongBan1s.ToPagedList(currentPageIndex, 10));
        }

        public ActionResult LoadNhanVienPhongBan(int? page, string searchString, string maPhongBan)
        {
            try
            {
                string parentID = linqDanhMuc.tbl_DM_PhongBans.Where(d => d.maPhongBan == maPhongBan).Select(d => d.maCha).FirstOrDefault() ?? string.Empty;
                if (String.IsNullOrEmpty(parentID))
                {
                    maPhongBan = string.Empty;
                }
                IList<sp_PB_DanhSachNhanVienResult> phongBan1s;
                phongBan1s = linqDanhMuc.sp_PB_DanhSachNhanVien(searchString, maPhongBan).ToList();
                ViewBag.isGet = "True";
                int currentPageIndex = page.HasValue ? page.Value : 1;
                ViewBag.Count = phongBan1s.Count();
                ViewBag.parrentId = maPhongBan;
                ViewBag.qSearchNV = searchString ?? string.Empty;
                return PartialView("LoadNhanVien",phongBan1s.ToPagedList(currentPageIndex, 10));
            }
            catch (Exception e)
            {
                ViewData["Message"] = e.Message;
                return View("Error");
            }
        }
        public ActionResult SoDoToChuc()
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return RedirectToAction("AccessDenied");
            #endregion


            try
            {
                buildTree = new StringBuilder();
                phongBans = linqDanhMuc.tbl_DM_PhongBans.ToList();
                buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
                ViewBag.departments = buildTree.ToString();

                ViewBag.shiftType = linqDanhMuc.tbl_NS_PhanCas.Where(t => t.tenPhanCa != "").ToList();
                ViewBag.page = 0;
                ViewBag.total = 0;
                return View(phongBans);
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("Error");
            }
        }
        public ActionResult LoadTreeNhanVien(string parrentId, string qSearch, int? soCapBac, int pageIndex, int pageSize)
        {


            try
            {

                var list = context.sp_PB_DanhSachNhanVien_Tree(qSearch, parrentId, soCapBac, pageIndex, pageSize).ToList();
                int totalList = context.sp_PB_DanhSachNhanVien_Tree(qSearch, parrentId, soCapBac, 0, 10000).Count();


                if (list.Count == 0)
                {
                    list = null;
                }
                
                return Json(new { count = totalList, lists = list });
               
            }
            catch
            {
                return Json(string.Empty);
            }

        }

        //public ActionResult LoadNhanVienLienKetView(int? page, string parrentId)
        //{
        //    try
        //    {
                
        //        string parentID = linqDanhMuc.tbl_DM_PhongBans.Where(d => d.maPhongBan == parrentId).Select(d => d.maCha).FirstOrDefault() ?? string.Empty;
        //        if (String.IsNullOrEmpty(parentID))
        //        {
        //            parrentId = string.Empty;
        //        }
        //        IList<sp_PB_DanhSachNhanVienLienKetResult> phongBan1s;
        //        phongBan1s = linqDanhMuc.sp_PB_DanhSachNhanVienLienKet(parrentId).ToList();
        //        ViewBag.isGet = "True";
        //        int currentPageIndex = page.HasValue ? page.Value : 1;
        //        ViewBag.Count = phongBan1s.Count();
        //        ViewBag.parrentId = parrentId;
        //        return PartialView("LoadNhanVienLienKetView",phongBan1s.ToPagedList(currentPageIndex, 10));
        //    }
        //    catch (Exception e)
        //    {
        //        ViewData["Message"] = e.Message;
        //        return View("Error");
        //    }


        //}
        public string removeNVLienKetPhongBan(string MaNV, string MaPB)
        {
            try
            {
                var list = linqNS.tbl_NS_NhanVienLienKetPhongBans.Where(d => d.employeeId == MaNV && d.departmentId == MaPB);
                linqNS.tbl_NS_NhanVienLienKetPhongBans.DeleteAllOnSubmit(list);
                linqNS.SubmitChanges();
                return "true";
            }
            catch (Exception e)
            {
                return "Lỗi: " + e.Message;
            }
        }
        public ActionResult ChonNhanVienLienKet()
        {
            buildTree = new StringBuilder();
            phongBans = linqDanhMuc.tbl_DM_PhongBans.ToList();
            buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
            ViewBag.NVPB = buildTree.ToString();
            return PartialView("_NhanVienLienKet");
        }
        public ActionResult LoadNhanVienLienKet(int? page, string searchString, string maPhongBan)
        {
            IList<sp_PB_DanhSachNhanVienResult> phongBan1s;
            phongBan1s = linqDanhMuc.sp_PB_DanhSachNhanVien(searchString, maPhongBan).ToList();
            ViewBag.isGet = "True";
            int currentPageIndex = page.HasValue ? page.Value : 1;
            ViewBag.Count = phongBan1s.Count();
            ViewBag.Search = searchString;
            ViewBag.MaPhongBan = maPhongBan;
            return PartialView("LoadNhanVienLienKet", phongBan1s.ToPagedList(currentPageIndex, 10));
        }
        public string updateListNVLienKetPhongBan(string MaNV, string MaPB)
        {

            try
            {
                // Update
                string[] MaNVs = MaNV.Split(',');
                List<tbl_NS_NhanVienLienKetPhongBan> lst_HR_EmployeeDepartment = new List<tbl_NS_NhanVienLienKetPhongBan>();
                tbl_NS_NhanVienLienKetPhongBan tblEmployeeDepart = null;
                if (MaNV != null)
                {
                    for (int i = 0; i < MaNVs.Length; i++)
                    {
                        if (linqNS.tbl_NS_NhanVienLienKetPhongBans.Where(d => d.employeeId.Contains(MaNVs[i])).FirstOrDefault() == null)
                        {
                            tblEmployeeDepart = new tbl_NS_NhanVienLienKetPhongBan();
                            tblEmployeeDepart.departmentId = MaPB;
                            tblEmployeeDepart.employeeId = MaNVs[i];
                            tblEmployeeDepart.userName = GetUser().userName;
                            tblEmployeeDepart.updateDate = DateTime.Now;
                            lst_HR_EmployeeDepartment.Add(tblEmployeeDepart);
                        }

                    }
                    linqNS.tbl_NS_NhanVienLienKetPhongBans.InsertAllOnSubmit(lst_HR_EmployeeDepartment);
                }

                linqNS.SubmitChanges();
                return "true";


            }
            catch (Exception e)
            {
                return "Lỗi: " + e.Message;
            }
        }

        public ActionResult chuyenNV(string arrMaNV)
        {
            buildTree = new StringBuilder();
            phongBans = linqDanhMuc.tbl_DM_PhongBans.ToList();
            buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
            ViewBag.PhongBanChuyen = buildTree.ToString();
            ViewBag.NhanVienChuyen = arrMaNV;
            return PartialView("_PhongBanChuyen");           
        }
        public string updateListNVPhongBan(string MaNV, string MaPB)
        {

            try
            {
                // Update
                string[] MaNVs = MaNV.Split(',');
                List<tbl_NS_NhanVienPhongBan> lst_HR_EmployeeDepartment = new List<tbl_NS_NhanVienPhongBan>();
                tbl_NS_NhanVienPhongBan tblEmployeeDepart = null;
                if (MaNV != null)
                {
                    for (int i = 0; i < MaNVs.Length; i++)
                    {

                        tblEmployeeDepart = new tbl_NS_NhanVienPhongBan();
                        tblEmployeeDepart.maPhongBan = MaPB;
                        tblEmployeeDepart.maNhanVien = MaNVs[i];
                        tblEmployeeDepart.nguoiLap = GetUser().userName;
                        tblEmployeeDepart.ngayLap = DateTime.Now;
                        lst_HR_EmployeeDepartment.Add(tblEmployeeDepart);

                    }
                    linqNS.tbl_NS_NhanVienPhongBans.InsertAllOnSubmit(lst_HR_EmployeeDepartment);
                }

                linqNS.SubmitChanges();
                return "true";


            }
            catch (Exception e)
            {
                return "Lỗi: " + e.Message;
            }
        }

        public ActionResult ListDSNhanVien(int? page, string qSearch, string parrentId)
        {
            try
            {
                
                var ParentGoc = linqDanhMuc.tbl_DM_PhongBans.Where(d => d.maCha == null).FirstOrDefault();
                var record = linqDanhMuc.sp_PB_DanhSachNhanVien(qSearch, string.Empty).Where(d=>d.departmentId!=parrentId).ToList();
                ViewBag.isGet = "True";
                int currentPageIndex = page.HasValue ? page.Value : 1;
                ViewBag.Count = record.Count();
                ViewBag.qSearchNVAdd = qSearch;
                ViewBag.parrentIdAdd = parrentId;
                return PartialView("_ListDSNhanVien", record.ToPagedList(currentPageIndex, 20));     
            }
            catch (Exception e)
            {
                ViewData["Message"] = e.Message;
                return View("Error");
            }
        }

        public string updateNVPhongBan(string MaNV, string MaPB)
        {
            try
            {
                tbl_NS_NhanVienPhongBan tblEmployeeDepart = new tbl_NS_NhanVienPhongBan();
                tblEmployeeDepart.maPhongBan = MaPB;
                tblEmployeeDepart.maNhanVien = MaNV;
                tblEmployeeDepart.nguoiLap = GetUser().userName;
                tblEmployeeDepart.ngayLap = DateTime.Now;

                linqNS.tbl_NS_NhanVienPhongBans.InsertOnSubmit(tblEmployeeDepart);
                linqNS.SubmitChanges();
                return "true";
            }
            catch (Exception e)
            {
                return "Lỗi: " + e.Message;
            }
        }
        // Load nhan vien lien ket.
        public ActionResult LoadNhanVienLienKetView(string parrentId)
        {
            try
            {
                string parentID = linqDanhMuc.tbl_DM_PhongBans.Where(d => d.maPhongBan == parrentId).Select(d => d.maCha).FirstOrDefault() ?? string.Empty;
                if (String.IsNullOrEmpty(parentID))
                {
                    parrentId = string.Empty;
                }
                int total = linqDanhMuc.sp_PB_DanhSachNhanVienLienKet(parrentId).Count();
               
                ViewData["nhanVien"] = linqDanhMuc.sp_PB_DanhSachNhanVienLienKet(parrentId).ToList();
                ViewData["total"] = total;

                return PartialView("LoadNhanVienLienKetView");
            }
            catch (Exception e)
            {
                ViewData["Message"] = e.Message;
                return View("Error");
            }


        }
        // End
    }
}
