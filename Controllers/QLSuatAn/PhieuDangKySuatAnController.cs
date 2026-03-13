using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Models.DanhMuc;
using BatDongSan.Models.QLSuatAn;
using BatDongSan.Helper.Utils;
using System.Globalization;
using System.IO;
using System.Text;
using BatDongSan.Models.HeThong;
using BatDongSan.Models.NhanSu;
using BatDongSan.Models.QLSuatAn.LuuVet;
using BatDongSan.Utils.Paging;
//using Xceed.Words.NET;

namespace BatDongSan.Controllers.QLSuatAn
{
    public class PhieuDangKySuatAnController : ApplicationController
    {
        private QLSuatAnDataContext context = new QLSuatAnDataContext();
        private IList<tbl_SA_CuaHang> cuaHangs;
        private tbl_SA_CuaHang cuaHang;
        private readonly string MCV = "PhieuDangKySuatAn";
        private bool? permission;
        private string mimeType;
        private string duongDan;
        private StringBuilder buildTree;
        private IList<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan> phongBans;
        public ActionResult Index(int? id, int? _page, string searchString, string maCuaHang)
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
                QLSALuuVetDataContext linqSALuuVet = new QLSALuuVetDataContext();
                var lstCuaHang = linqSALuuVet.tbl_SA_CuaHangs.ToList();
                ViewBag.danhSachCuaHang = lstCuaHang;
                lstCuaHang.Insert(0, new tbl_SA_CuaHang { maCuaHang = "", tenCuaHang = "[ Tất cả]" });


                ViewBag.lstCuaHangs = new SelectList(lstCuaHang, "maCuaHang", "tenCuaHang", maCuaHang);

                string param = "?maCuaHang=" + maCuaHang + "&searchString=" + searchString;
                int? page = _page == 0 ? 1 : _page;
                int? pIndex = page;
                var lstCuaHangMonAn = (from p in context.tbl_SA_PhieuAns
                                       //join nv in context.GetTable<tbl_NS_NhanVien>() on p.maNhanVien equals nv.maNhanVien
                                       //join q in context.GetTable<tbl_SA_CuaHang>() on p.maCuaHang equals q.maCuaHang
                                       where (p.maNhanVien == GetUser().manv || p.nguoiLapPhieu == GetUser().manv)
                                       && (maCuaHang == p.maCuaHang || maCuaHang == "" || maCuaHang == null)
                                       && p.maMonAn != null
                                       select new PhieuDangKySuatAnModel
                                       {
                                           maCuaHang = p.maCuaHang,
                                           //tenCuaHang = q.tenCuaHang,
                                           maPhieu = p.maPhieu,
                                           maNhanVien = p.maNhanVien,
                                           hoTenNhanVien = HoVaTen(p.maNhanVien),
                                           //hinhAnhQuan = q.avatar,
                                           ngayLapPhieu = p.ngayLapPhieu

                                       }
                                           ).OrderByDescending(d => d.ngayLapPhieu).ToList();
                int total = lstCuaHangMonAn.Count();
                PagingLoaderController("/PhieuDangKySuatAn/Index/", total, page ?? 1, param);
                var baoCaos = lstCuaHangMonAn.ToList();
                ViewBag.List = baoCaos.Skip(start).Take(offset).ToList();

                thang(DateTime.Now.Month);
                nam(DateTime.Now.Year);

                return View();
            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.Message;
                return View("Error");
            }
        }
        public ActionResult ViewIndex(int? id, int? _page, string searchString, string maCuaHang)
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
                QLSALuuVetDataContext linqSALuuVet = new QLSALuuVetDataContext();
                var lstCuaHang = linqSALuuVet.tbl_SA_CuaHangs.ToList();
                ViewBag.danhSachCuaHang = lstCuaHang;
                lstCuaHang.Insert(0, new tbl_SA_CuaHang { maCuaHang = "", tenCuaHang = "[ Tất cả]" });


                ViewBag.lstCuaHangs = new SelectList(lstCuaHang, "maCuaHang", "tenCuaHang", maCuaHang);

                string param = "?maCuaHang=" + maCuaHang + "&searchString=" + searchString;
                var lstCuaHangMonAn = (from p in context.tbl_SA_PhieuAns
                                       //join nv in context.GetTable<tbl_NS_NhanVien>() on p.maNhanVien equals nv.maNhanVien
                                       //join q in context.GetTable<tbl_SA_CuaHang>() on p.maCuaHang equals q.maCuaHang
                                       where (p.maNhanVien == GetUser().manv || p.nguoiLapPhieu == GetUser().manv)
                                       && (maCuaHang == p.maCuaHang || maCuaHang == "" || maCuaHang == null)
                                       && p.maMonAn != null
                                       select new PhieuDangKySuatAnModel
                                       {
                                           maCuaHang = p.maCuaHang,
                                           //tenCuaHang = q.tenCuaHang,
                                           maPhieu = p.maPhieu,
                                           maNhanVien = p.maNhanVien,
                                           hoTenNhanVien = HoVaTen(p.maNhanVien),
                                           //hinhAnhQuan = q.avatar,
                                           ngayLapPhieu = p.ngayLapPhieu

                                       }
                                           ).OrderByDescending(d => d.ngayLapPhieu).ToList();
                int? page = _page == 0 ? 1 : _page;
                int? pIndex = page;
                int total = lstCuaHangMonAn.Count();
                PagingLoaderController("/PhieuDangKySuatAn/Index/", total, page ?? 1, param);
                var baoCaos = lstCuaHangMonAn.ToList();

                ViewBag.List = baoCaos.Skip(start).Take(offset).ToList();
                ViewBag.Ajax = "true";
                return PartialView("PartialContent");
            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.Message;
                return View("Error");
            }
        }

        //
        // GET: /BangCap/Details/5

        public ActionResult Details(string id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            QLSALuuVetDataContext linqSALuuVet = new QLSALuuVetDataContext();
            PhieuDangKySuatAnModel phieuDangKySA = new PhieuDangKySuatAnModel();
            var thongTinPhieu = (from p in context.tbl_SA_PhieuAns
                                 where p.maPhieu == id && (p.maNhanVien == GetUser().manv || p.nguoiLapPhieu == GetUser().manv)
                                 select new PhieuDangKySuatAnModel
                                 {
                                     maPhieu = p.maPhieu,
                                     maCuaHang = p.maCuaHang,
                                     maNhanVienThuHuong = p.maNhanVien,
                                     ngayLapPhieu = p.ngayLapPhieu,
                                     hoTenNhanVienThuHuong = HoVaTen(p.maNhanVien),
                                     maMonAn = p.maMonAn,
                                     maVach = p.maVach,
                                     soThuTuPhucVuTaiQuan = p.soThuTuPhucVuTaiQuan,
                                     soSuatAnConLaiCuaNVTaiQuan = p.soSuatAnConLaiCuaNVTaiQuan,
                                     thoiGianDuKienPhucVu = p.thoiGianDuKienPhucVu
                                 }

                                     ).FirstOrDefault();
            ViewBag.thonTinCuaHang = linqSALuuVet.tbl_SA_CuaHangs.Where(d => d.maCuaHang == thongTinPhieu.maCuaHang).FirstOrDefault();
            ViewBag.thongTinMonAn = linqSALuuVet.tbl_SA_CuaHang_MonAns.Where(d => d.maMonAn == thongTinPhieu.maMonAn).FirstOrDefault();
            return PartialView("Details", thongTinPhieu);
        }

        //
        // GET: /BangCap/Create

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            int timeDongSuatAn = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["TimeDongSuatAn"]);
            int phutDongSuatAn = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["PhutDongSuatAn"]);

            PhieuDangKySuatAnModel phieuDangKySA = new PhieuDangKySuatAnModel();
            QLSALuuVetDataContext linqSALuuVet = new QLSALuuVetDataContext();
            phieuDangKySA.hoTenNhanVien = HoVaTen(GetUser().manv);
            phieuDangKySA.ngayLapPhieu = DateTime.Now;
            phieuDangKySA.maPhieu = GenerateUtil.CheckLetterPA("PSA", GetMax());
            phieuDangKySA.maNhanVien = GetUser().manv;
            phieuDangKySA.hoTenNhanVienThuHuong = HoVaTen(GetUser().manv);
            phieuDangKySA.maNhanVienThuHuong = GetUser().manv;
            Boolean checkQuyen = CheckQuyetInPhieuSA(GetUser().userId, "InPhieuSA");
            Boolean checkQuyenDangKySAGroup = CheckQuyetInPhieuSA(GetUser().userId, "DangKySAGroup");
            phieuDangKySA.themNhanVien = checkQuyen;
            var checkPhieu = context.tbl_SA_PhieuAns.Where(d => d.maNhanVien == GetUser().manv && d.ngayLapPhieu.Value.Day == DateTime.Now.Day && d.ngayLapPhieu.Value.Month == DateTime.Now.Month && d.ngayLapPhieu.Value.Year == DateTime.Now.Year).FirstOrDefault();
            if (checkPhieu != null && checkQuyen == false)
            {
                ViewBag.thongTinCuaHang = linqSALuuVet.tbl_SA_CuaHangs.Where(d => d.maCuaHang == checkPhieu.maCuaHang).FirstOrDefault();
                ViewBag.errror = "Ngày hôm nay bạn đã lập phiếu.";
                return View("ErrorSuatAn");
            }
            if (GetUser().manv != "NV-004")
            {
                var timeFrom = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 8, 30, 00);
                var timeTo = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, timeDongSuatAn, phutDongSuatAn, 00);
                // Check 1330 
                var maPhongBanHienTai = context.GetTable<tbl_NS_NhanVienPhongBan>().Where(d => d.maNhanVien == GetUser().manv)
                    .OrderByDescending(d => d.id).Select(d => d.maPhongBan).FirstOrDefault();
                var maPhongBanCha = context.GetTable<tbl_DM_PhongBan>().Where(d => d.maPhongBan == maPhongBanHienTai).Select(d => d.maCha).FirstOrDefault();
                if (maPhongBanHienTai == "ZTP56" || maPhongBanCha == "ZTP56")
                {
                    timeFrom = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 6, 30, 00);
                }

                if (DateTime.Now < timeFrom)
                {
                    ViewBag.errror = "Chưa đến thời gian lập phiếu. Thời gian là từ: " + timeFrom.ToString() + " đến: " + timeTo.ToString();
                    return View("ErrorSuatAn");
                }

                if ((DateTime.Now < timeFrom || DateTime.Now > timeTo))// && checkQuyen == false)
                {
                    ViewBag.errror = "Đã hết thời gian lập phiếu.  Thời gian là từ: " + timeFrom.ToString() + " đến: " + timeTo.ToString();
                    return View("ErrorSuatAn");
                }
            }
            return PartialView("Create", phieuDangKySA);
        }
        public ActionResult ChonNhanVienSuatAnTheoPB()
        {
            LinqDanhMucDataContext linqDanhMuc = new LinqDanhMucDataContext();
            buildTree = new StringBuilder();
            List<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan> phongBans = linqDanhMuc.tbl_DM_PhongBans.ToList();
            buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
            ViewBag.NVPB = buildTree.ToString();
            return PartialView("_NhanVienSuatAnPB");
        }
        public ActionResult LoadNhanVienSuatAnTheoPB(int? page, string searchString, string maPhongBan)
        {
            LinqDanhMucDataContext linqDanhMuc = new LinqDanhMucDataContext();
            IList<sp_PB_DanhSachNhanVienResult> phongBan1s;
            phongBan1s = linqDanhMuc.sp_PB_DanhSachNhanVien(searchString, maPhongBan).ToList();
            ViewBag.isGet = "True";
            int currentPageIndex = page.HasValue ? page.Value : 1;
            ViewBag.Count = phongBan1s.Count();
            ViewBag.Search = searchString;
            ViewBag.MaPhongBan = maPhongBan;
            return PartialView("LoadNhanVienSuatAnTheoPB", phongBan1s.ToPagedList(currentPageIndex, 10));
        }

        public ActionResult LoadDanhSachCuaHang(string id)
        {
            Boolean checkQuyen = CheckQuyetInPhieuSA(GetUser().userId, "InPhieuSA");
            if (checkQuyen == false)
            {
                id = GetUser().manv;
            }
            Boolean checkQuyenCQL = CheckQuyetInPhieuSA(GetUser().userId, "CapQuanLySA");
            QLSALuuVetDataContext linqSALuuVet = new QLSALuuVetDataContext();
            List<sp_SA_PhieuDangKySuatAn_CuaHang_LuuVet_PhanBoResult> listLuuVet = new List<sp_SA_PhieuDangKySuatAn_CuaHang_LuuVet_PhanBoResult>();
            listLuuVet = linqSALuuVet.sp_SA_PhieuDangKySuatAn_CuaHang_LuuVet_PhanBo(id, System.Configuration.ConfigurationManager.AppSettings["CongTy"].ToString()).Where(d => d.maCuaHang != "CH181106-001").ToList();

            //var data = linqSALuuVet.tbl_SA_CuaHangs.ToList();

            //var lstSuatAnCh = (from p in linqSALuuVet.tbl_SA_CuaHangs.ToList()
            //                   join q in listLuuVet on p.maCuaHang equals q.maCuaHang into cuaHangLuuVet
            //                   from luuVet in cuaHangLuuVet.DefaultIfEmpty()
            //                   //where (checkQuyenCQL == true && p.capQuanLy == true) || (checkQuyenCQL == false && p.capQuanLy == false)
            //                   select new PhieuDangKySuatAnModel
            //                   {
            //                       maCuaHang = p.maCuaHang,
            //                       tenCuaHang = p.tenCuaHang,
            //                       soSuatAnTrongNgayDaDangKyCH = (luuVet == null ? 0 : luuVet.soSuatAnTrongNgayDaDangKyCH),
            //                       soLanAnCuaNhanVienTrongThang = (luuVet == null ? 0 : luuVet.soLanAnCuaNhanVienTrongThang),
            //                       soSuatAnTrongNgayCuaCuaHang = p.soSuatAnTrongNgay,
            //                       suatAnTrenNhanVien = p.suatAnTrenNhanVien,
            //                       soSuatAnConLaiTrongNgayCH = ((p.soSuatAnTrongNgay ?? 0) - (luuVet == null ? 0 : luuVet.soSuatAnTrongNgayDaDangKyCH ?? 0)),
            //                       soLanAnConLaiTaiCuaHangNV = ((p.suatAnTrenNhanVien ?? 0) - (luuVet == null ? 0 : luuVet.soLanAnCuaNhanVienTrongThang ?? 0)),
            //                       avartar = p.avatar,
            //                       ghiChu = p.ghiChu,
            //                       tinhTrangHetHieuLuc = p.tinhTrangHetHieuLuc,

            //                   }).ToList();

            //var lstLuuVetMonAn = linqSALuuVet.tbl_SA_PhieuDangKySuatAn_LuuVets.Where(d => d.ngayDangKy.Day == DateTime.Now.Day
            // && d.ngayDangKy.Year == DateTime.Now.Year && d.ngayDangKy.Month == DateTime.Now.Month).ToList();
            //var lstMenu = (from ma in linqSALuuVet.tbl_SA_CuaHang_MonAns.ToList()
            //               select new PhieuDangKySuatAnModel
            //               {
            //                   maCuaHangMonAn = ma.maCuaHang,
            //                   tenMonAn = ma.tenMonAn,
            //                   maMonAn = ma.maMonAn,
            //                   hinhMonAn = ma.avatar,
            //                   soLuongMonAnBanDau = ma.soLuongMonAn,
            //                   soLuongMonAnDaDat = lstLuuVetMonAn.Where(d => d.maMonAn == ma.maMonAn).Count(),

            //               }).ToList();

            //ViewBag.lstMenu = lstMenu;
            //phieuDangKySuatAnModel.lstPhieuDangKySA = context.sp_SA_PhieuDangKy_GetDanhSachCuaHang(GetUser().manv).ToList();
            return PartialView("DanhSachCuaHang", listLuuVet);
        }
        public ActionResult ChiTietCuaHang(string maCuaHang, string maNhanVien)
        {
            Boolean checkQuyen = CheckQuyetInPhieuSA(GetUser().userId, "InPhieuSA");
            if (checkQuyen == false)
            {
                maNhanVien = GetUser().manv;
            }
            CuaHangModel cuaHangModel = new CuaHangModel();
            QLSALuuVetDataContext linqSALuuVet = new QLSALuuVetDataContext();
            var listLuuVet = linqSALuuVet.sp_SA_PhieuDangKySuatAn_CuaHang_LuuVet_PhanBo(maNhanVien, System.Configuration.ConfigurationManager.AppSettings["CongTy"].ToString()).Where(d => d.maCuaHang == maCuaHang).FirstOrDefault();




            return PartialView("ChiTietCuaHang", listLuuVet);
        }


        public string GetMax()
        {

            return context.tbl_SA_PhieuAns.OrderByDescending(d => d.ngayLapPhieu).Select(d => d.maPhieu).FirstOrDefault() ?? string.Empty;
        }
        public ActionResult ChonNhanVien()
        {
            #region Role user
            permission = GetPermission("InPhieuSA", BangPhanQuyen.QuyenIn);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            Boolean checkQuyen = CheckQuyetInPhieuSA(GetUser().userId, "InPhieuSA");
            if (checkQuyen == true)
            {
                //thang(DateTime.Now.Month);
                //nam(DateTime.Now.Year);
                var maPhongBanHienTai = context.GetTable<tbl_NS_NhanVienPhongBan>().Where(d => d.maNhanVien == GetUser().manv).OrderByDescending(d => d.id).Select(d => d.maPhongBan).FirstOrDefault();
                var maPhongBanCha = context.GetTable<tbl_DM_PhongBan>().Where(d => d.maPhongBan == maPhongBanHienTai).Select(d => d.maCha).FirstOrDefault();
                var tenPhongBanHienTai = context.GetTable<tbl_DM_PhongBan>().Where(d => d.maPhongBan == maPhongBanHienTai).Select(d => d.tenPhongBan).FirstOrDefault();
                var tenPhongBanCha = context.GetTable<tbl_DM_PhongBan>().Where(d => d.maPhongBan == maPhongBanCha).Select(d => d.tenPhongBan).FirstOrDefault();
                var lstPhongBans = context.GetTable<tbl_DM_PhongBan>().Where(d => d.maCha == "ABDC").ToList();
                //lstPhongBans.Insert(0, new BatDongSan.Models.DanhMuc.tbl_DM_PhongBan { maPhongBan = maPhongBanCha, tenPhongBan = tenPhongBanCha });
                //lstPhongBans.Insert(0, new BatDongSan.Models.DanhMuc.tbl_DM_PhongBan { maPhongBan = maPhongBanHienTai, tenPhongBan = tenPhongBanHienTai });
                if (maPhongBanHienTai == "PNhanSu" || CheckQuyetInPhieuSA(GetUser().userId, "InAllCty") == true)
                {
                    var maPhongBanCha2 = context.GetTable<tbl_DM_PhongBan>().Where(d => d.maPhongBan == maPhongBanCha).Select(d => d.maCha).FirstOrDefault();

                    var tenPhongBanCha2 = context.GetTable<tbl_DM_PhongBan>().Where(d => d.maPhongBan == maPhongBanCha2).Select(d => d.tenPhongBan).FirstOrDefault();
                    lstPhongBans.Insert(0, new BatDongSan.Models.DanhMuc.tbl_DM_PhongBan { maPhongBan = maPhongBanCha2, tenPhongBan = tenPhongBanCha2 });
                    lstPhongBans.Insert(1, new BatDongSan.Models.DanhMuc.tbl_DM_PhongBan { maPhongBan = maPhongBanCha, tenPhongBan = tenPhongBanCha });
                    lstPhongBans.Insert(2, new BatDongSan.Models.DanhMuc.tbl_DM_PhongBan { maPhongBan = maPhongBanHienTai, tenPhongBan = tenPhongBanHienTai });
                }
                else
                {
                    lstPhongBans.Insert(0, new BatDongSan.Models.DanhMuc.tbl_DM_PhongBan { maPhongBan = maPhongBanHienTai, tenPhongBan = tenPhongBanHienTai });
                }
                ViewBag.lstPhongBans = new SelectList(lstPhongBans, "maPhongBan", "tenPhongBan", maPhongBanHienTai);

                return View("ChonNhanVien");
            }
            else
            {
                return View("AccessDenied");
            }
        }
        public ActionResult LoadNhanVien(int? page, string searchString, string maPhongBan)
        {
            #region Role user
            permission = GetPermission("InPhieuSA", BangPhanQuyen.QuyenIn);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            LinqDanhMucDataContext linqDM = new LinqDanhMucDataContext();
            IList<sp_PB_DanhSachNhanVienResult> phongBan1s;
            phongBan1s = linqDM.sp_PB_DanhSachNhanVien(searchString, maPhongBan).ToList();
            ViewBag.isGet = "True";
            int currentPageIndex = page.HasValue ? page.Value : 1;
            ViewBag.Count = phongBan1s.Count();
            ViewBag.Search = searchString;
            ViewBag.MaPhongBan = maPhongBan;
            return PartialView("_LoadNhanVien", phongBan1s.ToPagedList(currentPageIndex, 10));
        }


        //
        // POST: /BangCap/Create

        [HttpPost]
        public ActionResult Create(FormCollection collection, HttpPostedFileBase[] files)
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
                QLSALuuVetDataContext linqSALuuVet = new QLSALuuVetDataContext();
                int flash = 0;
                tbl_SA_PhieuAn tblPhieuAn = new tbl_SA_PhieuAn();
                tblPhieuAn.maCuaHang = collection["maCuaHang"];
                tblPhieuAn.maMonAn = String.Empty;
                tblPhieuAn.nguoiLapPhieu = GetUser().manv;
                Boolean checkQuyen = CheckQuyetInPhieuSA(GetUser().userId, "InPhieuSA");
                if (checkQuyen == false)
                {
                    // Neu la nhan vien thuong thi 1 ngay / 1 
                    var checkPhieu = context.tbl_SA_PhieuAns.Where(d => d.maNhanVien == GetUser().manv
                        && d.ngayLapPhieu.Value.Day == DateTime.Now.Day && d.ngayLapPhieu.Value.Month == DateTime.Now.Month
                        && d.ngayLapPhieu.Value.Year == DateTime.Now.Year).FirstOrDefault();
                    if (checkPhieu != null)
                    {
                        ViewBag.errror = "Ngày hôm nay bạn đã lập phiếu.";
                        return View("ErrorSuatAn");
                    }
                    tblPhieuAn.maNhanVien = GetUser().manv;


                }
                else
                {
                    tblPhieuAn.maNhanVien = collection["maNhanVienThuHuong"];
                    var checkPhieu = context.tbl_SA_PhieuAns.Where(d => d.maNhanVien == tblPhieuAn.maNhanVien
                       && d.ngayLapPhieu.Value.Day == DateTime.Now.Day && d.ngayLapPhieu.Value.Month == DateTime.Now.Month
                       && d.ngayLapPhieu.Value.Year == DateTime.Now.Year).FirstOrDefault();
                    if (checkPhieu != null)
                    {
                        ViewBag.thongTinCuaHang = linqSALuuVet.tbl_SA_CuaHangs.Where(d => d.maCuaHang == checkPhieu.maCuaHang).FirstOrDefault();
                        ViewBag.errror = "Ngày hôm nay bạn đã lập phiếu.";
                        return View("ErrorSuatAn");
                    }
                }
                var listLuuVet = linqSALuuVet.sp_SA_PhieuDangKySuatAn_CuaHang_LuuVet_PhanBo(tblPhieuAn.maNhanVien, System.Configuration.ConfigurationManager.AppSettings["CongTy"].ToString())
                    .Where(d => d.maCuaHang == tblPhieuAn.maCuaHang).FirstOrDefault();



                if (listLuuVet.tongSoSuatConLaiCuaCuaHangTrongNgay > 0 && listLuuVet.soLanNhanVienDuocDangKyTrongThang > 0)
                {

                    tblPhieuAn.maPhieu = GenerateUtil.CheckLetterPA("PSA", GetMax());
                    tblPhieuAn.ngayLapPhieu = DateTime.Now;
                    tblPhieuAn.soThuTuPhucVuTaiQuan = listLuuVet.tongSoSuatDaDangKyCuaCuaHangTrongNgay + 1;
                    tblPhieuAn.soSuatAnConLaiCuaNVTaiQuan = listLuuVet.soLanNhanVienDuocDangKyTrongThang - 1;
                    tblPhieuAn.thoiGianDuKienPhucVu = "Từ 12h - 12h30";
                    if (tblPhieuAn.soThuTuPhucVuTaiQuan > 50)
                    {
                        tblPhieuAn.thoiGianDuKienPhucVu = "Từ 12h30 - 13h";
                    }
                    var dsMaVach = context.GetTable<tbl_HT_MaVach>().Where(d => d.maPhieu == tblPhieuAn.maPhieu).FirstOrDefault();
                    string maVachPhieu = String.Empty;
                    if (dsMaVach == null)
                    {
                        //Tạo mã vạch
                        GenerateBarcodeImage(tblPhieuAn.maPhieu, tblPhieuAn.maNhanVien);

                        var maVach = context.GetTable<tbl_HT_MaVach>()
                                        .Where(d => d.maPhieu == tblPhieuAn.maPhieu).FirstOrDefault();
                        maVachPhieu = maVach.tenFileDinhKemLuu;
                    }
                    else
                    {
                        maVachPhieu = dsMaVach.tenFileDinhKemLuu;
                    }
                    var tblPhongBanNV = context.GetTable<tbl_NS_NhanVienPhongBan>().Where(d => d.maNhanVien == tblPhieuAn.maNhanVien).OrderByDescending(d => d.id).Select(d => d.maPhongBan).FirstOrDefault();
                    var tenPhongBan = context.GetTable<tbl_DM_PhongBan>().Where(d => d.maPhongBan == tblPhongBanNV).Select(d => d.tenPhongBan).FirstOrDefault();
                    tblPhieuAn.maVach = maVachPhieu;

                    context.tbl_SA_PhieuAns.InsertOnSubmit(tblPhieuAn);
                    context.SubmitChanges();

                    tbl_SA_PhieuDangKySuatAn_LuuVet tblSALuuVet = new tbl_SA_PhieuDangKySuatAn_LuuVet();
                    tblSALuuVet.maNhanVien = tblPhieuAn.maNhanVien;
                    tblSALuuVet.maCongTy = System.Configuration.ConfigurationManager.AppSettings["CongTy"].ToString();
                    tblSALuuVet.maCuaHang = tblPhieuAn.maCuaHang;
                    tblSALuuVet.ngayDangKy = tblPhieuAn.ngayLapPhieu ?? DateTime.Now;
                    tblSALuuVet.maMonAn = tblPhieuAn.maMonAn;
                    tblSALuuVet.tenNhanVien = HoVaTen(tblSALuuVet.maNhanVien);
                    tblSALuuVet.tenPhongBan = tenPhongBan;
                    tblSALuuVet.sttPhucVu = tblPhieuAn.soThuTuPhucVuTaiQuan;
                    tblSALuuVet.thoiGianPhucVu = tblPhieuAn.thoiGianDuKienPhucVu;
                    tblSALuuVet.nguoiLap = GetUser().manv;
                    tblSALuuVet.tenNguoiLap = context.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == GetUser().manv).Select(d => d.ho + " " + d.ten).FirstOrDefault() ?? string.Empty;

                    linqSALuuVet.tbl_SA_PhieuDangKySuatAn_LuuVets.InsertOnSubmit(tblSALuuVet);
                    linqSALuuVet.SubmitChanges();
                    flash = 1;
                }
                if (flash == 1)
                {
                    return RedirectToAction("Details", new { id = tblPhieuAn.maPhieu });
                }
                else
                {
                    ViewBag.errror = "Cửa hàng đã hết suất hoặc bạn đã đăng ký hết số lần / cửa hàng.";
                    return View("ErrorSuatAn");
                }

            }
            catch
            {
                return View();
            }
        }
        //Save Image vào database
        private bool FileUploading(HttpPostedFileBase file)
        {
            duongDan = null;
            if (file != null && file.ContentLength > 0)
            {
                mimeType = file.ContentType;
                var date = DateTime.Now.ToString("yyyyMMdd-HHMMss") + ".jpg";
                var filePathOriginal = Server.MapPath("/Images/CuaHang/");
                string savedFileName = Path.Combine(filePathOriginal, date.ToString());
                duongDan = "/Images/CuaHang/" + date.ToString();
                file.SaveAs(savedFileName);
                return true;
            }
            else
                return false;
        }

        //
        // GET: /BangCap/Edit/5

        //public ActionResult Edit(string id)
        //{
        //    #region Role user
        //    permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
        //    if (!permission.HasValue)
        //        return View("LogIn");
        //    if (!permission.Value)
        //        return View("AccessDenied");
        //    #endregion

        //    PhieuDangKySuatAnModel phieuDangKySA = new PhieuDangKySuatAnModel();
        //    var thongTinPhieu = (from p in context.tbl_SA_PhieuAns
        //                         join q in context.tbl_SA_CuaHangs on p.maCuaHang equals q.maCuaHang
        //                         where p.maPhieu == id && p.nguoiLapPhieu == GetUser().manv
        //                         select new PhieuDangKySuatAnModel
        //                         {
        //                             maPhieu = p.maPhieu,
        //                             maNhanVien = p.nguoiLapPhieu,
        //                             ngayLapPhieu = q.ngayLap,
        //                             hoTenNhanVien = HoVaTen(p.nguoiLapPhieu),
        //                             tenCuaHang = q.tenCuaHang,
        //                             maCuaHang = q.maCuaHang,
        //                             soSuatAnTrongNgay = q.soSuatAnTrongNgay,
        //                             suatAnTrenNhanVien = q.suatAnTrenNhanVien,
        //                             tongSACuaHangDaDK = p.tongSACuaHangDaDK,
        //                             soLanNhanVienCuaHangDK = p.soLanNhanVienCuaHangDK,
        //                             avatar = q.avatar
        //                         }

        //                             ).FirstOrDefault();
        //    return PartialView("Edit", thongTinPhieu);
        //}

        //
        // POST: /BangCap/Edit/5

        //[HttpPost]
        //public ActionResult Edit(FormCollection collection, HttpPostedFileBase[] files)
        //{
        //    #region Role user
        //    permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
        //    if (!permission.HasValue)
        //        return View("LogIn");
        //    if (!permission.Value)
        //        return View("AccessDenied");
        //    #endregion
        //    try
        //    {
        //        int flash = 0;
        //        tbl_SA_PhieuAn tblPhieuAn = context.tbl_SA_PhieuAns.Where(d => d.maPhieu == collection["maPhieu"]).FirstOrDefault();
        //       var maCuaHangNew = collection["maCuaHang"];
        //       if (maCuaHangNew != tblPhieuAn.maCuaHang && maCuaHangNew != null)
        //       {
        //           var checkThongTinPhieu = context.sp_SA_PhieuDangKy_GetDanhSachCuaHang(GetUser().manv).Where(d => d.maCuaHang == maCuaHangNew).FirstOrDefault();
        //           if (checkThongTinPhieu.soLanAnConLaiTaiCuaHangNV > 0 && checkThongTinPhieu.soSuatAnConLaiTrongNgayCH > 0)
        //           {
        //               tblPhieuAn.maCuaHang = maCuaHangNew;
        //               tblPhieuAn.soLanNhanVienCuaHangDK = (checkThongTinPhieu.soLanDaAnTaiCuaHangNV ?? 0) + 1;
        //               tblPhieuAn.tongSACuaHangDaDK = (checkThongTinPhieu.soSuatAnTrongNgayDaDangKyCH ?? 0) + 1;
        //               tblPhieuAn.soThuTuPhucVuTaiQuan = (checkThongTinPhieu.soSuatAnTrongNgayDaDangKyCH ?? 0) + 1;
        //               tblPhieuAn.soSuatAnConLaiTaiQuan = (checkThongTinPhieu.soSuatAnConLaiTrongNgayCH - 1);

        //               context.SubmitChanges();
        //               flash = 1;
        //           }
        //       }
        //        if (flash == 1)
        //        {
        //            return RedirectToAction("Edit", new { id = tblPhieuAn.maPhieu });
        //        }
        //        else
        //        {
        //            return View("ErrorSuatAn");
        //        }
        //    }
        //    catch(Exception ex){
        //    return View("ErrorSuatAn");
        //    }
        //}
        public ActionResult PrintSuatAn()
        {
            #region Role user
            permission = GetPermission("InPhieuSA", BangPhanQuyen.QuyenIn);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            Boolean checkQuyen = CheckQuyetInPhieuSA(GetUser().userId, "InPhieuSA");
            if (checkQuyen == true)
            {
                //thang(DateTime.Now.Month);
                //nam(DateTime.Now.Year);
                var maPhongBanHienTai = context.GetTable<tbl_NS_NhanVienPhongBan>().Where(d => d.maNhanVien == GetUser().manv).OrderByDescending(d => d.id).Select(d => d.maPhongBan).FirstOrDefault();
                var maPhongBanCha = context.GetTable<tbl_DM_PhongBan>().Where(d => d.maPhongBan == maPhongBanHienTai).Select(d => d.maCha).FirstOrDefault();
                var tenPhongBanHienTai = context.GetTable<tbl_DM_PhongBan>().Where(d => d.maPhongBan == maPhongBanHienTai).Select(d => d.tenPhongBan).FirstOrDefault();
                var tenPhongBanCha = context.GetTable<tbl_DM_PhongBan>().Where(d => d.maPhongBan == maPhongBanCha).Select(d => d.tenPhongBan).FirstOrDefault();
                var lstPhongBans = context.GetTable<tbl_DM_PhongBan>().Where(d => d.maCha == "ABDC").ToList();
                if (maPhongBanHienTai == "PNhanSu" || CheckQuyetInPhieuSA(GetUser().userId, "InAllCty") == true)
                {
                    var maPhongBanCha2 = context.GetTable<tbl_DM_PhongBan>().Where(d => d.maPhongBan == maPhongBanCha).Select(d => d.maCha).FirstOrDefault();

                    var tenPhongBanCha2 = context.GetTable<tbl_DM_PhongBan>().Where(d => d.maPhongBan == maPhongBanCha2).Select(d => d.tenPhongBan).FirstOrDefault();
                    lstPhongBans.Insert(0, new BatDongSan.Models.DanhMuc.tbl_DM_PhongBan { maPhongBan = maPhongBanCha2, tenPhongBan = tenPhongBanCha2 });
                    lstPhongBans.Insert(1, new BatDongSan.Models.DanhMuc.tbl_DM_PhongBan { maPhongBan = maPhongBanCha, tenPhongBan = tenPhongBanCha });
                    lstPhongBans.Insert(2, new BatDongSan.Models.DanhMuc.tbl_DM_PhongBan { maPhongBan = maPhongBanHienTai, tenPhongBan = tenPhongBanHienTai });
                }
                else
                {
                    lstPhongBans.Insert(0, new BatDongSan.Models.DanhMuc.tbl_DM_PhongBan { maPhongBan = maPhongBanHienTai, tenPhongBan = tenPhongBanHienTai });
                }

                ViewBag.lstPhongBans = new SelectList(lstPhongBans, "maPhongBan", "tenPhongBan", maPhongBanHienTai);

                return View("PrintSuatAn");
            }
            else
            {
                return View("AccessDenied");
            }

        }
        private void thang(int value)
        {
            Dictionary<int, string> dics = new Dictionary<int, string>();
            for (int i = 0; i < 13; i++)
            {
                dics[i] = i.ToString();
            }
            ViewData["thang"] = new SelectList(dics, "Key", "Value", value);
            ViewData["thangBP"] = new SelectList(dics, "Key", "Value", value);

        }
        private void nam(int value)
        {
            Dictionary<int, string> dics = new Dictionary<int, string>();
            for (int i = 2015; i < 2031; i++)
            {
                dics[i] = i.ToString();
            }
            ViewData["nam"] = new SelectList(dics, "Key", "Value", value);
            ViewData["namBP"] = new SelectList(dics, "Key", "Value", value);

        }

        public ActionResult PartialPrintSuatAn(string searchString, string maPhongBan)
        {
            #region Role user
            permission = GetPermission("InPhieuSA", BangPhanQuyen.QuyenIn);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            QLSALuuVetDataContext linqSALuuVet = new QLSALuuVetDataContext();
            LinqHeThongDataContext contextHTNew = new LinqHeThongDataContext();
            var lstMaNhanVien = contextHTNew.sp_PB_DanhSachNhanVien_Tree(searchString, maPhongBan, null, 0, 1000000).Select(d => d.maNhanVien).ToList();
            ViewBag.lstCuaHang = linqSALuuVet.tbl_SA_CuaHangs.ToList();
            ViewBag.lstMonAn = linqSALuuVet.tbl_SA_CuaHang_MonAns.ToList();
            var thongTinPhieu = (from p in context.tbl_SA_PhieuAns

                                 where lstMaNhanVien.Contains(p.maNhanVien) && DateTime.Now.Month == p.ngayLapPhieu.Value.Month
                                 && DateTime.Now.Year == p.ngayLapPhieu.Value.Year && DateTime.Now.Day == p.ngayLapPhieu.Value.Day
                                 select new PhieuDangKySuatAnModel
                                 {
                                     maPhieu = p.maPhieu,
                                     maNhanVien = p.maNhanVien,
                                     ngayLapPhieu = p.ngayLapPhieu,
                                     hoTenNhanVien = HoVaTen(p.maNhanVien),
                                     maCuaHang = p.maCuaHang,
                                     maMonAn = p.maMonAn,
                                     //tenCuaHang = lstCuaHang.Where(d=>d.maCuaHang == p.maCuaHang).Select(d=>d.tenCuaHang).FirstOrDefault(),
                                     // maCuaHang = lstCuaHang.Where(d => d.maCuaHang == p.maCuaHang).Select(d => d.maCuaHang).FirstOrDefault(),
                                     //suatAnTrenNhanVien = lstCuaHang.Where(d => d.maCuaHang == p.maCuaHang).Select(d => d.suatAnTrenNhanVien).FirstOrDefault(),
                                     soSuatAnConLaiCuaNVTaiQuan = p.soSuatAnConLaiCuaNVTaiQuan,
                                     soThuTuPhucVuTaiQuan = p.soThuTuPhucVuTaiQuan,
                                     thoiGianDuKienPhucVu = p.thoiGianDuKienPhucVu,
                                     maVach = p.maVach,
                                     //avatar = lstCuaHang.Where(d => d.maCuaHang == p.maCuaHang).Select(d => d.avatar).FirstOrDefault(),
                                     //maMonAn = lstMonAn.Where(d=>d.maMonAn == p.maMonAn).Select(d=>d.maMonAn).FirstOrDefault(),
                                     //tenMonAn = lstMonAn.Where(d => d.maMonAn == p.maMonAn).Select(d => d.tenMonAn).FirstOrDefault(),
                                     // hinhMonAn = lstMonAn.Where(d => d.maMonAn == p.maMonAn).Select(d => d.avatar).FirstOrDefault(),
                                 }

                                         ).ToList();
            return View("PartialPrintSuatAn", thongTinPhieu);
        }

        //public ActionResult PrintDocumentSuatAn(string maPhieu)
        //{
        //    try
        //    {

        //        bool isPDF = true;
        //        string folderName = "BatDongSanAppGenerate";
        //        string driveLocal = Path.GetPathRoot(Environment.SystemDirectory);

        //        //Tạo folder trên ổ đĩa
        //        string folderPath = driveLocal + folderName;
        //        Directory.CreateDirectory(folderPath);

        //        //Set full quyền cho folder 
        //        GenerateUtil.SetFullControlPermissions(folderPath);

        //        String sourceFile = @"/UploadFiles/QLSA/mauPhieuAnTrua.docx",
        //               destinationFile = driveLocal + folderName + "/" + DateTime.Now.ToString("ddMMyyyyHHmmss") + "Phieu_An_Trua.docx";

        //        sourceFile = Server.MapPath(sourceFile);
        //        GenerateUtil.CopyFile(sourceFile, destinationFile);
        //        // Create the document in memory:
        //        var doc = DocX.Load(destinationFile);
        //        FillMauInSuatAn(doc, maPhieu);
        //        string ouputPDF = Path.GetDirectoryName(destinationFile) + "/" + DateTime.Now.ToString("ddMMyyyyHHmmss") + "Phieu_An_Trua.pdf";

        //        string downloadFileName = "Phieu_An_Trua" + maPhieu;

        //        byte[] buf;
        //        if (isPDF)
        //        {
        //            //Convert sang file PDF
        //            GenerateUtil.AsposeConvertTo(destinationFile, ouputPDF);

        //            //Xóa file word 
        //            System.IO.File.Delete(destinationFile);

        //            //Chuyển file pdf sang memorystream 

        //            using (MemoryStream memoryStream = new MemoryStream())
        //            {
        //                using (Stream stream = System.IO.File.OpenRead(ouputPDF))
        //                {
        //                    stream.CopyTo(memoryStream);
        //                }
        //                memoryStream.Position = 0;
        //                buf = new byte[memoryStream.Length];
        //                buf = memoryStream.ToArray();
        //            }

        //            //Xóa file PDF
        //            System.IO.File.Delete(ouputPDF);

        //            Response.AddHeader("Content-Type: ", "application/x-pdf ; charset=utf-8");
        //            Response.AddHeader("Content-Disposition", "attachment;filename=" + downloadFileName + ".pdf");
        //            Response.AddHeader("Content-Length", buf.Length.ToString());
        //            return File(buf, "application/x-pdf");
        //        }

        //        //Chuyển file pdf sang memorystream                 
        //        using (MemoryStream memoryStream = new MemoryStream())
        //        {
        //            using (Stream stream = System.IO.File.OpenRead(destinationFile))
        //            {
        //                stream.CopyTo(memoryStream);
        //            }
        //            memoryStream.Position = 0;
        //            buf = new byte[memoryStream.Length];
        //            buf = memoryStream.ToArray();
        //        }

        //        //Xóa file word
        //        System.IO.File.Delete(destinationFile);

        //        Response.AddHeader("Content-Type: ", "application/vnd.openxmlformats-officedocument.wordprocessingml.document ; charset=utf-8");
        //        Response.AddHeader("Content-Disposition", "attachment;filename=" + downloadFileName + ".docx");
        //        Response.AddHeader("Content-Length", buf.Length.ToString());
        //        return File(buf.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        //    }
        //    catch (Exception ex)
        //    {
        //        return View();
        //    }
        //}

        //public void FillMauInSuatAn(DocX doc, string maPhieu)
        //{

        //    var thongTinPhieu = (from p in context.tbl_SA_PhieuAns
        //                         join q in context.tbl_SA_CuaHangs on p.maCuaHang equals q.maCuaHang
        //                         where p.maPhieu == maPhieu && p.nguoiLapPhieu == GetUser().manv
        //                         select new PhieuDangKySuatAnModel
        //                         {
        //                             maPhieu = p.maPhieu,
        //                             maNhanVien = p.nguoiLapPhieu,
        //                             ngayLapPhieu = p.ngayLapPhieu,
        //                             hoTenNhanVien = HoVaTen(p.nguoiLapPhieu),
        //                             tenCuaHang = q.tenCuaHang,
        //                             maCuaHang = q.maCuaHang,
        //                             soSuatAnTrongNgay = q.soSuatAnTrongNgay,
        //                             suatAnTrenNhanVien = q.suatAnTrenNhanVien,
        //                             tongSACuaHangDaDK = p.tongSACuaHangDaDK,
        //                             soLanNhanVienCuaHangDK = p.soLanNhanVienCuaHangDK,
        //                             soThuTuPhucVuTaiQuan = p.soThuTuPhucVuTaiQuan,
        //                             soSuatAnConLaiTaiQuan = p.soSuatAnConLaiTaiQuan,
        //                             avatar = q.avatar
        //                         }

        //                            ).FirstOrDefault();
        //    doc.ReplaceText("<<ngayLapPhieu>>", String.Format("{0: dd/MM/yyyy}", thongTinPhieu.ngayLapPhieu));
        //    doc.ReplaceText("<<maNhanVien>>", thongTinPhieu.maNhanVien);
        //    doc.ReplaceText("<<hoTenNhanVien>>", thongTinPhieu.hoTenNhanVien);
        //    doc.ReplaceText("<<tenQuanAnDuocChon>>", thongTinPhieu.tenCuaHang);
        //    doc.ReplaceText("<<soThuTuPhucVu>>", String.Format("{0:#,##0}", Math.Round((decimal)thongTinPhieu.soThuTuPhucVuTaiQuan, 0, MidpointRounding.AwayFromZero)));
        //    doc.ReplaceText("<<soSuatAnConLaiDuocSuDungTaiQuan>>", String.Format("{0:#,##0}", Math.Round((decimal)thongTinPhieu.soSuatAnConLaiTaiQuan, 0, MidpointRounding.AwayFromZero)));



        //    doc.Save();
        //}
        //
        // POST: /BangCap/Delete/5

        //[HttpPost]
        //public ActionResult Delete(string id)
        //{
        //    #region Role user
        //    permission = GetPermission(MCV, BangPhanQuyen.QuyenXoa);
        //    if (!permission.HasValue)
        //        return View("LogIn");
        //    if (!permission.Value)
        //        return View("AccessDenied");
        //    #endregion

        //    try
        //    {


        //        var tblPhieu = context.tbl_SA_PhieuAns.Where(d => d.maPhieu == id).FirstOrDefault();
        //        context.tbl_SA_PhieuAns.DeleteOnSubmit(tblPhieu);
        //        context.SubmitChanges();
        //        return RedirectToAction("Index");
        //    }
        //    catch
        //    {
        //        return View();
        //    }
        //}


        public ActionResult LoadDanhSachPhanBo(int thang, int nam)
        {

            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            QLSuatAnDataContext qLSAContext = new QLSuatAnDataContext();

            var nhanVien = qLSAContext.sp_SA_DangKySuatAn_CongChuanNhanVien_Cienco6(thang, nam).Where(d => d.maNhanVien == GetUser().manv).FirstOrDefault();

            int congChuan = 0;

            if (nhanVien != null)
            {
                congChuan = nhanVien.congChuan ?? 0;
            }


            QLSALuuVetDataContext linqSALuuVet = new QLSALuuVetDataContext();
            var data = linqSALuuVet.sp_SA_NhanVienDangKySuatAnTrongThang_Index(thang, nam, congChuan).ToList();

            ViewData["lsDanhSach"] = data;

            return PartialView("_LoadDanhSachPhanBo");
        }
    }
}
