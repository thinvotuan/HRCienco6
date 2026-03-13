using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Utils.Paging;
using BatDongSan.Helper.Utils;
using System.Data;
using System.Globalization;
using System.Text;
using System.IO;
using BatDongSan.Models.NhanSu;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System.Configuration;
using System.Net.Mail;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel.Contrib;
using Worldsoft.Mvc.Web.Util;
using NPOI.SS.Util;
namespace BatDongSan.Controllers.TinhCong
{
    public class BangLuongNVController : ApplicationController
    {

        private LinqNhanSuDataContext nhanSuContext = new LinqNhanSuDataContext();
        private LinqThuanVietDataContext lqThuanViet = new LinqThuanVietDataContext();
        private readonly string MCV = "BangLuongNV";
        private IList<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan> phongBans;
        private StringBuilder buildTree;
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
            buildTree = new StringBuilder();
            phongBans = nhanSuContext.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan>().ToList();
            buildTree = TreePhongBanAjax.BuildTreeDepartment(phongBans);
            ViewBag.PhongBans = buildTree.ToString();
            thang(DateTime.Now.Month);
            nam(DateTime.Now.Year);
            return View("");

        }

        public ActionResult LoadBangLuongNV(string maPhongBan, string qSearch, int thang, int nam, int _page = 0)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            int page = _page == 0 ? 1 : _page;
            int pIndex = page;
            int total = nhanSuContext.sp_NS_BangLuongNhanVien(maPhongBan, thang, nam, qSearch).Count();
            PagingLoaderController("/BangLuongNV/LoadBangLuongNV/", total, page, "?qsearch=" + qSearch);
            ViewData["lsDanhSach"] = nhanSuContext.sp_NS_BangLuongNhanVien(maPhongBan, thang, nam, qSearch).Skip(start).Take(offset).ToList();

            ViewData["qSearch"] = qSearch;
            var kqCheck = 0;
            var checkEx = nhanSuContext.tbl_DuyetBangLuongNVs.Where(t => t.nam == nam && t.thang == thang).FirstOrDefault();
            if (checkEx != null)
            {
                kqCheck = 1;
            }
            ViewData["kqCheck"] = kqCheck;
            return PartialView("_LoadBangLuongNV");
        }
        public ActionResult ViewChiTietLuong(string thang, string nam, string maNhanVien)
        {
            #region Role user
            permission = GetPermission("XemBangLuong", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            var dsMauIn = nhanSuContext.GetTable<BatDongSan.Models.HeThong.Sys_PrintTemplate>().Where(d => d.maMauIn == "MIBLNV").FirstOrDefault();
            string noiDung = string.Empty;
            var ds = nhanSuContext.tbl_NS_BangLuongNhanViens
                               .Where(t => t.maNhanVien == maNhanVien && t.thang.ToString() == thang && nam == t.nam.ToString()).FirstOrDefault();
            ViewData["chiTiet"] = dsMauIn;
            if (dsMauIn != null)
            {
                var tongKhauTru = (ds.baoHiem) + (ds.thue ?? 0) + (ds.doanPhi ?? 0) + (ds.dangPhi ?? 0);
                noiDung = dsMauIn.html.Replace("{$thang}", thang)
                    .Replace("{$nam}", nam)
                    .Replace("{$hoVaTen}", ds.hoTen)
                    .Replace("{$tongLuong}", String.Format("{0:###,##0}", ds.tongLuong))
                    .Replace("{$khoanBSLVSLuongBH}", String.Format("{0:###,##0}", ds.luongDongBaoHiem))
                    .Replace("{$luongThoaThuan}", String.Format("{0:###,##0}", ds.khoanBoSungLuong))
                    .Replace("{$ngayCongChuan}", Convert.ToString(ds.ngayCongChuan))
                    .Replace("{$ngayCongTinhLuong}", Convert.ToString(ds.tongNgayCong))
                    .Replace("{$ngayCong}", Convert.ToString(ds.soNgayQuet))
                    .Replace("{$ngayCongTac}", Convert.ToString(ds.soNgayCongTac))
                    .Replace("{$ngayCongTangCa}", Convert.ToString(ds.soNgayTangCaThuong + ds.soNgayTangCaChuNhat + ds.soNgayTangCaNgayLe))
                    .Replace("{$soNgayPhepLuyKeThangTruoc}", Convert.ToString(ds.soNgayPhepLuyKeThangTruoc))

                    .Replace("{$nghiPhep}", Convert.ToString(ds.soNgayNghiPhep))
                    .Replace("{$nghiLeTet}", Convert.ToString(ds.soNgayNghiLe))

                    .Replace("{$luongTangCa}", String.Format("{0:###,##0}", (ds.luongTangCa ?? 0)))
                    .Replace("{$luongTheoCongChinh}", String.Format("{0:###,##0}", (ds.luongThucTe ?? 0) + (ds.luongNghiPhep ?? 0)))

                    .Replace("{$luongTheoNgayCong}", String.Format("{0:###,##0}", (ds.luongThucTe ?? 0) + (ds.luongNghiPhep ?? 0) + (ds.luongTangCa ?? 0)))

                    .Replace("{$congChoViec}", Convert.ToString(ds.soNgayCongChoViec))
                    .Replace("{$luongChoViec}", String.Format("{0:###,##0}", ds.luongChoViec))
                    .Replace("{$tongTienPhuCap}", String.Format("{0:###,##0}", (ds.phuCapTienAn ?? 0) + (ds.phuCapDienThoai ?? 0) + (ds.phuCapKhacThueNhaXeXang ?? 0)))
                    .Replace("{$tienAnGiuaCa}", String.Format("{0:###,##0}", ds.phuCapTienAn))
                    .Replace("{$tienDienThoai}", String.Format("{0:###,##0}", ds.phuCapDienThoai))
                    .Replace("{$phuCapKhacTNXX}", String.Format("{0:###,##0}", ds.phuCapKhacThueNhaXeXang))
                    .Replace("{$tienBHXH}", String.Format("{0:###,##0}", ds.baoHiem > 0 ? ds.baoHiemXH : 0))
                    .Replace("{$tienBHYT}", String.Format("{0:###,##0}", ds.baoHiem > 0 ? ds.baoHiemYTe : 0))
                    .Replace("{$tienBHTN}", String.Format("{0:###,##0}", ds.baoHiem > 0 ? ds.baoHiemTN : 0))
                    .Replace("{$tienTTNCT}", String.Format("{0:###,##0}", ds.thue))
                    .Replace("{$tienDoanPhi}", String.Format("{0:###,##0}", ds.doanPhi))
                    .Replace("{$tienDangPhi}", String.Format("{0:###,##0}", ds.dangPhi))
                    .Replace("{$tienPhuCap}", String.Format("{0:###,##0}", ds.phuCapKhac))
                    //double truyLanh = (item1.truyLanh ?? 0);
                    //    if ((item1.TTTLLuong ?? 0) > 0) truyLanh += (item1.TTTLLuong ?? 0);
                    //    if ((item1.TTTLThue ?? 0) > 0) truyLanh += (item1.TTTLThue ?? 0);

                    //    double truyThu = (item1.truyThu ?? 0);
                    //    if ((item1.TTTLLuong ?? 0) < 0) truyThu += (item1.TTTLLuong ?? 0);
                    //    if ((item1.TTTLThue ?? 0) < 0) truyThu += (item1.TTTLThue ?? 0);
                    .Replace("{$tienTruyThu}", String.Format("{0:###,##0}", ds.truyThu + ((ds.TTTLLuong ?? 0) < 0 ? ds.TTTLLuong : 0) + ((ds.TTTLThue ?? 0) < 0 ? ds.TTTLThue : 0)))
                    .Replace("{tongPhuCapTL}", String.Format("{0:###,##0}", ds.truyLanh - ds.truyThu + ds.phuCapKhac))
                    .Replace("{$tienTruyLanh}", String.Format("{0:###,##0}", ds.truyLanh + ((ds.TTTLLuong ?? 0) > 0 ? ds.TTTLLuong : 0) + ((ds.TTTLThue ?? 0) > 0 ? ds.TTTLThue : 0)))
                    .Replace("{$tienGiamTruGC}", String.Format("{0:###,##0}", (ds.giamTruBanThan ?? 0) + (ds.giamTruNguoiPhuThuoc ?? 0)))
                    .Replace("{$tongThucNhanLuong}", String.Format("{0:###,##0}", ds.thucLanh))
                    .Replace("{$tongKhauTru}", String.Format("{0:###,##0}", tongKhauTru));

            }
            ViewBag.NoiDung = noiDung;
            // return PartialView("_ViewChiTietLuong");
            return PartialView("_ViewChiTietLuongTemplate");
        }


        public ActionResult LoadBangLuongTheoBP(int thang, int nam, int _page = 0)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            int page = _page == 0 ? 1 : _page;
            int pIndex = page;
            int total = nhanSuContext.sp_NS_BangLuongTheoBoPhan(thang, nam).Count();
            PagingLoaderController("/BangLuongNV/LoadBangLuongTheoBP/", total, page, "?qsearch=" + null);
            ViewData["lsDanhSachBP"] = nhanSuContext.sp_NS_BangLuongTheoBoPhan(thang, nam).Skip(start).Take(offset).ToList();


            return PartialView("_LoadBangLuongTheoBP");
        }
        public ActionResult BangLuongNV()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            buildTree = new StringBuilder();
            phongBans = nhanSuContext.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan>().ToList();
            buildTree = TreePhongBanAjax.BuildTreeDepartment(phongBans);
            ViewBag.PhongBans = buildTree.ToString();
            thang(DateTime.Now.Month);
            nam(DateTime.Now.Year);
            return View("");

        }
        public ActionResult BangLuongChuyenNN()
        {
            #region Role user
            permission = GetPermission("BangLuongNVNN", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            buildTree = new StringBuilder();
            phongBans = nhanSuContext.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan>().ToList();
            buildTree = TreePhongBanAjax.BuildTreeDepartment(phongBans);
            ViewBag.PhongBans = buildTree.ToString();
            thang(DateTime.Now.Month);
            nam(DateTime.Now.Year);
            return View("");

        }

        public ActionResult LoadBangLuongChuyenNN(string maPhongBan, string qSearch, int thang, int nam, int _page = 0)
        {
            #region Role user
            permission = GetPermission("BangLuongNVNN", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            int page = _page == 0 ? 1 : _page;
            int pIndex = page;
            int total = nhanSuContext.sp_NS_BangLuongChuyenNganHang(thang, nam, maPhongBan, qSearch).Count();
            PagingLoaderController("/BangLuongNV/BangLuongChuyenNN/", total, page, "?qsearch=" + qSearch);
            ViewData["lsDanhSach"] = nhanSuContext.sp_NS_BangLuongChuyenNganHang(thang, nam, maPhongBan, qSearch).Skip(start).Take(offset).ToList();

            ViewData["qSearch"] = qSearch;
            return PartialView("_LoadBangLuongChuyenNN");
        }

        public ActionResult UpdateLuongNV(int? thang, int? nam)
        {


            try
            {
                #region Role user
                permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion
                //Check 
                var result = new { kq = false };
                var checkEx = nhanSuContext.tbl_DuyetBangLuongNVs.Where(t => t.nam == nam && t.thang == thang).FirstOrDefault();
                if (checkEx != null)
                {
                    return Json(result, JsonRequestBehavior.AllowGet);
                }
                //End check
                var list = nhanSuContext.sp_NS_TinhLuongNhanVien(thang, nam);
                result = new { kq = true };
                SaveActiveHistory("Tính lương nhân viên tháng: " + thang + " năm: " + nam);
                return Json(result, JsonRequestBehavior.AllowGet);

            }
            catch
            {
                return View();
            }

        }
        public ActionResult DuyetLuongNV(int thang, int nam)
        {


            try
            {
                #region Role user
                permission = GetPermission(MCV, BangPhanQuyen.QuyenDuyet);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion
                //Check 
                var result = new { kq = false };
                var checkEx = nhanSuContext.tbl_DuyetBangLuongNVs.Where(t => t.nam == nam && t.thang == thang).FirstOrDefault();
                if (checkEx != null)
                {
                    return Json(result, JsonRequestBehavior.AllowGet);
                }
                //End check
                // Insert Row 
                tbl_DuyetBangLuongNV tblDuyetBL = new tbl_DuyetBangLuongNV();
                tblDuyetBL.nam = nam;
                tblDuyetBL.thang = thang;
                tblDuyetBL.ngayDuyet = DateTime.Now;
                tblDuyetBL.nguoiDuyet = GetUser().manv;
                nhanSuContext.tbl_DuyetBangLuongNVs.InsertOnSubmit(tblDuyetBL);
                nhanSuContext.SubmitChanges();
                // End Insert Row
                // Check Exist
                var list = nhanSuContext.tbl_DuyetBangLuongNVs.Where(t => t.nam == nam && t.thang == thang).FirstOrDefault();

                if (list != null)
                {
                    result = new { kq = true };
                }
                SaveActiveHistory("Duyệt lương nhân viên tháng: " + thang + " năm: " + nam);
                return Json(result, JsonRequestBehavior.AllowGet);

            }
            catch
            {
                return View();
            }

        }
        public ActionResult SendBangLuongNV(int thang, int nam)
        {

            string maPhongBan = "";
            string qSearch = "";
            var list = nhanSuContext.tbl_DuyetBangLuongNVs.Where(t => t.nam == nam && t.thang == thang).FirstOrDefault();
            var result = new { kq = false };
            if (list == null)
            {
                return Json(result, JsonRequestBehavior.AllowGet);
            }

            var listSendMails = nhanSuContext.sp_NS_BangLuongNhanVien(maPhongBan, thang, nam, qSearch).ToList();
            foreach (var ds in listSendMails)
            {
                var dsMauIn = nhanSuContext.GetTable<BatDongSan.Models.HeThong.Sys_PrintTemplate>().Where(d => d.maMauIn == "MIBLNV").FirstOrDefault();
                string noiDung = string.Empty;
                var tongKhauTru = (ds.baoHiemXH ?? 0) + (ds.baoHiemYTe ?? 0) + (ds.baoHiemTN ?? 0) + (ds.thue ?? 0) + (ds.doanPhi ?? 0) + (ds.dangPhi ?? 0);
                //Replace bang luong send mail
                noiDung = dsMauIn.html.Replace("{$thang}", Convert.ToString(thang))
                    .Replace("{$nam}", Convert.ToString(nam))
                    .Replace("{$hoVaTen}", ds.hoTen)
                    .Replace("{$tongLuong}", String.Format("{0:###,##0}", ds.tongLuong))
                    .Replace("{$khoanBSLVSLuongBH}", String.Format("{0:###,##0}", ds.luongDongBaoHiem))
                    .Replace("{$luongThoaThuan}", String.Format("{0:###,##0}", ds.khoanBoSungLuong))
                    .Replace("{$ngayCongChuan}", Convert.ToString(ds.ngayCongChuan))
                    .Replace("{$ngayCongTinhLuong}", Convert.ToString(ds.tongNgayCong))
                    .Replace("{$ngayCong}", Convert.ToString(ds.soNgayQuet))
                    .Replace("{$ngayCongTac}", Convert.ToString(ds.soNgayCongTac))
                    .Replace("{$nghiPhep}", Convert.ToString(ds.soNgayNghiPhep))
                    .Replace("{$nghiLeTet}", Convert.ToString(ds.soNgayNghiLe))

                    .Replace("{$luongTangCa}", String.Format("{0:###,##0}", (ds.luongTangCa ?? 0)))
                    .Replace("{$luongTheoCongChinh}", String.Format("{0:###,##0}", (ds.luongThucTe ?? 0) + (ds.luongNghiPhep ?? 0)))
                    .Replace("{$luongTheoNgayCong}", String.Format("{0:###,##0}", (ds.luongThucTe ?? 0) + (ds.luongNghiPhep ?? 0) + (ds.luongTangCa ?? 0)))

                    .Replace("{$congChoViec}", Convert.ToString(ds.soNgayCongChoViec))
                    .Replace("{$luongChoViec}", String.Format("{0:###,##0}", ds.luongChoViec))
                    .Replace("{$tongTienPhuCap}", String.Format("{0:###,##0}", ds.phuCapTienAn + ds.phuCapDienThoai))
                    .Replace("{$tienAnGiuaCa}", String.Format("{0:###,##0}", ds.phuCapTienAn))
                    .Replace("{$tienDienThoai}", String.Format("{0:###,##0}", ds.phuCapDienThoai))
                    .Replace("{$tienBHXH}", String.Format("{0:###,##0}", ds.baoHiemXH))
                    .Replace("{$tienBHYT}", String.Format("{0:###,##0}", ds.baoHiemYTe))
                    .Replace("{$tienBHTN}", String.Format("{0:###,##0}", ds.baoHiemTN))
                    .Replace("{$tienTTNCT}", String.Format("{0:###,##0}", ds.thue))
                    .Replace("{$tienDoanPhi}", String.Format("{0:###,##0}", ds.doanPhi))
                    .Replace("{$tienDangPhi}", String.Format("{0:###,##0}", ds.dangPhi))
                    .Replace("{$tienPhuCap}", String.Format("{0:###,##0}", ds.phuCapKhac))
                    .Replace("{$tienTruyThu}", String.Format("{0:###,##0}", ds.truyThu))
                    .Replace("{tongPhuCapTL}", String.Format("{0:###,##0}", ds.truyLanh - ds.truyThu + ds.phuCapKhac))
                    .Replace("{$tienTruyLanh}", String.Format("{0:###,##0}", ds.truyLanh))
                    .Replace("{$tienGiamTruGC}", String.Format("{0:###,##0}", (ds.giamTruBanThan ?? 0) + (ds.giamTruNguoiPhuThuoc ?? 0)))
                    .Replace("{$tongThucNhanLuong}", String.Format("{0:###,##0}", ds.thucLanh))
                    .Replace("{$tongKhauTru}", String.Format("{0:###,##0}", tongKhauTru));
                //End
                // Code send mail
                MailHelper mailInit = new MailHelper(); // lay cac tham so trong webconfig
                System.Text.StringBuilder content = new System.Text.StringBuilder();

                content.Append("<h3>Email từ hệ thống nhân sự</h3>");
                content.Append("<p>Xin chào: " + ds.hoTen + " !</p>");
                //Content
                content.Append(noiDung);

                //End content
                content.Append("<p style='font-style:italic'>Thanks and Regards!</p>");
                //Send only email is @thuanviet.com.vn
                string[] array01 = ds.email.ToLower().Split('@');
                //string string2 = ConfigurationManager.AppSettings["OnlySend"]; //get domain from config files
                //string[] array1 = string2.Split(',');
                // bool EmailofThuanViet;
                //EmailofThuanViet = array1.Contains(array01[1]);
                // if (emailNV == "" || emailNV == null || EmailofThuanViet == false)
                // {
                //    return false;
                // }
                MailAddress toMail = new MailAddress(ds.email, ds.hoTen); // goi den mail
                mailInit.ToMail = toMail;
                mailInit.Body = content.ToString();
                mailInit.SendMail();
                // End code send mail
            }
            result = new { kq = true };
            SaveActiveHistory("Send mail bảng lương nhân viên: " + thang + " năm: " + nam);
            return Json(result, JsonRequestBehavior.AllowGet);
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

        #region Xuat File Bang Luong Nhan Vien
        public void XuatFileBLNV(int thang, int nam)
        {
            try
            {
                string maPhongBan = "";
                string qSearch = "";
                var filename = "";
                var virtualPath = HttpRuntime.AppDomainAppVirtualPath;

                var fileStream = new FileStream(System.Web.HttpContext.Current.Server.MapPath(virtualPath + @"\Content\Report\ReportTemplatecc.xls"), FileMode.Open, FileAccess.Read);

                var workbook = new HSSFWorkbook(fileStream, true);
                filename += "BangLuongNV_" + nam + "_" + thang + ".xls";

                #region format style excel cell
                /*style title start*/
                //tạo font cho các title
                //font tiêu đề 
                HSSFFont hFontTieuDe = (HSSFFont)workbook.CreateFont();
                hFontTieuDe.FontHeightInPoints = 11;
                hFontTieuDe.Boldweight = 100 * 10;
                hFontTieuDe.FontName = "Times New Roman";
                //hFontTieuDe.Color = HSSFColor.BLUE.index;

                //font tiêu đề 
                HSSFFont hFontTieuDeUnderline = (HSSFFont)workbook.CreateFont();
                hFontTieuDeUnderline.FontHeightInPoints = 11;
                hFontTieuDeUnderline.Boldweight = 100 * 10;
                hFontTieuDeUnderline.FontName = "Times New Roman";
                hFontTieuDeUnderline.Underline = 1;
                //hFontTieuDe.Color = HSSFColor.BLUE.index;


                HSSFFont hFontTieuDeItalic = (HSSFFont)workbook.CreateFont();
                hFontTieuDeItalic.FontHeightInPoints = 11;
                //hFontTieuDeItalic.Boldweight = 100 * 10;
                hFontTieuDeItalic.FontName = "Times New Roman";
                hFontTieuDeItalic.IsItalic = true;

                //hFontTieuDe.Color = HSSFColor.BLUE.index;

                HSSFFont hFontTieuDeNormal = (HSSFFont)workbook.CreateFont();
                hFontTieuDeNormal.FontHeightInPoints = 11;
                //hFontTieuDeItalic.Boldweight = 100 * 10;
                hFontTieuDeNormal.FontName = "Times New Roman";
                hFontTieuDeNormal.IsItalic = false;
                HSSFFont hFontTieuDeBold = (HSSFFont)workbook.CreateFont();
                hFontTieuDeBold.FontHeightInPoints = 11;
                hFontTieuDeItalic.Boldweight = (short)FontBoldWeight.BOLD;
                hFontTieuDeBold.FontName = "Times New Roman";
                hFontTieuDeBold.IsItalic = false;


                HSSFFont hFontTieuDeLarge = (HSSFFont)workbook.CreateFont();
                hFontTieuDeLarge.FontHeightInPoints = 16;
                hFontTieuDeLarge.Boldweight = 100 * 10;
                hFontTieuDeLarge.FontName = "Times New Roman";
                //hFontTieuDeLarge.Color = HSSFColor.BLUE.index;

                //font tiêu đề 
                HSSFFont hFontTongGiaTriHT = (HSSFFont)workbook.CreateFont();
                hFontTongGiaTriHT.FontHeightInPoints = 11;
                hFontTongGiaTriHT.Boldweight = (short)FontBoldWeight.BOLD;
                hFontTongGiaTriHT.FontName = "Times New Roman";
                hFontTongGiaTriHT.Color = HSSFColor.BLACK.index;

                //font thông tin bảng tính
                HSSFFont hFontTT = (HSSFFont)workbook.CreateFont();
                hFontTT.IsItalic = true;
                hFontTT.Boldweight = (short)FontBoldWeight.BOLD;
                hFontTT.Color = HSSFColor.BLACK.index;
                hFontTT.FontName = "Times New Roman";
                hFontTieuDe.FontHeightInPoints = 11;

                //font chứ hoa đậm
                HSSFFont hFontNommalUpper = (HSSFFont)workbook.CreateFont();
                hFontNommalUpper.Boldweight = (short)FontBoldWeight.BOLD;
                hFontNommalUpper.Color = HSSFColor.BLACK.index;
                hFontNommalUpper.FontName = "Times New Roman";

                //font chữ bình thường
                HSSFFont hFontNommal = (HSSFFont)workbook.CreateFont();
                hFontNommal.Color = HSSFColor.BLACK.index;
                hFontNommal.FontName = "Times New Roman";

                //font chữ bình thường đậm
                HSSFFont hFontNommalBold = (HSSFFont)workbook.CreateFont();
                hFontNommalBold.Color = HSSFColor.BLACK.index;
                hFontNommalBold.Boldweight = (short)FontBoldWeight.BOLD;
                hFontNommalBold.FontName = "Times New Roman";

                //tạo font cho các title end

                //Set style
                var styleTitle = workbook.CreateCellStyle();
                styleTitle.SetFont(hFontTieuDe);
                styleTitle.Alignment = HorizontalAlignment.LEFT;

                //Set styleUnderline
                var styleTitleUnderline = workbook.CreateCellStyle();
                styleTitleUnderline.SetFont(hFontTieuDeUnderline);
                styleTitleUnderline.Alignment = HorizontalAlignment.LEFT;

                //Set style In nghiêng
                var styleTitleItalic = workbook.CreateCellStyle();
                styleTitleItalic.SetFont(hFontTieuDeItalic);
                styleTitleItalic.Alignment = HorizontalAlignment.LEFT;
                //Set style In nghiêng
                var styleTitleNormal = workbook.CreateCellStyle();
                styleTitleNormal.SetFont(hFontTieuDeNormal);
                styleTitleNormal.Alignment = HorizontalAlignment.LEFT;
                var styleTitleBold = workbook.CreateCellStyle();
                styleTitleBold.SetFont(hFontTieuDeBold);
                styleTitleBold.Alignment = HorizontalAlignment.LEFT;


                //Set style Large font
                var styleTitleLarge = workbook.CreateCellStyle();
                styleTitleLarge.SetFont(hFontTieuDeLarge);
                styleTitleLarge.Alignment = HorizontalAlignment.LEFT;

                //style infomation
                var styleInfomation = workbook.CreateCellStyle();
                styleInfomation.SetFont(hFontTT);
                styleInfomation.Alignment = HorizontalAlignment.LEFT;

                //style header
                var styleheadedColumnTable = workbook.CreateCellStyle();
                styleheadedColumnTable.SetFont(hFontNommalUpper);
                styleheadedColumnTable.WrapText = true;
                styleheadedColumnTable.BorderBottom = CellBorderType.THIN;
                styleheadedColumnTable.BorderLeft = CellBorderType.THIN;
                styleheadedColumnTable.BorderRight = CellBorderType.THIN;
                styleheadedColumnTable.BorderTop = CellBorderType.THIN;
                styleheadedColumnTable.VerticalAlignment = VerticalAlignment.CENTER;
                styleheadedColumnTable.Alignment = HorizontalAlignment.CENTER;

                //style sum cell
                var styleCellSumary = workbook.CreateCellStyle();
                styleCellSumary.SetFont(hFontNommalUpper);
                styleCellSumary.WrapText = true;
                styleCellSumary.BorderBottom = CellBorderType.THIN;
                styleCellSumary.BorderLeft = CellBorderType.THIN;
                styleCellSumary.BorderRight = CellBorderType.THIN;
                styleCellSumary.BorderTop = CellBorderType.THIN;
                styleCellSumary.VerticalAlignment = VerticalAlignment.CENTER;
                styleCellSumary.Alignment = HorizontalAlignment.RIGHT;
                var styleCellSumaryLeft = workbook.CreateCellStyle();
                styleCellSumaryLeft.SetFont(hFontNommalUpper);
                styleCellSumaryLeft.WrapText = true;
                styleCellSumaryLeft.BorderBottom = CellBorderType.THIN;
                styleCellSumaryLeft.BorderLeft = CellBorderType.THIN;
                styleCellSumaryLeft.BorderRight = CellBorderType.THIN;
                styleCellSumaryLeft.BorderTop = CellBorderType.THIN;
                styleCellSumaryLeft.VerticalAlignment = VerticalAlignment.CENTER;
                styleCellSumaryLeft.Alignment = HorizontalAlignment.LEFT;

                var styleHeading1 = workbook.CreateCellStyle();
                styleHeading1.SetFont(hFontNommalBold);
                styleHeading1.WrapText = true;
                styleHeading1.BorderBottom = CellBorderType.THIN;
                styleHeading1.BorderLeft = CellBorderType.THIN;
                styleHeading1.BorderRight = CellBorderType.THIN;
                styleHeading1.BorderTop = CellBorderType.THIN;
                styleHeading1.VerticalAlignment = VerticalAlignment.CENTER;
                styleHeading1.Alignment = HorizontalAlignment.LEFT;

                var hStyleConLeft = (HSSFCellStyle)workbook.CreateCellStyle();
                hStyleConLeft.SetFont(hFontNommal);
                hStyleConLeft.VerticalAlignment = VerticalAlignment.TOP;
                hStyleConLeft.Alignment = HorizontalAlignment.LEFT;
                hStyleConLeft.WrapText = true;
                hStyleConLeft.BorderBottom = CellBorderType.THIN;
                hStyleConLeft.BorderLeft = CellBorderType.THIN;
                hStyleConLeft.BorderRight = CellBorderType.THIN;
                hStyleConLeft.BorderTop = CellBorderType.THIN;

                var hStyleConRight = (HSSFCellStyle)workbook.CreateCellStyle();
                hStyleConRight.SetFont(hFontNommal);
                hStyleConRight.VerticalAlignment = VerticalAlignment.TOP;
                hStyleConRight.Alignment = HorizontalAlignment.RIGHT;
                hStyleConRight.BorderBottom = CellBorderType.THIN;
                hStyleConRight.BorderLeft = CellBorderType.THIN;
                hStyleConRight.BorderRight = CellBorderType.THIN;
                hStyleConRight.BorderTop = CellBorderType.THIN;


                var hStyleConCenter = (HSSFCellStyle)workbook.CreateCellStyle();
                hStyleConCenter.SetFont(hFontNommal);
                hStyleConCenter.VerticalAlignment = VerticalAlignment.TOP;
                hStyleConCenter.Alignment = HorizontalAlignment.CENTER;
                hStyleConCenter.BorderBottom = CellBorderType.THIN;
                hStyleConCenter.BorderLeft = CellBorderType.THIN;
                hStyleConCenter.BorderRight = CellBorderType.THIN;
                hStyleConCenter.BorderTop = CellBorderType.THIN;
                //set style end
                #endregion

                //Khai báo row
                Row rowC = null;

                var sheet = workbook.CreateSheet("bangluongnam_" + nam + "_thang_" + thang);
                workbook.ActiveSheetIndex = 1;
                int count = 1;
                string cellTenCty = "Văn phòng cơ quan tổng công ty XDCTGT 6 - công ty cổ phần";
                string cellTitleMain = "BẢNG LƯƠNG NHÂN VIÊN THÁNG " + thang + "/" + nam;
                string cellTitleMainTH = "BẢNG TỔNG HỢP LƯƠNG THÁNG " + thang + "/" + nam;
                //Khai báo row đầu tiên
                int firstRowNumber = 1;
                //Group lại theo phòng ban, mỗi phòng ban là một sheet
                var danhSachBLGroupBys = nhanSuContext.sp_NS_BangLuongNhanVien(maPhongBan, thang, nam, qSearch).GroupBy(s => new { s.boPhanTinhLuong, s.phongBan });
                var idRowStart = firstRowNumber;
                double? TongLuongLe = 0;

                foreach (var item in danhSachBLGroupBys.OrderBy(d => d.Key.boPhanTinhLuong))
                {
                    var titleCellCty = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 0, cellTenCty.ToUpper());
                    titleCellCty.CellStyle = styleTitle;
                    idRowStart++;

                    var titleCellTitleMain = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 5, cellTitleMain.ToUpper());
                    titleCellTitleMain.CellStyle = styleTitle;
                    idRowStart++;

                    var stt = 0;
                    int dem = 0;
                    double? sumLuongLe = 0;
                    string tenPhongBan = string.Empty;
                    if (item.Count() > 0)
                    {
                        tenPhongBan = item.Key.phongBan;
                    }

                    string cellTitlePhongBan = count.ToString() + ". " + tenPhongBan;
                    var titleCellTitlePhongBan = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 0, cellTitlePhongBan);
                    titleCellTitlePhongBan.CellStyle = styleTitleUnderline;

                    count++;
                    idRowStart++;

                    var list1 = new List<string>();

                    list1.Add("STT");
                    list1.Add("Họ và Tên");
                    list1.Add("Lương");// trường này sẽ colspan = 2
                    list1.Add("Phụ cấp lương");
                    list1.Add("Tổng số");
                    list1.Add("Thưởng theo năng suất");

                    list1.Add("Công");
                    list1.Add("Tổng cộng");
                    list1.Add("Công thực tế(Ngày quét, công tác, lễ)");
                    list1.Add("Lương thực tế");
                    //list2.Add("Lễ, phép");
                    //list2.Add("Lương lễ, phép, tăng giờ");
                    list1.Add("Phép");
                    list1.Add("Lương phép");

                    list1.Add("Lễ");
                    list1.Add("Lương lễ");

                    list1.Add("Tăng ca");
                    list1.Add("Lương tăng ca");

                    list1.Add("Công chờ việc");
                    list1.Add("Lương chờ việc");

                    list1.Add("Truy lĩnh");
                    list1.Add("Thực lĩnh");
                    list1.Add("BHXH");
                    list1.Add("BH Y tế");
                    list1.Add("BHTN");
                    list1.Add("BHXH + Y tế + TN");
                    list1.Add("Thuế TNCN");

                    list1.Add("Truy trừ");
                    list1.Add("Đoàn phí");
                    //list2.Add("Đảng phí");
                    // list1.Add("Tiền ăn giữa ca");
                    list1.Add("Tiền điện thoại");
                    list1.Add("Còn lãnh");


                    var list2 = new List<string>();

                    list2.Add("STT");
                    list2.Add("Họ và Tên");
                    list2.Add("Lương");// trường này sẽ colspan = 2
                    list2.Add("Phụ cấp lương");
                    list2.Add("Tổng số");
                    list2.Add("Thưởng theo năng suất");

                    list2.Add("Công");
                    list2.Add("Tổng cộng");
                    list2.Add("Công thực tế(Ngày quét, công tác, lễ)");
                    list2.Add("Lương thực tế");
                    //list2.Add("Lễ, phép");
                    //list2.Add("Lương lễ, phép, tăng giờ");
                    list2.Add("Phép");
                    list2.Add("Lương phép");

                    list2.Add("Lễ");
                    list2.Add("Lương lễ");

                    list2.Add("Tăng ca");
                    list2.Add("Lương tăng ca");

                    list2.Add("Công chờ việc");
                    list2.Add("Lương chờ việc");

                    list2.Add("Truy lĩnh");
                    list2.Add("Thực lĩnh");
                    list2.Add("BHXH");
                    list2.Add("BH Y tế");
                    list2.Add("BHTN");
                    list2.Add("BHXH + Y tế + TN");
                    list2.Add("Thuế TNCN");

                    list2.Add("Truy trừ");
                    list2.Add("Đoàn phí");
                    //list2.Add("Đảng phí");
                    //list2.Add("Tiền ăn giữa ca");
                    list2.Add("Tiền điện thoại");
                    list2.Add("Còn lãnh");

                    var list3 = new List<string>();

                    list3.Add("1");
                    list3.Add("2");
                    list3.Add("3");// trường này sẽ colspan = 2
                    list3.Add("4");
                    list3.Add("");
                    list3.Add("5");

                    list3.Add("");
                    list3.Add("");
                    list3.Add("6");
                    list3.Add("7");
                    //phép
                    list3.Add("8");
                    list3.Add("9");
                    //le
                    list3.Add("");
                    list3.Add("");
                    //tang ca
                    list3.Add("");
                    list3.Add("");
                    //cho viec
                    list3.Add("10");
                    list3.Add("");

                    list3.Add("11");
                    list3.Add("12");
                    list3.Add("");
                    list3.Add("");
                    list3.Add("");
                    list3.Add("13");
                    list3.Add("14");

                    list3.Add("15");
                    list3.Add("16");
                    //list3.Add("17");
                    list3.Add("17");

                    list3.Add("18=12-13-14-15-16+17");

                    var headerRow = sheet.CreateRow(idRowStart);
                    int rowend = idRowStart;
                    ReportHelperExcel.CreateHeaderRow(headerRow, 0, styleheadedColumnTable, list1);
                    idRowStart++;
                    var headerRow1 = sheet.CreateRow(idRowStart);
                    ReportHelperExcel.CreateHeaderRow(headerRow1, 0, styleheadedColumnTable, list2);
                    idRowStart++;
                    var headerRow11 = sheet.CreateRow(idRowStart);
                    ReportHelperExcel.CreateHeaderRow(headerRow11, 0, styleheadedColumnTable, list3);
                    //sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 22, 22));
                    //sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 21, 21));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 29, 29));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 28, 28));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 27, 27));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 26, 26));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 25, 25));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 24, 24));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 23, 23));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 22, 22));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 21, 21));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 20, 20));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 19, 19));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 18, 18));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 17, 17));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 16, 16));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 15, 15));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 14, 14));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 13, 13));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 12, 12));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 11, 11));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 10, 10));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 9, 9));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 8, 8));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 7, 7));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 6, 6));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 5, 5));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 4, 4));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow.RowNum, 2, 3));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 1, 1));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 0, 0));



                    sheet.SetColumnWidth(0, 8 * 210);
                    sheet.SetColumnWidth(1, 30 * 210);
                    sheet.SetColumnWidth(2, 15 * 210);
                    sheet.SetColumnWidth(3, 15 * 210);
                    sheet.SetColumnWidth(4, 15 * 210);
                    sheet.SetColumnWidth(5, 15 * 210);
                    sheet.SetColumnWidth(6, 15 * 210);
                    sheet.SetColumnWidth(7, 15 * 210);
                    sheet.SetColumnWidth(8, 15 * 210);
                    sheet.SetColumnWidth(9, 15 * 210);
                    sheet.SetColumnWidth(10, 15 * 210);
                    sheet.SetColumnWidth(11, 15 * 210);

                    sheet.SetColumnWidth(12, 15 * 210);
                    sheet.SetColumnWidth(13, 15 * 210);
                    sheet.SetColumnWidth(14, 15 * 210);
                    sheet.SetColumnWidth(15, 15 * 210);
                    sheet.SetColumnWidth(16, 15 * 210);
                    sheet.SetColumnWidth(17, 15 * 210);
                    sheet.SetColumnWidth(18, 15 * 210);
                    sheet.SetColumnWidth(19, 15 * 210);
                    sheet.SetColumnWidth(20, 15 * 210);
                    sheet.SetColumnWidth(21, 15 * 210);
                    sheet.SetColumnWidth(22, 15 * 210);
                    sheet.SetColumnWidth(23, 15 * 210);
                    sheet.SetColumnWidth(24, 15 * 210);
                    sheet.SetColumnWidth(25, 15 * 210);
                    sheet.SetColumnWidth(26, 25 * 210);
                    sheet.SetColumnWidth(27, 25 * 210);
                    sheet.SetColumnWidth(28, 25 * 210);
                    sheet.SetColumnWidth(29, 25 * 210);

                    //double? sumLuongLePB = 0;
                    //sumLuongLePB = ((item.Sum(s => s.tongLuong ?? 0) * (item.Sum(s => s.soNgayNghiLe ?? 0) + (item.Sum(s => s.soNgayNghiPhep ?? 0)))) / item.Sum(s => s.ngayCongChuan ?? 0));
                    //Giai đoạn
                    double tongThucLinh = 0;

                    foreach (var item1 in item.OrderBy(d => d.hoTen))
                    {
                        dem = 0;
                        //double? luongNghiLe = ((item1.tongLuong ?? 0) * ((item1.soNgayNghiLe ?? 0) + (item1.soNgayNghiPhep ?? 0))) / item1.ngayCongChuan;
                        //double? luongNghiPhep = ((item1.tongLuong ?? 0) * ((item1.soNgayNghiLe ?? 0) + (item1.soNgayNghiPhep ?? 0))) / item1.ngayCongChuan;
                        //sumLuongLe += luongNghiLePhep;
                        stt++;
                        idRowStart++;

                        rowC = sheet.CreateRow(idRowStart);
                        ReportHelperExcel.SetAlignment(rowC, dem++, stt.ToString(), hStyleConCenter);
                        ReportHelperExcel.SetAlignment(rowC, dem++, item1.hoTen, hStyleConLeft);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.luongThoaThuan), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.phuCapLuong), hStyleConRight);
                        // Tong so
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", ((item1.luongThoaThuan ?? 0) + (item1.phuCapLuong ?? 0))), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.khoanBoSungLuong), hStyleConRight);
                        // Cong chuan 
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", (item1.ngayCongChuan ?? 0)), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.tongLuong), hStyleConRight);
                        // End cong chuan
                        // cong di lam
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", (item1.soNgayCongTac ?? 0) + (item1.soNgayQuet ?? 0) + (item1.soNgayNghiLe ?? 0)
                                + (item1.soNgayPhepLuyKeThangTruoc ?? 0) + (item1.soNgayLuyKeThangTruoc ?? 0)), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.luongThucTe), hStyleConRight);

                        //Phép
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", (item1.soNgayNghiPhep ?? 0)), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.luongNghiPhep), hStyleConRight);

                        //Lễ
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", (item1.soNgayNghiLe ?? 0)), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.luongNghiLe), hStyleConRight);

                        //tăng ca
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", (item1.soNgayTangCaThuong ?? 0) + (item1.soNgayTangCaChuNhat ?? 0) + (item1.soNgayTangCaNgayLe ?? 0)), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.luongTangCa), hStyleConRight);

                        //Công chờ việc
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.soNgayCongChoViec), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.luongChoViec), hStyleConRight);


                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.truyLanh), hStyleConRight);
                        //Löông thöïc teá,Löông pheùp,Truy lĩnh; tiền tăng giờ; löông chôø vieäc, 
                        double thucLinh = (item1.luongThucTe ?? 0) + (item1.luongNghiPhep ?? 0) + (item1.luongTangCa ?? 0) + (item1.luongChoViec ?? 0) + (item1.truyLanh ?? 0);

                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", thucLinh), hStyleConRight);


                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", (item1.ngayCongChuan - 14 >= item1.tongNgayCong) ? 0 : item1.baoHiemXH), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", (item1.ngayCongChuan - 14 >= item1.tongNgayCong) ? 0 : item1.baoHiemYTe), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", (item1.ngayCongChuan - 14 >= item1.tongNgayCong) ? 0 : item1.baoHiemTN), hStyleConRight);

                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", (item1.ngayCongChuan - 14 >= item1.tongNgayCong) ? 0 : (item1.baoHiemTN + item1.baoHiemXH + item1.baoHiemYTe)), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.thue), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.truyThu), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.doanPhi), hStyleConRight);
                        //ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.dangPhi), hStyleConRight);
                        //ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.phuCapTienAn), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.phuCapDienThoai), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.thucLanh), hStyleConRight);

                        tongThucLinh += thucLinh;
                    }

                    int demT = 0;
                    idRowStart++;
                    Row rowT = sheet.CreateRow(idRowStart);
                    ReportHelperExcel.SetAlignment(rowT, demT++, (stt).ToString(), styleheadedColumnTable);
                    ReportHelperExcel.SetAlignment(rowT, demT++, "", styleCellSumary);

                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.luongThoaThuan)), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.phuCapLuong)), styleCellSumary);
                    // Sum tong so
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", ((item.Sum(s => s.luongThoaThuan) ?? 0) + (item.Sum(s => s.phuCapLuong) ?? 0))), styleCellSumary);

                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", (item.Sum(s => s.khoanBoSungLuong) ?? 0)), styleCellSumary);
                    // Cong chuan 
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0.##}", item.Sum(s => s.ngayCongChuan ?? 0)), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.tongLuong)), styleCellSumary);
                    // End cong chuan
                    // cong di lam
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0.##}", item.Sum(s => ((s.soNgayCongTac ?? 0) + (s.soNgayQuet ?? 0) + (s.soNgayNghiLe ?? 0)
                                + (s.soNgayPhepLuyKeThangTruoc ?? 0) + (s.soNgayLuyKeThangTruoc ?? 0)))), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => (s.luongThucTe ?? 0))), styleCellSumary);
                    // End cong di lam

                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0.##}", item.Sum(s => ((s.soNgayNghiPhep ?? 0)))), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.luongNghiPhep)), styleCellSumary);

                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0.##}", item.Sum(s => ((s.soNgayNghiLe ?? 0)))), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.luongNghiLe)), styleCellSumary);

                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0.##}", item.Sum(s => ((s.soNgayTangCaThuong ?? 0) + (s.soNgayTangCaChuNhat ?? 0) + (s.soNgayTangCaNgayLe ?? 0)))), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.luongTangCa)), styleCellSumary);

                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.soNgayCongChoViec)), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.luongChoViec)), styleCellSumary);

                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.truyLanh)), styleCellSumary);

                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", tongThucLinh), styleCellSumary);

                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => (s.baoHiemXH))), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => (s.baoHiemYTe))), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => (s.baoHiemTN))), styleCellSumary);

                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.baoHiem)), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.thue)), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.truyThu)), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.doanPhi)), styleCellSumary);
                    //ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.dangPhi)), styleCellSumary);
                    //ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.phuCapTienAn)), styleCellSumary);

                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.phuCapDienThoai)), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.thucLanh)), styleCellSumary);

                    idRowStart = idRowStart + 2;
                    string cellFooterSoTien = "(" + CharacterHelper.DocTienBangChu((decimal)item.Sum(s => s.thucLanh), string.Empty) + ")";
                    var titleCellFooterSoTien = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 12, cellFooterSoTien);
                    titleCellFooterSoTien.CellStyle = styleTitleItalic;

                    idRowStart = idRowStart + 1;
                    string cellFooterNgayLap = "Tp.Hồ Chí Minh, ngày   tháng  năm " + nam;
                    var titleCellFooterNgayLap = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 13, cellFooterNgayLap);
                    titleCellFooterNgayLap.CellStyle = styleTitleItalic;

                    idRowStart = idRowStart + 1;
                    string cellFooterPTC = "PHÒNG TÀI CHÍNH KẾ TOÁN";
                    var titleCellFooterPTC = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 1, cellFooterPTC);
                    titleCellFooterPTC.CellStyle = styleTitleNormal;

                    string cellFooterKT = "TRƯỞNG PHÒNG HÀNH CHÍNH - NHÂN SỰ";
                    var titleCellFooterKT = HSSFCellUtil.CreateCell(sheet.GetRow(idRowStart), 13, cellFooterKT);
                    titleCellFooterKT.CellStyle = styleTitleNormal;

                    //string cellFooterTGD = "TỔNG GIÁM ĐỐC";
                    //var titleCellFooterTGD = HSSFCellUtil.CreateCell(sheet.GetRow(idRowStart), 14, cellFooterTGD);
                    //titleCellFooterTGD.CellStyle = styleTitle;
                    idRowStart = idRowStart + 2;
                }

                // Sum tong theo bo phan tinh luong
                var danhSachBLGroupBysTong = nhanSuContext.sp_NS_BangLuongNhanVien(maPhongBan, thang, nam, qSearch).GroupBy(s => new { s.boPhanTinhLuong });
                var flashHeadTong = 0;
                double tongThucLinhBoPhan = 0;

                foreach (var item in danhSachBLGroupBysTong)
                {
                    #region định nghĩa
                    var stt = 0;
                    int dem = 0;
                    //double? sumLuongLePB = 0;
                    //sumLuongLePB = ((item.Sum(s => s.tongLuong ?? 0) * (item.Sum(s => s.soNgayNghiLe ?? 0) + (item.Sum(s => s.soNgayNghiPhep ?? 0)))) / item.Sum(s => s.ngayCongChuan ?? 0));
                    count++;
                    if (flashHeadTong == 0)
                    {
                        var titleCellCty = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 0, cellTenCty.ToUpper());
                        titleCellCty.CellStyle = styleTitle;
                        idRowStart++;

                        var titleCellTitleMain = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 5, cellTitleMainTH.ToUpper());
                        titleCellTitleMain.CellStyle = styleTitle;
                        idRowStart++;
                        var list1 = new List<string>();

                        list1.Add("TS  NV");
                        list1.Add("Đơn vị");
                        list1.Add("Tổng mức lương");// trường này sẽ colspan = 2
                        list1.Add("");
                        list1.Add("Tổng số");
                        list1.Add("Thưởng theo năng suất");

                        list1.Add("Công");
                        list1.Add("Lương tháng");
                        list1.Add("Công thực tế");
                        list1.Add("Lương thực tế");


                        list1.Add("Phép");
                        list1.Add("Lương phép");

                        list1.Add("Lễ");
                        list1.Add("Lương lễ");

                        list1.Add("Tăng ca");
                        list1.Add("Lương tăng ca");

                        list1.Add("Công chờ việc");
                        list1.Add("Lương chờ việc");

                        list1.Add("Truy lĩnh");
                        list1.Add("Thực lĩnh");
                        list1.Add("BHXH + Y tế + TN");
                        list1.Add("Thuế TNCN");
                        list1.Add("Truy trừ");
                        list1.Add("Đoàn phí");
                        //list1.Add("Đảng phí");
                        //list1.Add("Tiền ăn giữa ca");
                        list1.Add("Tiền điện thoại");
                        list1.Add("Còn lãnh");


                        var list2 = new List<string>();

                        list2.Add("TS  NV");
                        list2.Add("Đơn vị");
                        list2.Add("Lương");// trường này sẽ colspan = 2
                        list2.Add("Phụ cấp lương");
                        list2.Add("Tổng số");
                        list2.Add("Thưởng theo năng suất");

                        list2.Add("Công");
                        list2.Add("Lương tháng");
                        list2.Add("Công thực tế");
                        list2.Add("Lương thực tế");

                        list2.Add("Phép");
                        list2.Add("Lương phép");

                        list2.Add("Lễ");
                        list2.Add("Lương lễ");

                        list2.Add("Tăng ca");
                        list2.Add("Lương tăng ca");

                        list2.Add("Công chờ việc");
                        list2.Add("Lương chờ việc");


                        list2.Add("Truy lĩnh");
                        list2.Add("Thực lĩnh");
                        list2.Add("BHXH + Y tế + TN");
                        list2.Add("Thuế TNCN");
                        list2.Add("Truy trừ");
                        list2.Add("Đoàn phí");
                        //list2.Add("Đảng phí");
                        //list2.Add("Tiền ăn giữa ca");
                        list2.Add("Tiền điện thoại");
                        list2.Add("Còn lãnh");


                        var headerRow = sheet.CreateRow(idRowStart);
                        int rowend = idRowStart;
                        ReportHelperExcel.CreateHeaderRow(headerRow, 0, styleheadedColumnTable, list1);
                        idRowStart++;
                        var headerRow1 = sheet.CreateRow(idRowStart);
                        ReportHelperExcel.CreateHeaderRow(headerRow1, 0, styleheadedColumnTable, list2);

                        //sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 22, 22));
                        //sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 21, 21));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 26, 26));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 25, 25));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 24, 24));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 23, 23));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 22, 22));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 21, 21));

                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 20, 20));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 19, 19));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 18, 18));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 17, 17));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 16, 16));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 15, 15));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 14, 14));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 13, 13));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 12, 12));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 11, 11));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 10, 10));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 9, 9));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 8, 8));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 7, 7));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 6, 6));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 5, 5));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 4, 4));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow.RowNum, 2, 3));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 1, 1));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 0, 0));



                        sheet.SetColumnWidth(0, 8 * 210);
                        sheet.SetColumnWidth(1, 30 * 210);
                        sheet.SetColumnWidth(2, 15 * 210);
                        sheet.SetColumnWidth(3, 15 * 210);
                        sheet.SetColumnWidth(4, 15 * 210);
                        sheet.SetColumnWidth(5, 15 * 210);
                        sheet.SetColumnWidth(6, 15 * 210);
                        sheet.SetColumnWidth(7, 15 * 210);
                        sheet.SetColumnWidth(8, 15 * 210);
                        sheet.SetColumnWidth(9, 15 * 210);
                        sheet.SetColumnWidth(10, 15 * 210);
                        sheet.SetColumnWidth(11, 25 * 210);

                        sheet.SetColumnWidth(12, 15 * 210);
                        sheet.SetColumnWidth(13, 15 * 210);
                        sheet.SetColumnWidth(14, 25 * 210);
                        sheet.SetColumnWidth(15, 25 * 210);
                        sheet.SetColumnWidth(16, 15 * 210);
                        sheet.SetColumnWidth(17, 25 * 210);
                        sheet.SetColumnWidth(18, 25 * 210);
                        sheet.SetColumnWidth(19, 25 * 210);
                        sheet.SetColumnWidth(20, 25 * 210);
                        sheet.SetColumnWidth(21, 25 * 210);
                        sheet.SetColumnWidth(22, 25 * 210);
                        sheet.SetColumnWidth(23, 25 * 210);
                        sheet.SetColumnWidth(24, 25 * 210);
                        sheet.SetColumnWidth(25, 25 * 210);
                        sheet.SetColumnWidth(26, 25 * 210);

                    }
                    flashHeadTong = 1;

                    string tenPhongBan = string.Empty;
                    if (item.Count() > 0)
                    {
                        tenPhongBan = item.Key.boPhanTinhLuong;
                    }
                    int demT = 0;
                    idRowStart++;

                    #endregion

                    Row rowT = sheet.CreateRow(idRowStart);
                    ReportHelperExcel.SetAlignment(rowT, demT++, item.Count().ToString(), styleheadedColumnTable);
                    ReportHelperExcel.SetAlignment(rowT, demT++, tenPhongBan, styleCellSumaryLeft);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.luongThang)), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.phuCapLuong)), hStyleConRight);
                    // Sum tong so
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", ((item.Sum(s => s.luongThang) ?? 0) + (item.Sum(s => s.phuCapLuong) ?? 0))), hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", (item.Sum(s => s.khoanBoSungLuong) ?? 0)), hStyleConRight);
                    // Cong chuan 
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0.##}", item.Sum(s => s.ngayCongChuan ?? 0)), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.tongLuong)), hStyleConRight);
                    // End cong chuan
                    // cong di lam
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0.##}", item.Sum(s => ((s.soNgayCongTac ?? 0) + (s.soNgayQuet ?? 0) + (s.soNgayNghiLe ?? 0)
                            + (s.soNgayPhepLuyKeThangTruoc ?? 0) + (s.soNgayLuyKeThangTruoc ?? 0)))), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => (s.luongThucTe ?? 0))), hStyleConRight);
                    // End cong di lam
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0.##}", item.Sum(s => ((s.soNgayNghiPhep ?? 0)))), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.luongNghiPhep)), hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0.##}", item.Sum(s => ((s.soNgayNghiLe ?? 0)))), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.luongNghiLe)), hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0.##}", item.Sum(s => ((s.soNgayTangCaThuong ?? 0) + (s.soNgayTangCaChuNhat ?? 0) + (s.soNgayTangCaNgayLe ?? 0)))), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.luongTangCa)), hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0.##}", item.Sum(s => ((s.soNgayCongChoViec ?? 0)))), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.luongChoViec)), hStyleConRight);


                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.truyLanh)), hStyleConRight);

                    double thucLinh = item.Sum(s => (s.luongThucTe ?? 0) + (s.luongNghiPhep ?? 0) + (s.luongTangCa ?? 0) + (s.luongChoViec ?? 0) + (s.truyLanh ?? 0));

                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", thucLinh), hStyleConRight);

                    //ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => (s.baoHiemXH))), hStyleConRight);
                    //ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => (s.baoHiemYTe))), hStyleConRight);
                    //ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => (s.baoHiemTN))), hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => (s.baoHiemTN + s.baoHiemXH + s.baoHiemYTe))), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.thue)), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.truyThu)), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.doanPhi)), styleCellSumary);
                    //ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.dangPhi)), styleCellSumary);
                    //ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.phuCapTienAn)), hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.phuCapDienThoai)), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.thucLanh)), styleCellSumary);

                    tongThucLinhBoPhan += thucLinh;

                }
                var danhSachBLGroupBysTongSum = nhanSuContext.sp_NS_BangLuongNhanVien(maPhongBan, thang, nam, qSearch).ToList();

                int demTong = 0;
                idRowStart++;
                Row rowTong = sheet.CreateRow(idRowStart);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0.##}", danhSachBLGroupBysTongSum.Count()), styleheadedColumnTable);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, "Tổng cộng", styleCellSumaryLeft);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.luongThang)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.phuCapLuong)), styleCellSumary);
                // Sum tong so
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", (danhSachBLGroupBysTongSum.Sum(s => s.luongThang + s.phuCapLuong) ?? 0)), styleCellSumary);

                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", (danhSachBLGroupBysTongSum.Sum(s => s.khoanBoSungLuong) ?? 0)), styleCellSumary);
                // Cong chuan 
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0.##}", danhSachBLGroupBysTongSum.Sum(s => s.ngayCongChuan ?? 0)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.tongLuong)), styleCellSumary);
                // End cong chuan
                // cong di lam
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0.##}", danhSachBLGroupBysTongSum.Sum(s => ((s.soNgayCongTac ?? 0) + (s.soNgayQuet ?? 0) + (s.soNgayNghiLe ?? 0)
                        + (s.soNgayPhepLuyKeThangTruoc ?? 0) + (s.soNgayLuyKeThangTruoc ?? 0)))), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.luongThucTe)), styleCellSumary);
                // End cong di lam


                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0.##}", danhSachBLGroupBysTongSum.Sum(s => ((s.soNgayNghiPhep ?? 0)))), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.luongNghiPhep)), styleCellSumary);

                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0.##}", danhSachBLGroupBysTongSum.Sum(s => ((s.soNgayNghiLe ?? 0)))), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.luongNghiLe)), styleCellSumary);

                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0.##}", danhSachBLGroupBysTongSum.Sum(s => ((s.soNgayTangCaThuong ?? 0) + (s.soNgayTangCaChuNhat ?? 0) + (s.soNgayTangCaNgayLe ?? 0)))), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.luongTangCa)), styleCellSumary);

                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0.##}", danhSachBLGroupBysTongSum.Sum(s => ((s.soNgayCongChoViec ?? 0)))), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.luongChoViec)), styleCellSumary);

                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.truyLanh)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", tongThucLinhBoPhan), styleCellSumary);

                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => (s.baoHiemTN + s.baoHiemXH + s.baoHiemYTe))), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.thue)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.truyThu)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.doanPhi)), styleCellSumary);
                //ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.dangPhi)), styleCellSumary);
                //ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.phuCapTienAn)), styleCellSumary);

                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.phuCapDienThoai)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.thucLanh)), styleCellSumary);
                idRowStart = idRowStart + 2;
                string cellFooterSoTienTong = "(" + CharacterHelper.DocTienBangChu((decimal)danhSachBLGroupBysTongSum.Sum(s => s.thucLanh), string.Empty) + ")";
                var titleCellFooterSoTienTong = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 12, cellFooterSoTienTong);
                titleCellFooterSoTienTong.CellStyle = styleTitleItalic;

                idRowStart = idRowStart + 1;
                string cellFooterNgayLapTong = "Tp.Hồ Chí Minh, ngày   tháng  năm " + nam;
                var titleCellFooterNgayLapTong = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 13, cellFooterNgayLapTong);
                titleCellFooterNgayLapTong.CellStyle = styleTitleItalic;

                idRowStart = idRowStart + 1;
                string cellFooterPTCTong = "TRƯỞNG PHÒNG HÀNH CHÍNH - NHÂN SỰ";
                var titleCellFooterPTCTong = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 1, cellFooterPTCTong);
                titleCellFooterPTCTong.CellStyle = styleTitleBold;

                string cellFooterKTTong = "PHÒNG TÀI CHÍNH - KẾ TOÁN";
                var titleCellFooterKTTong = HSSFCellUtil.CreateCell(sheet.GetRow(idRowStart), 7, cellFooterKTTong);
                titleCellFooterKTTong.CellStyle = styleTitleBold;
                string cellFooterKTTongTGD = "TỔNG GIÁM ĐỐC";
                var titleCellFooterKTTongTGD = HSSFCellUtil.CreateCell(sheet.GetRow(idRowStart), 13, cellFooterKTTongTGD);
                titleCellFooterKTTongTGD.CellStyle = styleTitleBold;
                // End sum 


                var stream = new MemoryStream();
                workbook.Write(stream);

                Response.ContentType = "application/vnd.ms-excel";
                Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));
                Response.Clear();

                Response.BinaryWrite(stream.GetBuffer());
                Response.End();
            }
            catch
            {

            }
        }
        #endregion



        #region Xuat File Bang Luong Nhan Vien
        public void XuatFileBLNV_TrinhKy(int thang, int nam)
        {
            try
            {
                string maPhongBan = "";
                string qSearch = "";
                var filename = "";
                var virtualPath = HttpRuntime.AppDomainAppVirtualPath;

                var fileStream = new FileStream(System.Web.HttpContext.Current.Server.MapPath(virtualPath + @"\Content\Report\ReportTemplatecc.xls"), FileMode.Open, FileAccess.Read);

                var workbook = new HSSFWorkbook(fileStream, true);
                filename += "BangLuongNV_" + nam + "_" + thang + ".xls";

                #region format style excel cell
                /*style title start*/
                //tạo font cho các title
                //font tiêu đề 
                HSSFFont hFontTieuDe = (HSSFFont)workbook.CreateFont();
                hFontTieuDe.FontHeightInPoints = 11;
                hFontTieuDe.Boldweight = 100 * 10;
                hFontTieuDe.FontName = "Times New Roman";
                //hFontTieuDe.Color = HSSFColor.BLUE.index;

                //font tiêu đề 
                HSSFFont hFontTieuDeUnderline = (HSSFFont)workbook.CreateFont();
                hFontTieuDeUnderline.FontHeightInPoints = 11;
                hFontTieuDeUnderline.Boldweight = 100 * 10;
                hFontTieuDeUnderline.FontName = "Times New Roman";
                hFontTieuDeUnderline.Underline = 1;
                //hFontTieuDe.Color = HSSFColor.BLUE.index;


                HSSFFont hFontTieuDeItalic = (HSSFFont)workbook.CreateFont();
                hFontTieuDeItalic.FontHeightInPoints = 11;
                //hFontTieuDeItalic.Boldweight = 100 * 10;
                hFontTieuDeItalic.FontName = "Times New Roman";
                hFontTieuDeItalic.IsItalic = true;

                //hFontTieuDe.Color = HSSFColor.BLUE.index;

                HSSFFont hFontTieuDeNormal = (HSSFFont)workbook.CreateFont();
                hFontTieuDeNormal.FontHeightInPoints = 11;
                //hFontTieuDeItalic.Boldweight = 100 * 10;
                hFontTieuDeNormal.FontName = "Times New Roman";
                hFontTieuDeNormal.IsItalic = false;
                HSSFFont hFontTieuDeBold = (HSSFFont)workbook.CreateFont();
                hFontTieuDeBold.FontHeightInPoints = 11;
                hFontTieuDeItalic.Boldweight = (short)FontBoldWeight.BOLD;
                hFontTieuDeBold.FontName = "Times New Roman";
                hFontTieuDeBold.IsItalic = false;


                HSSFFont hFontTieuDeLarge = (HSSFFont)workbook.CreateFont();
                hFontTieuDeLarge.FontHeightInPoints = 16;
                hFontTieuDeLarge.Boldweight = 100 * 10;
                hFontTieuDeLarge.FontName = "Times New Roman";
                //hFontTieuDeLarge.Color = HSSFColor.BLUE.index;

                //font tiêu đề 
                HSSFFont hFontTongGiaTriHT = (HSSFFont)workbook.CreateFont();
                hFontTongGiaTriHT.FontHeightInPoints = 11;
                hFontTongGiaTriHT.Boldweight = (short)FontBoldWeight.BOLD;
                hFontTongGiaTriHT.FontName = "Times New Roman";
                hFontTongGiaTriHT.Color = HSSFColor.BLACK.index;

                //font thông tin bảng tính
                HSSFFont hFontTT = (HSSFFont)workbook.CreateFont();
                hFontTT.IsItalic = true;
                hFontTT.Boldweight = (short)FontBoldWeight.BOLD;
                hFontTT.Color = HSSFColor.BLACK.index;
                hFontTT.FontName = "Times New Roman";
                hFontTieuDe.FontHeightInPoints = 11;

                //font chứ hoa đậm
                HSSFFont hFontNommalUpper = (HSSFFont)workbook.CreateFont();
                hFontNommalUpper.Boldweight = (short)FontBoldWeight.BOLD;
                hFontNommalUpper.Color = HSSFColor.BLACK.index;
                hFontNommalUpper.FontName = "Times New Roman";

                //font chữ bình thường
                HSSFFont hFontNommal = (HSSFFont)workbook.CreateFont();
                hFontNommal.Color = HSSFColor.BLACK.index;
                hFontNommal.FontName = "Times New Roman";

                //font chữ bình thường đậm
                HSSFFont hFontNommalBold = (HSSFFont)workbook.CreateFont();
                hFontNommalBold.Color = HSSFColor.BLACK.index;
                hFontNommalBold.Boldweight = (short)FontBoldWeight.BOLD;
                hFontNommalBold.FontName = "Times New Roman";

                //tạo font cho các title end

                //Set style
                var styleTitle = workbook.CreateCellStyle();
                styleTitle.SetFont(hFontTieuDe);
                styleTitle.Alignment = HorizontalAlignment.LEFT;

                //Set styleUnderline
                var styleTitleUnderline = workbook.CreateCellStyle();
                styleTitleUnderline.SetFont(hFontTieuDeUnderline);
                styleTitleUnderline.Alignment = HorizontalAlignment.LEFT;

                //Set style In nghiêng
                var styleTitleItalic = workbook.CreateCellStyle();
                styleTitleItalic.SetFont(hFontTieuDeItalic);
                styleTitleItalic.Alignment = HorizontalAlignment.LEFT;
                //Set style In nghiêng
                var styleTitleNormal = workbook.CreateCellStyle();
                styleTitleNormal.SetFont(hFontTieuDeNormal);
                styleTitleNormal.Alignment = HorizontalAlignment.LEFT;
                var styleTitleBold = workbook.CreateCellStyle();
                styleTitleBold.SetFont(hFontTieuDeBold);
                styleTitleBold.Alignment = HorizontalAlignment.LEFT;


                //Set style Large font
                var styleTitleLarge = workbook.CreateCellStyle();
                styleTitleLarge.SetFont(hFontTieuDeLarge);
                styleTitleLarge.Alignment = HorizontalAlignment.LEFT;

                //style infomation
                var styleInfomation = workbook.CreateCellStyle();
                styleInfomation.SetFont(hFontTT);
                styleInfomation.Alignment = HorizontalAlignment.LEFT;

                //style header
                var styleheadedColumnTable = workbook.CreateCellStyle();
                styleheadedColumnTable.SetFont(hFontNommalUpper);
                styleheadedColumnTable.WrapText = true;
                styleheadedColumnTable.BorderBottom = CellBorderType.THIN;
                styleheadedColumnTable.BorderLeft = CellBorderType.THIN;
                styleheadedColumnTable.BorderRight = CellBorderType.THIN;
                styleheadedColumnTable.BorderTop = CellBorderType.THIN;
                styleheadedColumnTable.VerticalAlignment = VerticalAlignment.CENTER;
                styleheadedColumnTable.Alignment = HorizontalAlignment.CENTER;

                //style sum cell
                var styleCellSumary = workbook.CreateCellStyle();
                styleCellSumary.SetFont(hFontNommalUpper);
                styleCellSumary.WrapText = true;
                styleCellSumary.BorderBottom = CellBorderType.THIN;
                styleCellSumary.BorderLeft = CellBorderType.THIN;
                styleCellSumary.BorderRight = CellBorderType.THIN;
                styleCellSumary.BorderTop = CellBorderType.THIN;
                styleCellSumary.VerticalAlignment = VerticalAlignment.CENTER;
                styleCellSumary.Alignment = HorizontalAlignment.RIGHT;
                var styleCellSumaryLeft = workbook.CreateCellStyle();
                styleCellSumaryLeft.SetFont(hFontNommalUpper);
                styleCellSumaryLeft.WrapText = true;
                styleCellSumaryLeft.BorderBottom = CellBorderType.THIN;
                styleCellSumaryLeft.BorderLeft = CellBorderType.THIN;
                styleCellSumaryLeft.BorderRight = CellBorderType.THIN;
                styleCellSumaryLeft.BorderTop = CellBorderType.THIN;
                styleCellSumaryLeft.VerticalAlignment = VerticalAlignment.CENTER;
                styleCellSumaryLeft.Alignment = HorizontalAlignment.LEFT;

                var styleHeading1 = workbook.CreateCellStyle();
                styleHeading1.SetFont(hFontNommalBold);
                styleHeading1.WrapText = true;
                styleHeading1.BorderBottom = CellBorderType.THIN;
                styleHeading1.BorderLeft = CellBorderType.THIN;
                styleHeading1.BorderRight = CellBorderType.THIN;
                styleHeading1.BorderTop = CellBorderType.THIN;
                styleHeading1.VerticalAlignment = VerticalAlignment.CENTER;
                styleHeading1.Alignment = HorizontalAlignment.LEFT;

                var hStyleConLeft = (HSSFCellStyle)workbook.CreateCellStyle();
                hStyleConLeft.SetFont(hFontNommal);
                hStyleConLeft.VerticalAlignment = VerticalAlignment.TOP;
                hStyleConLeft.Alignment = HorizontalAlignment.LEFT;
                hStyleConLeft.WrapText = true;
                hStyleConLeft.BorderBottom = CellBorderType.THIN;
                hStyleConLeft.BorderLeft = CellBorderType.THIN;
                hStyleConLeft.BorderRight = CellBorderType.THIN;
                hStyleConLeft.BorderTop = CellBorderType.THIN;

                var hStyleConRight = (HSSFCellStyle)workbook.CreateCellStyle();
                hStyleConRight.SetFont(hFontNommal);
                hStyleConRight.VerticalAlignment = VerticalAlignment.TOP;
                hStyleConRight.Alignment = HorizontalAlignment.RIGHT;
                hStyleConRight.BorderBottom = CellBorderType.THIN;
                hStyleConRight.BorderLeft = CellBorderType.THIN;
                hStyleConRight.BorderRight = CellBorderType.THIN;
                hStyleConRight.BorderTop = CellBorderType.THIN;


                var hStyleConCenter = (HSSFCellStyle)workbook.CreateCellStyle();
                hStyleConCenter.SetFont(hFontNommal);
                hStyleConCenter.VerticalAlignment = VerticalAlignment.TOP;
                hStyleConCenter.Alignment = HorizontalAlignment.CENTER;
                hStyleConCenter.BorderBottom = CellBorderType.THIN;
                hStyleConCenter.BorderLeft = CellBorderType.THIN;
                hStyleConCenter.BorderRight = CellBorderType.THIN;
                hStyleConCenter.BorderTop = CellBorderType.THIN;
                //set style end
                #endregion

                //Khai báo row
                Row rowC = null;

                var sheet = workbook.CreateSheet("bangluongnam_" + nam + "_thang_" + thang);
                workbook.ActiveSheetIndex = 1;
                int count = 1;
                string cellTenCty = "Văn phòng cơ quan tổng công ty XDCTGT 6 - công ty cổ phần";
                string cellTitleMain = "BẢNG LƯƠNG NHÂN VIÊN THÁNG " + thang + "/" + nam;
                string cellTitleMainTH = "BẢNG TỔNG HỢP LƯƠNG THÁNG " + thang + "/" + nam;
                //Khai báo row đầu tiên
                int firstRowNumber = 1;
                //Group lại theo phòng ban, mỗi phòng ban là một sheet
                var danhSachBLGroupBys = nhanSuContext.sp_NS_BangLuongNhanVien(maPhongBan, thang, nam, qSearch).GroupBy(s => new { s.boPhanTinhLuong, s.phongBan });
                var idRowStart = firstRowNumber;

                foreach (var item in danhSachBLGroupBys.OrderBy(d => d.Key.boPhanTinhLuong))
                {
                    var titleCellCty = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 0, cellTenCty.ToUpper());
                    titleCellCty.CellStyle = styleTitle;
                    idRowStart++;

                    var titleCellTitleMain = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 5, cellTitleMain.ToUpper());
                    titleCellTitleMain.CellStyle = styleTitle;
                    idRowStart++;

                    var stt = 0;
                    int dem = 0;

                    string tenPhongBan = string.Empty;
                    if (item.Count() > 0)
                    {
                        tenPhongBan = item.Key.phongBan;
                    }

                    string cellTitlePhongBan = count.ToString() + ". " + tenPhongBan;
                    var titleCellTitlePhongBan = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 0, cellTitlePhongBan);
                    titleCellTitlePhongBan.CellStyle = styleTitleUnderline;

                    count++;
                    idRowStart++;

                    var list1 = new List<string>();

                    list1.Add("STT");
                    list1.Add("Họ và Tên");
                    list1.Add("Lương");// trường này sẽ colspan = 2
                    list1.Add("Phụ cấp lương");
                    list1.Add("Thưởng theo năng suất");
                    list1.Add("Công thực tế");
                    list1.Add("Lương thực tế");
                    list1.Add("Phép");
                    list1.Add("Lương phép");
                    list1.Add("Lương tăng ca");
                    list1.Add("Lương chờ việc");
                    list1.Add("Truy lĩnh");
                    list1.Add("Thực lĩnh");
                    list1.Add("BHXH + Y tế + TN");
                    list1.Add("Thuế TNCN");
                    list1.Add("Truy trừ");
                    list1.Add("Đoàn phí");
                    //list1.Add("Tiền ăn giữa ca");
                    list1.Add("Tiền điện thoại");
                    list1.Add("Còn lãnh");

                    var list2 = new List<string>();

                    list2.Add("STT");
                    list2.Add("Họ và Tên");
                    list2.Add("Lương 1");// trường này sẽ colspan = 2
                    list2.Add("Phụ cấp lương");
                    list2.Add("Thưởng theo năng suất");
                    list2.Add("Công thực tế");
                    list2.Add("Lương thực tế");
                    list2.Add("Phép");
                    list2.Add("Lương phép");
                    list2.Add("Lương tăng ca");
                    list2.Add("Lương chờ việc");
                    list2.Add("Truy lĩnh");
                    list2.Add("Thực lĩnh");
                    list2.Add("BHXH + Y tế + TN");
                    list2.Add("Thuế TNCN");
                    list2.Add("Truy trừ");
                    list2.Add("Đoàn phí");
                    //list2.Add("Tiền ăn giữa ca");
                    list2.Add("Tiền điện thoại");
                    list2.Add("Còn lãnh");

                    var list3 = new List<string>();

                    list3.Add("1");
                    list3.Add("2");
                    list3.Add("3");// trường này sẽ colspan = 2
                    list3.Add("4");
                    list3.Add("5");
                    list3.Add("6");
                    list3.Add("7");
                    list3.Add("8");
                    list3.Add("9");
                    list3.Add("10");
                    list3.Add("11");
                    list3.Add("12");
                    list3.Add("13");
                    list3.Add("14");
                    list3.Add("15");
                    list3.Add("16");
                    list3.Add("17");
                    // list3.Add("18");
                    list3.Add("18");
                    list3.Add("19=13-14-15-16-17+18");

                    var headerRow = sheet.CreateRow(idRowStart);
                    int rowend = idRowStart;
                    ReportHelperExcel.CreateHeaderRow(headerRow, 0, styleheadedColumnTable, list1);
                    idRowStart++;
                    var headerRow1 = sheet.CreateRow(idRowStart);
                    ReportHelperExcel.CreateHeaderRow(headerRow1, 0, styleheadedColumnTable, list2);
                    idRowStart++;
                    var headerRow11 = sheet.CreateRow(idRowStart);
                    ReportHelperExcel.CreateHeaderRow(headerRow11, 0, styleheadedColumnTable, list3);

                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 19, 19));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 18, 18));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 17, 17));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 16, 16));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 15, 15));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 14, 14));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 13, 13));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 12, 12));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 11, 11));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 10, 10));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 9, 9));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 8, 8));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 7, 7));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 6, 6));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 5, 5));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 4, 4));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow.RowNum, 2, 3));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 1, 1));
                    sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 0, 0));

                    sheet.SetColumnWidth(0, 8 * 210);
                    sheet.SetColumnWidth(1, 30 * 210);
                    sheet.SetColumnWidth(2, 15 * 210);
                    sheet.SetColumnWidth(3, 15 * 210);
                    sheet.SetColumnWidth(4, 15 * 210);
                    sheet.SetColumnWidth(5, 15 * 210);
                    sheet.SetColumnWidth(6, 15 * 210);
                    sheet.SetColumnWidth(7, 15 * 210);
                    sheet.SetColumnWidth(8, 15 * 210);
                    sheet.SetColumnWidth(9, 15 * 210);
                    sheet.SetColumnWidth(10, 15 * 210);
                    sheet.SetColumnWidth(11, 15 * 210);
                    sheet.SetColumnWidth(12, 15 * 210);
                    sheet.SetColumnWidth(13, 15 * 210);
                    sheet.SetColumnWidth(14, 15 * 210);
                    sheet.SetColumnWidth(15, 15 * 210);
                    sheet.SetColumnWidth(16, 15 * 210);
                    sheet.SetColumnWidth(17, 15 * 210);
                    sheet.SetColumnWidth(18, 15 * 210);
                    sheet.SetColumnWidth(19, 15 * 210);


                    double tongThucLinh = 0, tongTruyLanh = 0, tongTruyThu = 0;

                    foreach (var item1 in item.OrderBy(d => d.hoTen))
                    {
                        dem = 0;
                        stt++;
                        idRowStart++;

                        rowC = sheet.CreateRow(idRowStart);
                        ReportHelperExcel.SetAlignment(rowC, dem++, stt.ToString(), hStyleConCenter);
                        ReportHelperExcel.SetAlignment(rowC, dem++, item1.hoTen, hStyleConLeft);

                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.luongThoaThuan), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.phuCapLuong), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.khoanBoSungLuong), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", (item1.soNgayCongTac ?? 0) + (item1.soNgayQuet ?? 0) + (item1.soNgayNghiLe ?? 0)
                                + (item1.soNgayPhepLuyKeThangTruoc ?? 0) + (item1.soNgayLuyKeThangTruoc ?? 0)), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.luongThucTe), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", (item1.soNgayNghiPhep ?? 0)), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.luongNghiPhep), hStyleConRight);
                        //tăng ca
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.luongTangCa), hStyleConRight);
                        //Công chờ việc
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.luongChoViec), hStyleConRight);
                        double truyLanh = (item1.truyLanh ?? 0);
                        if ((item1.TTTLLuong ?? 0) > 0) truyLanh += (item1.TTTLLuong ?? 0);
                        if ((item1.TTTLThue ?? 0) > 0) truyLanh += (item1.TTTLThue ?? 0);

                        double truyThu = (item1.truyThu ?? 0);
                        if ((item1.TTTLLuong ?? 0) < 0) truyThu += (item1.TTTLLuong ?? 0);
                        if ((item1.TTTLThue ?? 0) < 0) truyThu += (item1.TTTLThue ?? 0);

                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", truyLanh), hStyleConRight);
                        double thucLinh = (item1.luongThucTe ?? 0) + (item1.luongNghiPhep ?? 0) + (item1.luongTangCa ?? 0) + (item1.luongChoViec ?? 0) + truyLanh;
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", thucLinh), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", (item1.baoHiem)), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.thue), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", truyThu), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.doanPhi), hStyleConRight);
                        //ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.phuCapTienAn), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.phuCapDienThoai), hStyleConRight);
                        ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item1.thucLanh), hStyleConRight);

                        tongThucLinh += thucLinh;
                        tongTruyLanh += truyLanh;
                        tongTruyThu += truyThu;
                    }

                    int demT = 0;
                    idRowStart++;
                    Row rowT = sheet.CreateRow(idRowStart);
                    ReportHelperExcel.SetAlignment(rowT, demT++, (stt).ToString(), styleheadedColumnTable);
                    ReportHelperExcel.SetAlignment(rowT, demT++, "", styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.luongThoaThuan)), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.phuCapLuong)), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", (item.Sum(s => s.khoanBoSungLuong) ?? 0)), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0.##}", item.Sum(s => ((s.soNgayCongTac ?? 0) + (s.soNgayQuet ?? 0) + (s.soNgayNghiLe ?? 0)
                                + (s.soNgayPhepLuyKeThangTruoc ?? 0) + (s.soNgayLuyKeThangTruoc ?? 0)))), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => (s.luongThucTe ?? 0))), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0.##}", item.Sum(s => ((s.soNgayNghiPhep ?? 0)))), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.luongNghiPhep)), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.luongTangCa)), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.luongChoViec)), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", tongTruyLanh), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", tongThucLinh), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => (s.baoHiem))), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.thue)), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", tongTruyThu), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.doanPhi)), styleCellSumary);
                    //ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.phuCapTienAn)), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.phuCapDienThoai)), styleCellSumary);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.thucLanh)), styleCellSumary);

                    idRowStart = idRowStart + 2;
                    string cellFooterSoTien = "(" + CharacterHelper.DocTienBangChu((decimal)item.Sum(s => s.thucLanh), string.Empty) + ")";
                    var titleCellFooterSoTien = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 12, cellFooterSoTien);
                    titleCellFooterSoTien.CellStyle = styleTitleItalic;

                    idRowStart = idRowStart + 1;
                    string cellFooterNgayLap = "Tp.Hồ Chí Minh, ngày   tháng  năm " + nam;
                    var titleCellFooterNgayLap = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 13, cellFooterNgayLap);
                    titleCellFooterNgayLap.CellStyle = styleTitleItalic;

                    idRowStart = idRowStart + 1;
                    string cellFooterPTC = "PHÒNG TÀI CHÍNH KẾ TOÁN";
                    var titleCellFooterPTC = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 1, cellFooterPTC);
                    titleCellFooterPTC.CellStyle = styleTitleNormal;

                    string cellFooterKT = "TRƯỞNG PHÒNG HÀNH CHÍNH - NHÂN SỰ";
                    var titleCellFooterKT = HSSFCellUtil.CreateCell(sheet.GetRow(idRowStart), 13, cellFooterKT);
                    titleCellFooterKT.CellStyle = styleTitleNormal;

                    idRowStart = idRowStart + 2;
                }

                // Sum tong theo bo phan tinh luong
                var danhSachBLGroupBysTong = nhanSuContext.sp_NS_BangLuongNhanVien(maPhongBan, thang, nam, qSearch).GroupBy(s => new { s.boPhanTinhLuong });
                var flashHeadTong = 0;
                double tongThucLinhBoPhan = 0, tongTruyLanhBoPhan = 0, tongTruyThuBoPhan = 0;

                foreach (var item in danhSachBLGroupBysTong)
                {
                    #region định nghĩa

                    count++;
                    if (flashHeadTong == 0)
                    {
                        var titleCellCty = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 0, cellTenCty.ToUpper());
                        titleCellCty.CellStyle = styleTitle;
                        idRowStart++;

                        var titleCellTitleMain = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 5, cellTitleMainTH.ToUpper());
                        titleCellTitleMain.CellStyle = styleTitle;
                        idRowStart++;
                        var list1 = new List<string>();

                        list1.Add("TS  NV");
                        list1.Add("Đơn vị");
                        list1.Add("Lương");// trường này sẽ colspan = 2
                        list1.Add("Lương phụ cấp");
                        list1.Add("Thưởng theo năng suất");
                        list1.Add("Công thực tế");
                        list1.Add("Lương thực tế");
                        list1.Add("Phép");
                        list1.Add("Lương phép");
                        list1.Add("Lương tăng ca");
                        list1.Add("Lương chờ việc");
                        list1.Add("Truy lĩnh");
                        list1.Add("Thực lĩnh");
                        list1.Add("BHXH + Y tế + TN");
                        list1.Add("Thuế TNCN");
                        list1.Add("Truy trừ");
                        list1.Add("Đoàn phí");
                        //list1.Add("Tiền ăn giữa ca");
                        list1.Add("Tiền điện thoại");
                        list1.Add("Còn lãnh");

                        var list2 = new List<string>();
                        list2.Add("TS  NV");
                        list2.Add("Đơn vị");
                        list2.Add("Lương");// trường này sẽ colspan = 2
                        list2.Add("Lương phụ cấp");
                        list2.Add("Thưởng theo năng suất");
                        list2.Add("Công thực tế");
                        list2.Add("Lương thực tế");
                        list2.Add("Phép");
                        list2.Add("Lương phép");
                        list2.Add("Lương tăng ca");
                        list2.Add("Lương chờ việc");
                        list2.Add("Truy lĩnh");
                        list2.Add("Thực lĩnh");
                        list2.Add("BHXH + Y tế + TN");
                        list2.Add("Thuế TNCN");
                        list2.Add("Truy trừ");
                        list2.Add("Đoàn phí");
                        //list2.Add("Tiền ăn giữa ca");
                        list2.Add("Tiền điện thoại");
                        list2.Add("Còn lãnh");

                        var headerRow = sheet.CreateRow(idRowStart);
                        int rowend = idRowStart;
                        ReportHelperExcel.CreateHeaderRow(headerRow, 0, styleheadedColumnTable, list1);
                        idRowStart++;
                        var headerRow1 = sheet.CreateRow(idRowStart);
                        ReportHelperExcel.CreateHeaderRow(headerRow1, 0, styleheadedColumnTable, list2);

                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 19, 19));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 18, 18));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 17, 17));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 16, 16));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 15, 15));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 14, 14));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 13, 13));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 12, 12));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 11, 11));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 10, 10));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 9, 9));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 8, 8));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 7, 7));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 6, 6));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 5, 5));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 4, 4));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow.RowNum, 2, 3));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 1, 1));
                        sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 0, 0));



                        sheet.SetColumnWidth(0, 8 * 210);
                        sheet.SetColumnWidth(1, 30 * 210);
                        sheet.SetColumnWidth(2, 15 * 210);
                        sheet.SetColumnWidth(3, 15 * 210);
                        sheet.SetColumnWidth(4, 15 * 210);
                        sheet.SetColumnWidth(5, 15 * 210);
                        sheet.SetColumnWidth(6, 15 * 210);
                        sheet.SetColumnWidth(7, 15 * 210);
                        sheet.SetColumnWidth(8, 15 * 210);
                        sheet.SetColumnWidth(9, 15 * 210);
                        sheet.SetColumnWidth(10, 15 * 210);
                        sheet.SetColumnWidth(11, 15 * 210);
                        sheet.SetColumnWidth(12, 15 * 210);
                        sheet.SetColumnWidth(13, 15 * 210);
                        sheet.SetColumnWidth(14, 15 * 210);
                        sheet.SetColumnWidth(15, 15 * 210);
                        sheet.SetColumnWidth(16, 15 * 210);
                        sheet.SetColumnWidth(17, 15 * 210);
                        sheet.SetColumnWidth(18, 15 * 210);
                        sheet.SetColumnWidth(19, 20 * 210);

                    }
                    flashHeadTong = 1;

                    string tenPhongBan = string.Empty;
                    if (item.Count() > 0)
                    {
                        tenPhongBan = item.Key.boPhanTinhLuong;
                    }
                    int demT = 0;
                    idRowStart++;

                    #endregion

                    Row rowT = sheet.CreateRow(idRowStart);
                    ReportHelperExcel.SetAlignment(rowT, demT++, item.Count().ToString(), styleheadedColumnTable);
                    ReportHelperExcel.SetAlignment(rowT, demT++, tenPhongBan, styleCellSumaryLeft);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.luongThoaThuan)), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.phuCapLuong)), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", (item.Sum(s => s.khoanBoSungLuong) ?? 0)), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0.##}", item.Sum(s => ((s.soNgayCongTac ?? 0) + (s.soNgayQuet ?? 0) + (s.soNgayNghiLe ?? 0)
                            + (s.soNgayPhepLuyKeThangTruoc ?? 0) + (s.soNgayLuyKeThangTruoc ?? 0)))), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => (s.luongThucTe ?? 0))), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0.##}", item.Sum(s => ((s.soNgayNghiPhep ?? 0)))), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.luongNghiPhep)), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.luongTangCa)), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.luongChoViec)), hStyleConRight);

                    //double truyLanh = (item1.truyLanh ?? 0);
                    //if ((item1.TTTLLuong ?? 0) > 0) truyLanh += (item1.TTTLLuong ?? 0);
                    //if ((item1.TTTLThue ?? 0) > 0) truyLanh += (item1.TTTLThue ?? 0);

                    //double truyThu = (item1.truyThu ?? 0);
                    //if ((item1.TTTLLuong ?? 0) < 0) truyThu += (item1.TTTLLuong ?? 0);
                    //if ((item1.TTTLThue ?? 0) < 0) truyThu += (item1.TTTLThue ?? 0);

                    double truyLanh = item.Sum(s => (s.truyLanh ?? 0) + ((s.TTTLLuong ?? 0) > 0 ? (s.TTTLLuong ?? 0) : 0) + ((s.TTTLThue ?? 0) > 0 ? (s.TTTLThue ?? 0) : 0));
                    double truyThu = item.Sum(s => (s.truyThu ?? 0) + ((s.TTTLLuong ?? 0) < 0 ? (s.TTTLLuong ?? 0) : 0) + ((s.TTTLThue ?? 0) < 0 ? (s.TTTLThue ?? 0) : 0));

                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", truyLanh), hStyleConRight);

                    double thucLinh = item.Sum(s => (s.luongThucTe ?? 0) + (s.luongNghiPhep ?? 0) + (s.luongTangCa ?? 0) + (s.luongChoViec ?? 0)) + truyLanh;

                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", thucLinh), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => (s.baoHiem))), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.thue)), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", truyThu), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.doanPhi)), hStyleConRight);
                    //ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.dangPhi)), styleCellSumary);
                    //ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.phuCapTienAn)), hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.phuCapDienThoai)), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", item.Sum(s => s.thucLanh)), styleCellSumary);

                    tongThucLinhBoPhan += thucLinh;
                    tongTruyLanhBoPhan += truyLanh;
                    tongTruyThuBoPhan += truyThu;

                }
                var danhSachBLGroupBysTongSum = nhanSuContext.sp_NS_BangLuongNhanVien(maPhongBan, thang, nam, qSearch).ToList();

                int demTong = 0;
                idRowStart++;
                Row rowTong = sheet.CreateRow(idRowStart);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0.##}", danhSachBLGroupBysTongSum.Count()), styleheadedColumnTable);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, "Tổng cộng", styleCellSumaryLeft);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.luongThoaThuan)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.phuCapLuong)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", (danhSachBLGroupBysTongSum.Sum(s => s.khoanBoSungLuong) ?? 0)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0.##}", danhSachBLGroupBysTongSum.Sum(s => ((s.soNgayCongTac ?? 0) + (s.soNgayQuet ?? 0) + (s.soNgayNghiLe ?? 0)
                        + (s.soNgayPhepLuyKeThangTruoc ?? 0) + (s.soNgayLuyKeThangTruoc ?? 0)))), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.luongThucTe)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0.##}", danhSachBLGroupBysTongSum.Sum(s => ((s.soNgayNghiPhep ?? 0)))), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.luongNghiPhep)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.luongTangCa)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.luongChoViec)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", tongTruyLanhBoPhan), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", tongThucLinhBoPhan), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => (s.baoHiem))), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.thue)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", tongTruyThuBoPhan), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.doanPhi)), styleCellSumary);
                //ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.phuCapTienAn)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.phuCapDienThoai)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowTong, demTong++, String.Format("{0:#,##0}", danhSachBLGroupBysTongSum.Sum(s => s.thucLanh)), styleCellSumary);

                idRowStart = idRowStart + 2;
                string cellFooterSoTienTong = "(" + CharacterHelper.DocTienBangChu((decimal)danhSachBLGroupBysTongSum.Sum(s => s.thucLanh), string.Empty) + ")";
                var titleCellFooterSoTienTong = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 12, cellFooterSoTienTong);
                titleCellFooterSoTienTong.CellStyle = styleTitleItalic;

                idRowStart = idRowStart + 1;
                string cellFooterNgayLapTong = "Tp.Hồ Chí Minh, ngày   tháng  năm " + nam;
                var titleCellFooterNgayLapTong = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 13, cellFooterNgayLapTong);
                titleCellFooterNgayLapTong.CellStyle = styleTitleItalic;

                idRowStart = idRowStart + 1;
                string cellFooterPTCTong = "TRƯỞNG PHÒNG HÀNH CHÍNH - NHÂN SỰ";
                var titleCellFooterPTCTong = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 1, cellFooterPTCTong);
                titleCellFooterPTCTong.CellStyle = styleTitleBold;

                string cellFooterKTTong = "PHÒNG TÀI CHÍNH - KẾ TOÁN";
                var titleCellFooterKTTong = HSSFCellUtil.CreateCell(sheet.GetRow(idRowStart), 7, cellFooterKTTong);
                titleCellFooterKTTong.CellStyle = styleTitleBold;
                string cellFooterKTTongTGD = "TỔNG GIÁM ĐỐC";
                var titleCellFooterKTTongTGD = HSSFCellUtil.CreateCell(sheet.GetRow(idRowStart), 13, cellFooterKTTongTGD);
                titleCellFooterKTTongTGD.CellStyle = styleTitleBold;
                // End sum 


                var stream = new MemoryStream();
                workbook.Write(stream);

                Response.ContentType = "application/vnd.ms-excel";
                Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));
                Response.Clear();

                Response.BinaryWrite(stream.GetBuffer());
                Response.End();
            }
            catch
            {

            }
        }
        #endregion



        #region  Xuat File Bang Luong Nhan Vien theo bộ phận
        public void XuatFileBLNVBoPhan(int thang, int nam)
        {
            try
            {
                var filename = "";
                var virtualPath = HttpRuntime.AppDomainAppVirtualPath;

                var fileStream = new FileStream(System.Web.HttpContext.Current.Server.MapPath(virtualPath + @"\Content\Report\ReportTemplate.xls"), FileMode.Open, FileAccess.Read);

                var workbook = new HSSFWorkbook(fileStream, true);
                filename += "BangTongHopLuongCacBoPhan_" + nam + "_" + thang + ".xls";

                #region format style excel cell
                /*style title start*/
                //tạo font cho các title
                //font tiêu đề 
                HSSFFont hFontTieuDe = (HSSFFont)workbook.CreateFont();
                hFontTieuDe.FontHeightInPoints = 11;
                hFontTieuDe.Boldweight = 100 * 10;
                hFontTieuDe.FontName = "Times New Roman";
                //hFontTieuDe.Color = HSSFColor.BLUE.index;

                //font tiêu đề 
                HSSFFont hFontTieuDeUnderline = (HSSFFont)workbook.CreateFont();
                hFontTieuDeUnderline.FontHeightInPoints = 11;
                hFontTieuDeUnderline.Boldweight = 100 * 10;
                hFontTieuDeUnderline.FontName = "Times New Roman";
                hFontTieuDeUnderline.Underline = 1;
                //hFontTieuDe.Color = HSSFColor.BLUE.index;


                HSSFFont hFontTieuDeItalic = (HSSFFont)workbook.CreateFont();
                hFontTieuDeItalic.FontHeightInPoints = 11;
                //hFontTieuDeItalic.Boldweight = 100 * 10;
                hFontTieuDeItalic.FontName = "Times New Roman";
                hFontTieuDeItalic.IsItalic = true;
                //hFontTieuDe.Color = HSSFColor.BLUE.index;


                HSSFFont hFontTieuDeLarge = (HSSFFont)workbook.CreateFont();
                hFontTieuDeLarge.FontHeightInPoints = 16;
                hFontTieuDeLarge.Boldweight = 100 * 10;
                hFontTieuDeLarge.FontName = "Times New Roman";
                //hFontTieuDeLarge.Color = HSSFColor.BLUE.index;

                //font tiêu đề 
                HSSFFont hFontTongGiaTriHT = (HSSFFont)workbook.CreateFont();
                hFontTongGiaTriHT.FontHeightInPoints = 11;
                hFontTongGiaTriHT.Boldweight = (short)FontBoldWeight.BOLD;
                hFontTongGiaTriHT.FontName = "Times New Roman";
                hFontTongGiaTriHT.Color = HSSFColor.BLACK.index;

                //font thông tin bảng tính
                HSSFFont hFontTT = (HSSFFont)workbook.CreateFont();
                hFontTT.IsItalic = true;
                hFontTT.Boldweight = (short)FontBoldWeight.BOLD;
                hFontTT.Color = HSSFColor.BLACK.index;
                hFontTT.FontName = "Times New Roman";
                hFontTieuDe.FontHeightInPoints = 11;

                //font chứ hoa đậm
                HSSFFont hFontNommalUpper = (HSSFFont)workbook.CreateFont();
                hFontNommalUpper.Boldweight = (short)FontBoldWeight.BOLD;
                hFontNommalUpper.Color = HSSFColor.BLACK.index;
                hFontNommalUpper.FontName = "Times New Roman";

                //font chữ bình thường
                HSSFFont hFontNommal = (HSSFFont)workbook.CreateFont();
                hFontNommal.Color = HSSFColor.BLACK.index;
                hFontNommal.FontName = "Times New Roman";

                //font chữ bình thường đậm
                HSSFFont hFontNommalBold = (HSSFFont)workbook.CreateFont();
                hFontNommalBold.Color = HSSFColor.BLACK.index;
                hFontNommalBold.Boldweight = (short)FontBoldWeight.BOLD;
                hFontNommalBold.FontName = "Times New Roman";

                //tạo font cho các title end

                //Set style
                var styleTitle = workbook.CreateCellStyle();
                styleTitle.SetFont(hFontTieuDe);
                styleTitle.Alignment = HorizontalAlignment.LEFT;

                //Set styleUnderline
                var styleTitleUnderline = workbook.CreateCellStyle();
                styleTitleUnderline.SetFont(hFontTieuDeUnderline);
                styleTitleUnderline.Alignment = HorizontalAlignment.LEFT;

                //Set style In nghiêng
                var styleTitleItalic = workbook.CreateCellStyle();
                styleTitleItalic.SetFont(hFontTieuDeItalic);
                styleTitleItalic.Alignment = HorizontalAlignment.LEFT;

                //Set style Large font
                var styleTitleLarge = workbook.CreateCellStyle();
                styleTitleLarge.SetFont(hFontTieuDeLarge);
                styleTitleLarge.Alignment = HorizontalAlignment.LEFT;

                //style infomation
                var styleInfomation = workbook.CreateCellStyle();
                styleInfomation.SetFont(hFontTT);
                styleInfomation.Alignment = HorizontalAlignment.LEFT;

                //style header
                var styleheadedColumnTable = workbook.CreateCellStyle();
                styleheadedColumnTable.SetFont(hFontNommalUpper);
                styleheadedColumnTable.WrapText = true;
                styleheadedColumnTable.BorderBottom = CellBorderType.THIN;
                styleheadedColumnTable.BorderLeft = CellBorderType.THIN;
                styleheadedColumnTable.BorderRight = CellBorderType.THIN;
                styleheadedColumnTable.BorderTop = CellBorderType.THIN;
                styleheadedColumnTable.VerticalAlignment = VerticalAlignment.CENTER;
                styleheadedColumnTable.Alignment = HorizontalAlignment.CENTER;

                //style sum cell
                var styleCellSumary = workbook.CreateCellStyle();
                styleCellSumary.SetFont(hFontNommalUpper);
                styleCellSumary.WrapText = true;
                styleCellSumary.BorderBottom = CellBorderType.THIN;
                styleCellSumary.BorderLeft = CellBorderType.THIN;
                styleCellSumary.BorderRight = CellBorderType.THIN;
                styleCellSumary.BorderTop = CellBorderType.THIN;
                styleCellSumary.VerticalAlignment = VerticalAlignment.CENTER;
                styleCellSumary.Alignment = HorizontalAlignment.RIGHT;

                var styleHeading1 = workbook.CreateCellStyle();
                styleHeading1.SetFont(hFontNommalBold);
                styleHeading1.WrapText = true;
                styleHeading1.BorderBottom = CellBorderType.THIN;
                styleHeading1.BorderLeft = CellBorderType.THIN;
                styleHeading1.BorderRight = CellBorderType.THIN;
                styleHeading1.BorderTop = CellBorderType.THIN;
                styleHeading1.VerticalAlignment = VerticalAlignment.CENTER;
                styleHeading1.Alignment = HorizontalAlignment.LEFT;

                var hStyleConLeft = (HSSFCellStyle)workbook.CreateCellStyle();
                hStyleConLeft.SetFont(hFontNommal);
                hStyleConLeft.VerticalAlignment = VerticalAlignment.TOP;
                hStyleConLeft.Alignment = HorizontalAlignment.LEFT;
                hStyleConLeft.WrapText = true;
                hStyleConLeft.BorderBottom = CellBorderType.THIN;
                hStyleConLeft.BorderLeft = CellBorderType.THIN;
                hStyleConLeft.BorderRight = CellBorderType.THIN;
                hStyleConLeft.BorderTop = CellBorderType.THIN;

                var hStyleConRight = (HSSFCellStyle)workbook.CreateCellStyle();
                hStyleConRight.SetFont(hFontNommal);
                hStyleConRight.VerticalAlignment = VerticalAlignment.TOP;
                hStyleConRight.Alignment = HorizontalAlignment.RIGHT;
                hStyleConRight.BorderBottom = CellBorderType.THIN;
                hStyleConRight.BorderLeft = CellBorderType.THIN;
                hStyleConRight.BorderRight = CellBorderType.THIN;
                hStyleConRight.BorderTop = CellBorderType.THIN;


                var hStyleConCenter = (HSSFCellStyle)workbook.CreateCellStyle();
                hStyleConCenter.SetFont(hFontNommal);
                hStyleConCenter.VerticalAlignment = VerticalAlignment.TOP;
                hStyleConCenter.Alignment = HorizontalAlignment.CENTER;
                hStyleConCenter.BorderBottom = CellBorderType.THIN;
                hStyleConCenter.BorderLeft = CellBorderType.THIN;
                hStyleConCenter.BorderRight = CellBorderType.THIN;
                hStyleConCenter.BorderTop = CellBorderType.THIN;
                //set style end
                #endregion

                //Khai báo row
                Row rowC = null;



                var sheet = workbook.CreateSheet("BangLuongTheoBoPhan");

                //Khai báo row đầu tiên
                int firstRowNumber = 3;

                string cellTenCty = "TỔNG CÔNG TY XDCTGT 6 - CÔNG TY CỔ PHẦN";
                var titleCellCty = HSSFCellUtil.CreateCell(sheet.CreateRow(0), 0, cellTenCty.ToUpper());
                titleCellCty.CellStyle = styleTitle;

                string cellTenCacBanDH = "CÁC BAN ĐIỀU HÀNH DỰ ÁN";
                var titleCellTenCacBanDH = HSSFCellUtil.CreateCell(sheet.CreateRow(1), 1, cellTenCacBanDH.ToUpper());
                titleCellTenCacBanDH.CellStyle = styleTitle;

                string cellTitleMain = "BẢNG TỔNG HỢP LƯƠNG THÁNG " + thang + "/" + nam;
                var titleCellTitleMain = HSSFCellUtil.CreateCell(sheet.CreateRow(2), 5, cellTitleMain.ToUpper());
                titleCellTitleMain.CellStyle = styleTitle;

                firstRowNumber++;

                var list1 = new List<string>();
                list1.Add("TS");
                list1.Add("Đơn vị");
                list1.Add("Tổng mức lương");// trường này sẽ colspan = 2
                list1.Add("");
                list1.Add("Khoản bổ sung");
                list1.Add("Công");
                list1.Add("Tổng cộng");
                list1.Add("Lễ, phép");
                list1.Add("Lương lễ, phép");
                list1.Add("Truy lĩnh");
                list1.Add("Thực lĩnh");
                list1.Add("BHXH + Y tế + TN");
                list1.Add("Thuế TNCN");
                list1.Add("Truy thu");
                list1.Add("Đoàn phí");
                list1.Add("Đảng phí");
                //list1.Add("Tiền ăn giữa ca");
                list1.Add("Còn lãnh");


                var list2 = new List<string>();
                list2.Add("TS");
                list2.Add("Đơn vị");
                list2.Add("Lương");// trường này sẽ colspan = 2
                list2.Add("Phụ cấp lương");
                list2.Add("Khoản bổ sung");
                list2.Add("Công");
                list2.Add("Tổng cộng");
                list2.Add("Lễ, phép");
                list2.Add("Lương lễ, phép");
                list2.Add("Truy lĩnh");
                list2.Add("Thực lĩnh");
                list2.Add("BHXH + Y tế + TN");
                list2.Add("Thuế TNCN");
                list2.Add("Truy thu");
                list2.Add("Đoàn phí");
                list2.Add("Đảng phí");
                //list2.Add("Tiền ăn giữa ca");
                list2.Add("Còn lãnh");

                var idRowStart = firstRowNumber; // bat dau o dong thu 4
                var headerRow = sheet.CreateRow(idRowStart);
                int rowend = idRowStart;
                ReportHelperExcel.CreateHeaderRow(headerRow, 0, styleheadedColumnTable, list1);
                idRowStart++;
                var headerRow1 = sheet.CreateRow(idRowStart);
                ReportHelperExcel.CreateHeaderRow(headerRow1, 0, styleheadedColumnTable, list2);

                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 17, 17));
                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 16, 16));
                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 15, 15));
                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 14, 14));
                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 13, 13));
                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 12, 12));
                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 11, 11));
                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 10, 10));
                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 9, 9));
                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 8, 8));
                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 7, 7));
                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 6, 6));
                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 5, 5));
                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 4, 4));
                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow.RowNum, 2, 3));
                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 1, 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 0, 0));



                sheet.SetColumnWidth(0, 5 * 210);
                sheet.SetColumnWidth(1, 30 * 210);
                sheet.SetColumnWidth(2, 15 * 210);
                sheet.SetColumnWidth(3, 15 * 210);
                sheet.SetColumnWidth(4, 15 * 210);
                sheet.SetColumnWidth(5, 15 * 210);
                sheet.SetColumnWidth(6, 15 * 210);
                sheet.SetColumnWidth(7, 15 * 210);
                sheet.SetColumnWidth(8, 15 * 210);
                sheet.SetColumnWidth(9, 15 * 210);
                sheet.SetColumnWidth(10, 15 * 210);
                sheet.SetColumnWidth(11, 15 * 210);

                sheet.SetColumnWidth(12, 15 * 210);
                sheet.SetColumnWidth(13, 15 * 210);
                sheet.SetColumnWidth(14, 15 * 210);
                sheet.SetColumnWidth(15, 10 * 210);
                sheet.SetColumnWidth(16, 15 * 210);
                sheet.SetColumnWidth(17, 15 * 210);


                var data = nhanSuContext.sp_NS_BangLuongTheoBoPhan(thang, nam).ToList();
                var stt = 0;
                int dem = 0;
                double? sumLuongLe = 0;

                foreach (var item in data)
                {
                    dem = 0;
                    double? luongNghiLePhep = ((item.tongLuong ?? 0) * ((item.soNgayNghiLe ?? 0) + (item.soNgayNghiPhep ?? 0))) / item.ngayCongChuan;
                    sumLuongLe += luongNghiLePhep;
                    stt++;
                    idRowStart++;

                    rowC = sheet.CreateRow(idRowStart);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item.tongSo.ToString(), hStyleConCenter);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item.phongBan, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item.luongThang), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item.phuCapLuong), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item.khoanBoSungLuong), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", (item.soNgayCongTac ?? 0) + (item.soNgayQuet ?? 0)), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item.tongLuong), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", (item.soNgayNghiLe ?? 0) + (item.soNgayNghiPhep ?? 0)), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", luongNghiLePhep), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item.truyLanh), hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item.thucLanh), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", (item.baoHiemTN + item.baoHiemXH + item.baoHiemYTe)), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item.thue), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item.truyThu), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item.doanPhi), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item.dangPhi), hStyleConRight);
                    // ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item.phuCapTienAn), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0}", item.thucLanh), hStyleConRight);
                }

                int demT = 0;
                idRowStart++;
                Row rowT = sheet.CreateRow(idRowStart);

                ReportHelperExcel.SetAlignment(rowT, demT++, data.Sum(s => s.tongSo).ToString(), hStyleConCenter);
                ReportHelperExcel.SetAlignment(rowT, demT++, "TỔNG CỘNG", styleCellSumary);
                ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", data.Sum(s => s.luongThang)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", data.Sum(s => s.phuCapLuong)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", data.Sum(s => s.khoanBoSungLuong)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0.##}", data.Sum(s => ((s.soNgayCongTac ?? 0) + (s.soNgayQuet ?? 0)))), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", data.Sum(s => s.tongLuong)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0.##}", data.Sum(s => ((s.soNgayNghiLe ?? 0) + (s.soNgayNghiPhep ?? 0)))), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", sumLuongLe), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", data.Sum(s => s.truyLanh)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", data.Sum(s => s.thucLanh)), styleCellSumary);

                ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", data.Sum(s => s.baoHiem)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", data.Sum(s => s.thue)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", data.Sum(s => s.truyThu)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", data.Sum(s => s.doanPhi)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", data.Sum(s => s.dangPhi)), styleCellSumary);
                //ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", data.Sum(s => s.phuCapTienAn)), styleCellSumary);
                ReportHelperExcel.SetAlignment(rowT, demT++, String.Format("{0:#,##0}", data.Sum(s => s.thucLanh)), styleCellSumary);

                idRowStart = idRowStart + 2;
                string cellFooterSoTien = "(" + CharacterHelper.DocTienBangChu((decimal)data.Sum(s => s.thucLanh), string.Empty) + ")";
                var titleCellFooterSoTien = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 12, cellFooterSoTien);
                titleCellFooterSoTien.CellStyle = styleTitleItalic;

                idRowStart = idRowStart + 2;
                string cellFooterNgayLap = "Tp.Hồ Chí Minh, ngày   tháng  năm " + nam;
                var titleCellFooterNgayLap = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 13, cellFooterNgayLap);
                titleCellFooterNgayLap.CellStyle = styleTitleItalic;

                idRowStart = idRowStart + 2;
                string cellFooterPTC = "PHÒNG TỔ CHỨC CB-LĐ";
                var titleCellFooterPTC = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 1, cellFooterPTC);
                titleCellFooterPTC.CellStyle = styleTitle;

                string cellFooterKT = "PHÒNG TÀI CHÍNH KẾ TOÁN";
                var titleCellFooterKT = HSSFCellUtil.CreateCell(sheet.GetRow(idRowStart), 7, cellFooterKT);
                titleCellFooterKT.CellStyle = styleTitle;

                string cellFooterTGD = "TỔNG GIÁM ĐỐC";
                var titleCellFooterTGD = HSSFCellUtil.CreateCell(sheet.GetRow(idRowStart), 14, cellFooterTGD);
                titleCellFooterTGD.CellStyle = styleTitle;


                var stream = new MemoryStream();
                workbook.Write(stream);

                Response.ContentType = "application/vnd.ms-excel";
                Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));
                Response.Clear();

                Response.BinaryWrite(stream.GetBuffer());
                Response.End();
            }
            catch
            {

            }



            //var filename = "";
            //var virtualPath = HttpRuntime.AppDomainAppVirtualPath;

            //filename += "BangLuongNVBoPhan_" + nam + "_" + thang + ".xlsx";

            //using (ExcelPackage package = new ExcelPackage())
            //{
            //    //Create a sheet
            //    package.Workbook.Worksheets.Add("BangLuongNVBoPhan_" + nam + "_" + thang);
            //    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
            //    //Header
            //    //insert từ dòng nào, bao nhiêu row
            //    var rowFrom = 1;
            //    worksheet.InsertRow(rowFrom, 1);
            //    worksheet.Cells[1, 1].Value = "STT";
            //    worksheet.Cells[1, 2].Value = "Phòng ban";
            //    worksheet.Cells[1, 3].Value = "Tháng";
            //    worksheet.Cells[1, 4].Value = "Năm";
            //    worksheet.Cells[1, 5].Value = "Tổng lương";
            //    worksheet.Cells[1, 6].Value = "Bảo hiểm";
            //    worksheet.Cells[1, 7].Value = "Thuế";
            //    worksheet.Cells[1, 8].Value = "Truy thu";
            //    worksheet.Cells[1, 9].Value = "Truy lãnh";
            //    worksheet.Cells[1, 10].Value = "Thực lãnh";
            //    worksheet.Column(2).Width = 35;
            //    worksheet.Column(3).Width = 35;
            //    worksheet.Column(4).Width = 20;
            //    worksheet.Column(5).Width = 20;
            //    worksheet.Column(6).Width = 20;
            //    worksheet.Column(7).Width = 20;
            //    worksheet.Column(8).Width = 20;
            //    worksheet.Column(9).Width = 20;
            //    worksheet.Column(10).Width = 20;
            //    //// Formatting style of the header
            //    using (var range = worksheet.Cells[1, 1, 2, 10])
            //    {
            //        // Setting bold font
            //        range.Style.Font.Bold = true;
            //        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //        // Setting fill type solid
            //        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            //        // Setting background color dark blue
            //        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            //        // Setting font color
            //        range.Style.Font.Color.SetColor(System.Drawing.Color.White);
            //    }



            //    #region
            //    //Body
            //    var data = nhanSuContext.sp_NS_BangLuongTheoBoPhan(thang, nam).ToList();

            //    if (data != null && data.Count > 0)
            //    {
            //        var countSTT = 1;
            //        foreach (var item in data)
            //        {

            //            rowFrom = rowFrom + 1;
            //            worksheet.InsertRow(rowFrom, 1);
            //            worksheet.Cells[rowFrom, 1].Value = countSTT++;
            //            worksheet.Cells[rowFrom, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
            //            worksheet.Cells[rowFrom, 1].Style.Font.Bold = true;
            //            worksheet.Cells[rowFrom, 2].Value = item.boPhanTinhLuong;
            //            worksheet.Cells[rowFrom, 3].Value = item.thang;
            //            worksheet.Cells[rowFrom, 4].Value = item.nam;


            //            worksheet.Cells[rowFrom, 5].Value = item.tongLuong;
            //            worksheet.Cells[rowFrom, 5].Style.Numberformat.Format = "#,##0.000";

            //            worksheet.Cells[rowFrom, 6].Value = item.baoHiem;
            //            worksheet.Cells[rowFrom, 6].Style.Numberformat.Format = "#,##0.000";

            //            worksheet.Cells[rowFrom, 7].Value = item.thue;
            //            worksheet.Cells[rowFrom, 7].Style.Numberformat.Format = "#,##0";

            //            worksheet.Cells[rowFrom, 8].Value = item.truyThu;
            //            worksheet.Cells[rowFrom, 8].Style.Numberformat.Format = "#,##0.000";

            //            worksheet.Cells[rowFrom, 9].Value = item.truyLanh;
            //            worksheet.Cells[rowFrom, 9].Style.Numberformat.Format = "#,##0.000";

            //            worksheet.Cells[rowFrom, 10].Value = item.thucLanh;
            //            worksheet.Cells[rowFrom, 10].Style.Numberformat.Format = "#,##0.000";
            //        }
            //    }

            //    #endregion

            //    //Generate A File
            //    Byte[] bin = package.GetAsByteArray();

            //    Response.BinaryWrite(bin);
            //    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //    Response.AddHeader("content-disposition", string.Format("attachment;filename={0}", filename));
            //}
        }


        #endregion End Xuat File
        // Xuat File Bang Luong Nhan Vien theo bộ phận
        public void XuatFileBLNN(int thang, int nam)
        {

            string maPhongBan = "";
            string qSearch = "";
            var filename = "";
            var virtualPath = HttpRuntime.AppDomainAppVirtualPath;

            filename += "BangLuongChuyenNN_" + nam + "_" + thang + ".xlsx";

            using (ExcelPackage package = new ExcelPackage())
            {
                //Create a sheet
                package.Workbook.Worksheets.Add("BangLuongChuyenNN_" + nam + "_" + thang);
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                //Header
                //insert từ dòng nào, bao nhiêu row
                var rowFrom = 1;
                worksheet.InsertRow(rowFrom, 1);
                worksheet.Cells[1, 1].Value = "STT";
                worksheet.Cells[1, 2].Value = "Họ tên";
                worksheet.Cells[1, 3].Value = "Số tài khoản";
                worksheet.Cells[1, 4].Value = "Tên ngân hàng";
                worksheet.Cells[1, 5].Value = "Số CMND";
                worksheet.Cells[1, 6].Value = "Thực lãnh";

                worksheet.Column(2).Width = 35;
                worksheet.Column(3).Width = 35;
                worksheet.Column(4).Width = 20;
                worksheet.Column(5).Width = 20;
                worksheet.Column(6).Width = 20;
                //// Formatting style of the header
                using (var range = worksheet.Cells[1, 1, 2, 6])
                {
                    // Setting bold font
                    range.Style.Font.Bold = true;
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    // Setting fill type solid
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    // Setting background color dark blue
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                    // Setting font color
                    range.Style.Font.Color.SetColor(System.Drawing.Color.White);
                }



                #region
                //Body
                var data = nhanSuContext.sp_NS_BangLuongChuyenNganHang(thang, nam, maPhongBan, qSearch).ToList();

                if (data != null && data.Count > 0)
                {
                    var countSTT = 1;
                    foreach (var item in data)
                    {

                        rowFrom = rowFrom + 1;
                        worksheet.InsertRow(rowFrom, 1);
                        worksheet.Cells[rowFrom, 1].Value = countSTT++;
                        worksheet.Cells[rowFrom, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                        worksheet.Cells[rowFrom, 1].Style.Font.Bold = true;
                        worksheet.Cells[rowFrom, 2].Value = item.hoTen;
                        worksheet.Cells[rowFrom, 3].Value = item.soTaiKhoan;
                        worksheet.Cells[rowFrom, 4].Value = item.tenNganHang;
                        worksheet.Cells[rowFrom, 5].Value = item.soCMND;
                        worksheet.Cells[rowFrom, 6].Value = item.thucLanh;
                        worksheet.Cells[rowFrom, 6].Style.Numberformat.Format = "#,##0.000";
                        worksheet.Cells[rowFrom, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    }
                }

                #endregion

                //Generate A File
                Byte[] bin = package.GetAsByteArray();

                Response.BinaryWrite(bin);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", string.Format("attachment;filename={0}", filename));
            }
        }

        // End Xuat File

        public ActionResult CheckTaoDNCL(int? thang, int? nam)
        {
            string hasValue = string.Empty;
            var isTonTai = lqThuanViet.tbl_DeNghiChiLuongs.Where(d => d.thang == thang && d.nam == nam).FirstOrDefault();
            if (isTonTai != null)
            {
                hasValue = "true";
            }
            return Json(hasValue);
        }
        public ActionResult CreateCL(int? thang, int? nam)
        {
            DeNghiChiLuong deNghiCL = new DeNghiChiLuong();
            deNghiCL.soPhieu = CheckLetterDNCL("DNCL", GetMaxDeNghiChiLuong(), 3);
            deNghiCL.thang = thang;
            deNghiCL.nam = nam;
            deNghiCL.maNguoiLap = GetUser().manv;
            deNghiCL.tenNguoiLap = HoVaTen(GetUser().manv);
            deNghiCL.ngayLap = DateTime.Now;
            var chiTietLuong = (from p in nhanSuContext.sp_NS_ChiTietLuong_DeNghiChiLuong(thang, nam)
                                select new DeNghiChiLuongChiTiet
                                {
                                    BHXH = p.BHXH,
                                    ChuyenKhoan = p.chuyenKhoan,
                                    LuongThang = p.luongThang,
                                    TenBoPhanTinhLuong = p.boPhanTinhLuong,
                                    ThueTNCN = p.thueTNCN,
                                    TruyLanhTruyThu = p.truyLanhTruyThu,
                                    PhuCapCT = p.phuCapCT,
                                    PhuCapKhac = p.phuCapKhac,
                                    DoanPhi = p.doanPhi,
                                }).ToList();
            deNghiCL.DeNghiChiLuongChiTiet = chiTietLuong;
            return View(deNghiCL);
        }
        public string GetMaxDeNghiChiLuong()
        {
            return lqThuanViet.tbl_DeNghiChiLuongs.OrderByDescending(d => d.maPhieu).Select(d => d.maPhieu).FirstOrDefault();
        }
        [HttpPost]
        public ActionResult CreateCL(FormCollection coll)
        {
            try
            {
                tbl_DeNghiChiLuongChiTiet deNghiChiTiet;
                List<tbl_DeNghiChiLuongChiTiet> lsChiTiet = new List<tbl_DeNghiChiLuongChiTiet>();
                tbl_DeNghiChiLuong deNghi = new tbl_DeNghiChiLuong();
                deNghi.maPhieu = CheckLetterDNCL("DNCL", GetMaxDeNghiChiLuong(), 3);
                deNghi.thang = Convert.ToInt32(coll.Get("thang"));
                deNghi.nam = Convert.ToInt32(coll.Get("nam"));
                deNghi.ngayLap = DateTime.Now;
                deNghi.nguoiLap = GetUser().manv;
                deNghi.tenNguoiLap = HoVaTen(GetUser().manv);
                deNghi.noiDung = coll.Get("noiDung");

                lqThuanViet.tbl_DeNghiChiLuongs.InsertOnSubmit(deNghi);

                string[] chiTiet = coll.GetValues("TenBoPhanTinhLuong");
                if (chiTiet != null)
                {
                    for (int i = 0; i < chiTiet.Count(); i++)
                    {
                        deNghiChiTiet = new tbl_DeNghiChiLuongChiTiet();
                        deNghiChiTiet.maPhieu = deNghi.maPhieu;
                        deNghiChiTiet.tenBoPhanTinhLuong = coll.GetValues("TenBoPhanTinhLuong")[i];
                        deNghiChiTiet.luongThang = Convert.ToDecimal(coll.GetValues("LuongThang")[i]);
                        deNghiChiTiet.phuCapCongTrinh = Convert.ToDecimal(coll.GetValues("PhuCapCT")[i]);
                        deNghiChiTiet.phuCapKhac = Convert.ToDecimal(coll.GetValues("PhuCapKhac")[i]);
                        deNghiChiTiet.BHXH = Convert.ToDecimal(coll.GetValues("BHXH")[i]);
                        deNghiChiTiet.thueTNCN = Convert.ToDecimal(coll.GetValues("ThueTNCN")[i]);
                        deNghiChiTiet.truyLanhTruyThu = Convert.ToDecimal(coll.GetValues("TruyLanhTruyThu")[i]);
                        deNghiChiTiet.chuyenKhoan = Convert.ToDecimal(coll.GetValues("ChuyenKhoan")[i]);
                        deNghiChiTiet.maCongTrinh = lqThuanViet.tbl_DeNghiChiLuongChiTiets.Where(d => d.tenBoPhanTinhLuong.Contains(coll.GetValues("TenBoPhanTinhLuong")[i])).Select(d => d.maCongTrinh).FirstOrDefault();
                        deNghiChiTiet.phiCongDoan = Convert.ToDouble(coll.GetValues("DoanPhi")[i]);

                        lsChiTiet.Add(deNghiChiTiet);
                    }
                    lqThuanViet.tbl_DeNghiChiLuongChiTiets.InsertAllOnSubmit(lsChiTiet);
                }
                lqThuanViet.SubmitChanges();

                return RedirectToAction("Index", "BangLuongNV");
            }
            catch
            {
                return View("Error");
            }
        }
        public string CheckLetterDNCL(string preString, string maxValue, int length)
        {
            string yearCurrent = DateTime.Now.Year.ToString().Substring(2, 2);
            string monthCurrent = DateTime.Now.Month.ToString(); // "4"
            //khi thang hien tai nho hon 9 thi cong them "0" vao
            if (Convert.ToInt32(monthCurrent) <= 9)
            {
                monthCurrent = "0" + monthCurrent;
            }
            //Khi tham so select o database la null khoi tao so dau tien
            if (String.IsNullOrEmpty(maxValue))
            {
                string ret = "1";
                while (ret.Length < length)
                {
                    ret = "0" + ret;
                }
                return preString + yearCurrent + monthCurrent + "-" + ret;
            }
            else
            {
                string preStringMax = maxValue.Substring(0, maxValue.IndexOf("-") - 4);
                string maxNumber = maxValue.Substring(maxValue.IndexOf("-") + 1);
                string monthYear = maxValue.Substring(maxValue.IndexOf("-") - 4, 4);
                string monthDb = monthYear.Substring(2, 2); //as "04"

                string stringTemp = maxNumber;
                //Khi thang trong gia tri max bang voi thang create thi cong len cho 1
                if (monthDb == monthCurrent)
                {
                    int strToInt = Convert.ToInt32(maxNumber);
                    maxNumber = Convert.ToString(strToInt + 1);
                    while (maxNumber.Length < stringTemp.Length)
                        maxNumber = "0" + maxNumber;
                }
                else //reset
                {
                    maxNumber = "1";
                    while (maxNumber.Length < stringTemp.Length)
                    {
                        maxNumber = "0" + maxNumber;
                    }
                }

                return preStringMax + yearCurrent + monthCurrent + "-" + maxNumber;
            }
        }
    }
}
