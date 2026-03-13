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
using BatDongSan.Models.VIP;
namespace BatDongSan.Controllers.TinhCong
{
    public class BangLuongThuongTetVipController : ApplicationController
    {

        private LinqVIPDataContext VipContext = new LinqVIPDataContext();
        private LinqNhanSuDataContext nhanSuContext = new LinqNhanSuDataContext();
        private readonly string MCV = "BangLuongTTVip";
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

        public ActionResult LoadIndex(string maPhongBan, string qSearch, int nam, int _page = 0)
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
            int total = VipContext.sp_BangLuongThuongTet(nam, maPhongBan, qSearch).Count();
            PagingLoaderFullController("/BangLuongThuongTetVip/Index/", total, page, "?qsearch=" + qSearch);
            ViewData["lsDanhSach"] = VipContext.sp_BangLuongThuongTet(nam, maPhongBan, qSearch).Skip(start).Take(offset).ToList();
            var ckDuyet = "khong";
            var checkDuyet = VipContext.tbl_DuyetLuongThuongTetVIPs.Where(d => d.nam == nam).FirstOrDefault();
            if (checkDuyet != null)
            {
                ckDuyet = "duyet";
            }
            ViewBag.ckDuyet = ckDuyet;
            ViewData["qSearch"] = qSearch;
            return PartialView("_LoadIndex");
        }
        public ActionResult SendBangLuongNV(int nam)
        {

            string maPhongBan = "";
            string qSearch = "";
            var listSendMails = VipContext.sp_BangLuongThuongTet(nam, maPhongBan, qSearch).ToList();
            foreach (var ds in listSendMails)
            {
                var dsMauIn = nhanSuContext.GetTable<BatDongSan.Models.HeThong.Sys_PrintTemplate>().Where(d => d.maMauIn == "MIBLTTVIP").FirstOrDefault();
                string noiDung = string.Empty;
                //Replace bang luong send mail

                noiDung = dsMauIn.html.Replace("{$nam}", Convert.ToString(nam))
                    .Replace("{$maNhanVien}", Convert.ToString(ds.maNhanVien))
                    .Replace("{$hoVaTen}", Convert.ToString(ds.hoTen))
                     .Replace("{$phongBan}", Convert.ToString(ds.tenPhongBan))
                      .Replace("{$chucDanh}", Convert.ToString(ds.tenChucVu))
                    .Replace("{$diemDanhGia}", String.Format("{0:###,##0}", ds.diemDanhGia))
                    .Replace("{$xepLoai}", Convert.ToString(ds.xepLoai))
                    .Replace("{$tongThuNhapTrongNam}", String.Format("{0:###,##0}", ds.tongThuNhapTrongNam))
                    .Replace("{$soThangLuong}", String.Format("{0:###,##0}", ds.soThangThuong))
                    .Replace("{$luongThangTrungBinh}", String.Format("{0:###,##0}", ds.luongThangTrungBinh))
                    .Replace("{$soTienThuong}", String.Format("{0:###,##0}", ds.soTienThuong))
                .Replace("{$daChuyenDot1}", String.Format("{0:###,##0}", (ds.soTienDaChuyen??0)))
                .Replace("{$conLaiPhaiChuyen}", String.Format("{0:###,##0}", (ds.conLaiPhaiChuyen??0)));
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

                MailAddress toMail = new MailAddress(ds.Email, ds.hoTen); // goi den mail
                mailInit.ToMail = toMail;
                mailInit.Body = content.ToString();
                mailInit.SendMail();
                // End code send mail
            }
            var result = new { kq = true };
            SaveActiveHistory("Send mail bảng lương thưởng tết nhân viên: năm: " + nam);
            return Json(result, JsonRequestBehavior.AllowGet);
        }
        public void XuatFileBLThuongTet(int nam, string maPhongBan, string qSearch)
        {

            var filename = "";
            var virtualPath = HttpRuntime.AppDomainAppVirtualPath;

            filename += "BangLuongThuongTetVIP_" + nam + ".xlsx";

            using (ExcelPackage package = new ExcelPackage())
            {
                //Create a sheet
                package.Workbook.Worksheets.Add("BangLuongThuongTetVIP_" + nam);
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                //Header
                //insert từ dòng nào, bao nhiêu row
                var rowFrom = 1;
                worksheet.InsertRow(rowFrom, 1);
                worksheet.Cells[1, 1].Value = "STT";
                worksheet.Cells[1, 2].Value = "Mã nhân viên";
                worksheet.Cells[1, 3].Value = "Họ tên";
                worksheet.Cells[1, 4].Value = "Họ tên không dấu";
                worksheet.Cells[1, 5].Value = "Số CMND";
                worksheet.Cells[1, 6].Value = "Số TK";
                worksheet.Cells[1, 7].Value = "Điểm đánh giá";
                worksheet.Cells[1, 8].Value = "Xếp loại";
                worksheet.Cells[1, 9].Value = "Tổng thu nhập trong năm";
                worksheet.Cells[1, 10].Value = "Số tháng thưởng";
                worksheet.Cells[1, 11].Value = "Lương tháng trung bình";
                worksheet.Cells[1, 12].Value = "Số tiền thưởng";
                worksheet.Cells[1, 13].Value = "Đã chuyển đợt 1";
                worksheet.Cells[1, 14].Value = "Còn lại phải chuyển";

                worksheet.Column(2).Width = 30;
                worksheet.Column(3).Width = 30;
                worksheet.Column(4).Width = 30;
                worksheet.Column(5).Width = 30;
                worksheet.Column(6).Width = 30;
                worksheet.Column(7).Width = 30;
                worksheet.Column(8).Width = 30;
                worksheet.Column(9).Width = 30;
                worksheet.Column(10).Width = 30;
                worksheet.Column(11).Width = 30;
                worksheet.Column(12).Width = 30;
                worksheet.Column(13).Width = 30;
                worksheet.Column(14).Width = 30;
                //// Formatting style of the header
                using (var range = worksheet.Cells[1, 1, 1, 14])
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
                var data = VipContext.sp_BangLuongThuongTet(nam, maPhongBan, qSearch).ToList();

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
                        worksheet.Cells[rowFrom, 2].Value = item.maNhanVien;
                        worksheet.Cells[rowFrom, 3].Value = item.hoTen;
                        worksheet.Cells[rowFrom, 4].Value = item.tenKhongDau;
                        worksheet.Cells[rowFrom, 5].Value = item.soCMND;

                        worksheet.Cells[rowFrom, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        worksheet.Cells[rowFrom, 6].Value = item.soTaiKhoanNganHang;

                        worksheet.Cells[rowFrom, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        worksheet.Cells[rowFrom, 7].Value = item.diemDanhGia;
                        worksheet.Cells[rowFrom, 8].Value = item.xepLoai;
                        worksheet.Cells[rowFrom, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[rowFrom, 9].Value = item.tongThuNhapTrongNam;
                        worksheet.Cells[rowFrom, 9].Style.Numberformat.Format = "#,##0.00";
                        worksheet.Cells[rowFrom, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        worksheet.Cells[rowFrom, 10].Value = item.soThangThuong;
                        worksheet.Cells[rowFrom, 10].Style.Numberformat.Format = "#,##0.00";
                        worksheet.Cells[rowFrom, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        worksheet.Cells[rowFrom, 11].Value = item.luongThangTrungBinh;
                        worksheet.Cells[rowFrom, 11].Style.Numberformat.Format = "#,##0.00";
                        worksheet.Cells[rowFrom, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        worksheet.Cells[rowFrom, 12].Value = item.soTienThuong;
                        worksheet.Cells[rowFrom, 12].Style.Numberformat.Format = "#,##0.00";
                        worksheet.Cells[rowFrom, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        worksheet.Cells[rowFrom, 13].Value = item.soTienDaChuyen;
                        worksheet.Cells[rowFrom, 13].Style.Numberformat.Format = "#,##0.00";
                        worksheet.Cells[rowFrom, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        worksheet.Cells[rowFrom, 14].Value = item.conLaiPhaiChuyen;
                        worksheet.Cells[rowFrom, 14].Style.Numberformat.Format = "#,##0.00";
                        worksheet.Cells[rowFrom, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

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
        public ActionResult DuyetLuongThuongTetVIP(int nam)
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
                var checkEx = VipContext.tbl_DuyetLuongThuongTetVIPs.Where(t => t.nam == nam).FirstOrDefault();
                if (checkEx != null)
                {
                    return Json(result, JsonRequestBehavior.AllowGet);
                }
                //End check
                // Insert Row 
                tbl_DuyetLuongThuongTetVIP tblDuyetBL = new tbl_DuyetLuongThuongTetVIP();
                tblDuyetBL.nam = nam;
                tblDuyetBL.ngayDuyet = DateTime.Now;
                tblDuyetBL.nguoiDuyet = GetUser().manv;
                VipContext.tbl_DuyetLuongThuongTetVIPs.InsertOnSubmit(tblDuyetBL);
                VipContext.SubmitChanges();
                // End Insert Row
                // Check Exist
                var list = VipContext.tbl_DuyetLuongThuongTetVIPs.Where(t => t.nam == nam).FirstOrDefault();

                if (list != null)
                {
                    result = new { kq = true };
                }
                SaveActiveHistory("Duyệt bảng lương thưởng tết VIP năm: " + nam + ". User duyệt: " + GetUser().userName);
                return Json(result, JsonRequestBehavior.AllowGet);

            }
            catch
            {
                return View();
            }

        }
        public ActionResult UpdateLuongNV(int nam)
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

                // Call Store

                nhanSuContext.tbl_NhanVienVIPs.DeleteAllOnSubmit(nhanSuContext.tbl_NhanVienVIPs.ToList());
                nhanSuContext.SubmitChanges();

                var lsNhanVien = VipContext.tbl_NhanVienVIPs.ToList().
                         Select(t => new BatDongSan.Models.NhanSu.tbl_NhanVienVIP
                         {
                             Email = t.Email,
                             GhiChu = t.GhiChu,
                             LuongCoBan = t.LuongCoBan,
                             MaNhanVien = t.MaNhanVien,
                             NgayVaoLam = t.NgayVaoLam,
                             PhuCapChucVu = t.PhuCapChucVu,
                             PhuCapCongTrinh = t.PhuCapCongTrinh,
                             PhuCapDacBiet = t.PhuCapDacBiet,
                             PhuCapThamNien = t.PhuCapThamNien,
                             PhuCapThuHut = t.PhuCapThuHut,
                             TenNhanVien = t.TenNhanVien,
                             ThoiGianCongTac = t.ThoiGianCongTac,
                             TongLuong = t.TongLuong,
                             TruyLanh = t.TruyLanh,
                             TruyThu = t.TruyThu,
                             Net = t.Net
                         }).ToList();
                nhanSuContext.tbl_NhanVienVIPs.InsertAllOnSubmit(lsNhanVien);
                nhanSuContext.SubmitChanges();

                VipContext.tbl_BangLuongThuongTets.DeleteAllOnSubmit(VipContext.tbl_BangLuongThuongTets.Where(d => d.nam == nam).ToList());
                VipContext.SubmitChanges();

                var lsLuong = nhanSuContext.sp_TinhLuongThuongVIP(nam).ToList()
                    .Select(t => new BatDongSan.Models.VIP.tbl_BangLuongThuongTet
                    {
                        maNhanVien = t.MaNhanVien,
                        hoTen = t.TenNhanVien,
                        nam = nam,
                        diemDanhGia = t.diemDanhGia,
                        xepLoai = t.xepLoai,
                        soThangThuong = (decimal)t.soThangThuong,
                        soTienDaChuyen = (decimal)t.tienDaChuyen

                    }).ToList();

                VipContext.tbl_BangLuongThuongTets.InsertAllOnSubmit(lsLuong);
                VipContext.SubmitChanges();

                VipContext.sp_TinhLuongThuongTet(nam);

                nhanSuContext.tbl_NhanVienVIPs.DeleteAllOnSubmit(nhanSuContext.tbl_NhanVienVIPs.ToList());
                nhanSuContext.SubmitChanges();
                //End store

                var result = new { kq = true };
                SaveActiveHistory("Tính lương nhân viên VIP năm: " + nam);
                return Json(result, JsonRequestBehavior.AllowGet);

            }
            catch
            {
                return View();
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



    }
}
