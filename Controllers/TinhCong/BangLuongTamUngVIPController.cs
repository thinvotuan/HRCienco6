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
    public class BangLuongTamUngVIPController : ApplicationController
    {
        private LinqVIPDataContext VipContext = new LinqVIPDataContext();
        private LinqNhanSuDataContext nhanSuContext = new LinqNhanSuDataContext();
        private readonly string MCV = "BangLuongTUVip";
        private bool? permission;
        //
        // GET: /BangLuongTamUngVIP/

        public ActionResult Index()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            thang(DateTime.Now.Month);
            nam(DateTime.Now.Year);
            return View("");

        }
        public ActionResult LoadIndex(int thang, string qSearch, int nam, int _page = 0)
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
            int total = VipContext.sp_BangLuongTamUngVip(thang, nam, qSearch).Count();
            int CkDuyet = 0;
            int CkSendMail = 0;
            var checkDuyet = VipContext.tbl_DuyetLuongTamUngVIPs.Where(d => d.nam == nam && d.thang == thang).FirstOrDefault();
           if(checkDuyet != null){
               CkDuyet = 1;
           }
           var checkSendMail = VipContext.tbl_SendBangLuongTamUngVIPs.Where(d => d.nam == nam && d.thang == thang).FirstOrDefault();
           if (checkSendMail != null) {
               CkSendMail = 1;
           }
           ViewBag.CkSendMail = CkSendMail;
           ViewBag.CkDuyet = CkDuyet;
           PagingLoaderFullController("/BangLuongTamUngVIP/Index/", total, page, "?qsearch=" + qSearch);

            ViewData["lsDanhSach"] = VipContext.sp_BangLuongTamUngVip(thang,nam, qSearch).Skip(start).Take(offset).ToList();

            ViewData["qSearch"] = qSearch;
            return PartialView("_LoadIndex");
        }
        public ActionResult UpdateLuongTUNV(int nam, int thang)
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
                //End store

                var result = new { kq = true };
                SaveActiveHistory("Tính lương tạm ứng VIP tháng: " + thang + " năm: " + nam);
                return Json(result, JsonRequestBehavior.AllowGet);

            }
            catch
            {
                return View();
            }

        }
        public ActionResult DuyetLuongTamUngVip(int thang, int nam)
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
                var checkEx = VipContext.tbl_DuyetLuongTamUngVIPs.Where(t => t.nam == nam && t.thang == thang).FirstOrDefault();
                if (checkEx != null)
                {
                    return Json(result, JsonRequestBehavior.AllowGet);
                }
                //End check
                // Insert Row 
                tbl_DuyetLuongTamUngVIP tblDuyetBL = new tbl_DuyetLuongTamUngVIP();
                tblDuyetBL.nam = nam;
                tblDuyetBL.thang = thang;
                tblDuyetBL.ngayDuyet = DateTime.Now;
                tblDuyetBL.nguoiDuyet = GetUser().manv;
                VipContext.tbl_DuyetLuongTamUngVIPs.InsertOnSubmit(tblDuyetBL);
                VipContext.SubmitChanges();
                // End Insert Row
                // Check Exist
                var list = VipContext.tbl_DuyetLuongTamUngVIPs.Where(t => t.nam == nam && t.thang == thang).FirstOrDefault();

                if (list != null)
                {
                    result = new { kq = true };
                }
                SaveActiveHistory("Duyệt lương tạm ứng VIP tháng: " + thang + " năm: " + nam + ". User duyệt: " + GetUser().userName);
                return Json(result, JsonRequestBehavior.AllowGet);

            }
            catch
            {
                return View();
            }

        }
        public void XuatFileBLNN(int thang, int nam)
        {

            var filename = "";
            var virtualPath = HttpRuntime.AppDomainAppVirtualPath;

            filename += "BangLuongTamUngVIP_" + nam + "_" + thang + ".xlsx";

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
                worksheet.Cells[1, 4].Value = "Số CMND";
                worksheet.Cells[1, 5].Value = "Tổng tạm ứng";
                worksheet.Cells[1, 6].Value = "Đã chuyển đợt 1";
                worksheet.Cells[1, 7].Value = "Còn lại phải chuyển";

                worksheet.Column(2).Width = 35;
                worksheet.Column(3).Width = 35;
                worksheet.Column(4).Width = 20;
                worksheet.Column(5).Width = 20;
                worksheet.Column(6).Width = 20;
                worksheet.Column(7).Width = 20;
                //// Formatting style of the header
                using (var range = worksheet.Cells[1, 1, 1, 7])
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
                var data = VipContext.sp_BangLuongNhanVienTamUng(thang, nam).ToList();

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
                        worksheet.Cells[rowFrom, 2].Value = item.tenKhongDau;
                        worksheet.Cells[rowFrom, 3].Value = item.soTaiKhoanNganHang;
                        worksheet.Cells[rowFrom, 4].Value = item.soCMND;
                        worksheet.Cells[rowFrom, 5].Value = item.tongLuong;
                        worksheet.Cells[rowFrom, 6].Value = item.soTienChuyenDot1;
                        worksheet.Cells[rowFrom, 6].Style.Numberformat.Format = "#,##0.000";
                        worksheet.Cells[rowFrom, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        worksheet.Cells[rowFrom, 7].Value = item.soTienConLai;
                        worksheet.Cells[rowFrom, 7].Style.Numberformat.Format = "#,##0.000";
                        worksheet.Cells[rowFrom, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

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


        public ActionResult SendBangLuongTamUngVip(int thang, int nam)
        {

            string qSearch = "";
            var list = VipContext.tbl_DuyetLuongTamUngVIPs.Where(t => t.nam == nam && t.thang == thang).FirstOrDefault();
            var result = new { kq = false };
            if (list == null)
            {
                return Json(result, JsonRequestBehavior.AllowGet);
            }

            var listSendMails = VipContext.sp_BangLuongTamUngVip(thang,nam, qSearch).ToList();
            foreach (var ds in listSendMails)
            {
                if (ds.Email != null && ds.Email != "")
                {
                    var dsMauIn = nhanSuContext.GetTable<BatDongSan.Models.HeThong.Sys_PrintTemplate>().Where(d => d.maMauIn == "MIBLTamUngVip").FirstOrDefault();
                    string noiDung = string.Empty;
                    //Replace bang luong send mail
                    noiDung = dsMauIn.html.Replace("{$thang}", Convert.ToString(thang))
                        .Replace("{$nam}", Convert.ToString(nam))
                        .Replace("{$maNhanVien}", ds.maNhanVien)
                        .Replace("{$hoVaTen}", ds.hoTen)
                        .Replace("{$tongLuong}", String.Format("{0:###,##0}", ds.tongLuong))
                        .Replace("{$soTienChuyenDot1}", String.Format("{0:###,##0}", ds.soTienChuyenDot1))
                        .Replace("{$soTienConLai}", String.Format("{0:###,##0}", ds.soTienConLai));

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
                    //string[] array01 = ds.email.ToLower().Split('@');
                    //string string2 = ConfigurationManager.AppSettings["OnlySend"]; //get domain from config files
                    //string[] array1 = string2.Split(',');
                    // bool EmailofThuanViet;
                    //EmailofThuanViet = array1.Contains(array01[1]);
                    // if (emailNV == "" || emailNV == null || EmailofThuanViet == false)
                    // {
                    //    return false;
                    // }
                    MailAddress toMail = new MailAddress(ds.Email, ds.hoTen); // goi den mail
                    mailInit.ToMail = toMail;
                    mailInit.Body = content.ToString();
                    mailInit.SendMail();
                    // End code send mail
                }
            }
            result = new { kq = true };
            SaveActiveHistory("Send mail bảng lương tạm ứng VIP: " + thang + " năm: " + nam + ". User send: " + GetUser().manv);
            var tblSendMail = VipContext.tbl_SendBangLuongTamUngVIPs.Where(d => d.nam == nam && d.thang == thang).FirstOrDefault();
            if (tblSendMail == null)
            {
                // Insert Row 
                tbl_SendBangLuongTamUngVIP tblDuyetBL = new tbl_SendBangLuongTamUngVIP();
                tblDuyetBL.nam = nam;
                tblDuyetBL.thang = thang;
                tblDuyetBL.ngaySend = DateTime.Now;
                tblDuyetBL.nguoiSend = GetUser().manv;
                VipContext.tbl_SendBangLuongTamUngVIPs.InsertOnSubmit(tblDuyetBL);
                VipContext.SubmitChanges();
                // End Insert Row
            }

            return Json(result, JsonRequestBehavior.AllowGet);
        }

        private void thang(int value)
        {
            Dictionary<int, string> dics = new Dictionary<int, string>();
            for (int i = 1; i <= 12; i++)
            {
                dics[i] = i.ToString();
            }
            ViewData["lsthang"] = new SelectList(dics, "Key", "Value", value);

        }
        private void nam(int value)
        {
            Dictionary<int, string> dics = new Dictionary<int, string>();
            for (int i = DateTime.Now.Year - 5; i < DateTime.Now.Year + 2; i++)
            {
                dics[i] = i.ToString();
            }
            ViewData["lsnam"] = new SelectList(dics, "Key", "Value", value);

        }
    }
}
