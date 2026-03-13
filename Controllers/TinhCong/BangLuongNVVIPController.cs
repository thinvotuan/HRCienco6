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
    public class BangLuongNVVIPController : ApplicationController
    {

        private LinqNhanSuDataContext nhanSuContext = new LinqNhanSuDataContext();
        private LinqVIPDataContext vipContext = new LinqVIPDataContext();
        private readonly string MCV = "BangLuongNVVIP";
        private IList<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan> phongBans;
        private StringBuilder buildTree;
        private bool? permission;
        public ActionResult XemBangLuong()
        {
            #region Role user
            permission = GetPermission("XemBangLuong", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            thang(DateTime.Now.Month);
            nam(DateTime.Now.Year);
            return View("");
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
            var dsMauIn = nhanSuContext.GetTable<BatDongSan.Models.HeThong.Sys_PrintTemplate>().Where(d => d.maMauIn == "MIBLNVVIP").FirstOrDefault();
            string noiDung = string.Empty;
            var ds = vipContext.tbl_BangLuongVIPs.Where(t => t.maNhanVien == maNhanVien && t.thang.ToString() == thang && nam == t.nam.ToString()).FirstOrDefault();
            var phuCapPhatSinh = vipContext.tbl_NhanVienVIPs.Where(d => d.MaNhanVien == maNhanVien).Select(d => d.PhuCapPhatSinh).FirstOrDefault() ??0;
            ViewData["chiTiet"] = dsMauIn;
            if (dsMauIn != null)
            {
                noiDung = dsMauIn.html.Replace("{$thang}", thang)
                    .Replace("{$nam}", nam)
                    .Replace("{$maNhanVien}", ds.maNhanVien)
                    .Replace("{$HoVaTen}", ds.tenNhanVien)
                    .Replace("{$tongLuongVaPhuCap}", String.Format("{0:###,##0}", ds.tongLuong))
                    .Replace("{$luongCoBan}", String.Format("{0:###,##0}", ds.luongCoban))
                    .Replace("{$phuCapThamNien}", String.Format("{0:###,##0}", ds.phuCapThamNien))
                    .Replace("{$ngayCongTac}", String.Format("{0: dd/MM/yyyy}",ds.ngayVaoLam))
                    .Replace("{$ngayCongTinhLuong}", Convert.ToString(ds.ngayTinhLuong))
                    .Replace("{$thoiGianCongTac}", Convert.ToString(ds.thoiGianCongTac))
                    .Replace("{$thanhTienPhuCapThamNien}", String.Format("{0:###,##0}", ds.phuCapThamNien))
                    .Replace("{$phuCapChucVu}", String.Format("{0:###,##0}", ds.phuCapChucVu))
                    .Replace("{$phuCapCongTrinh}", String.Format("{0:###,##0}", ds.phuCapCongTrinh))
                    .Replace("{$phuCapThuHut}", String.Format("{0:###,##0}", ds.phuCapThuHut))
                    .Replace("{$phuCapDacBiet}", String.Format("{0:###,##0}", ds.phuCapDacBiet))
                    .Replace("{$phuCapPhatSinh}", String.Format("{0:###,##0}", phuCapPhatSinh))
                    .Replace("{$cacKhoanGiamTru}", String.Format("{0:###,##0}", ((ds.baoHiemXH ?? 0) + (ds.thueTNCN ?? 0))))
                    .Replace("{$khauTruBHXH}", String.Format("{0:###,##0}", ds.baoHiemXH))
                    .Replace("{$thueTNCNThang}", String.Format("{0:###,##0}", ds.thueTNCN))
                    .Replace("{$ngayCongChuan}", String.Format("{0:###,##0}", ds.ngayCongChuan))
                    .Replace("{$ngayCongTinhLuong}", String.Format("{0:###,##0}", ds.ngayTinhLuong))
                     .Replace("{$chiTietThucNhanKyNay}", String.Format("{0:###,##0}", ds.conLaiPhaiChuyen))
                    .Replace("{$luongNgayCongERP}", String.Format("{0:###,##0}", ds.luongThang))
                    .Replace("{$luongChuyenDot1}", String.Format("{0:###,##0}", ds.luongDaChuyenDot1))
                    .Replace("{$truyThuKyTruoc}", String.Format("{0:###,##0}", ds.truyThu))
                    .Replace("{$truyLanhKyTruoc}", String.Format("{0:###,##0}", ds.truyLanh));
                   

            }
            ViewBag.NoiDung = noiDung;
            // return PartialView("_ViewChiTietLuong");
            return PartialView("_ViewChiTietLuongTemplate");
        }
        public void XuatFileBLNV(int thang, int nam)
        {
            var filename = "";
            var virtualPath = HttpRuntime.AppDomainAppVirtualPath;
            var fileStream = new FileStream(System.Web.HttpContext.Current.Server.MapPath(virtualPath + @"\Content\Report\ReportTemplate.xls"), FileMode.Open, FileAccess.Read);

            var workbook = new HSSFWorkbook(fileStream, true);
            filename += "BangLuongNVVIP_" + thang + "_" + nam + ".xls";


            var sheet = workbook.GetSheet("danhsachnhanvien");

            /*style title start*/
            //tạo font cho các title
            //font tiêu đề 
            HSSFFont hFontTieuDe = (HSSFFont)workbook.CreateFont();
            hFontTieuDe.FontHeightInPoints = 18;
            hFontTieuDe.Boldweight = 100 * 10;
            hFontTieuDe.FontName = "Times New Roman";
            hFontTieuDe.Color = HSSFColor.BLUE.index;
            HSSFFont hFontTieuDe2 = (HSSFFont)workbook.CreateFont();
            hFontTieuDe2.FontHeightInPoints = 15;
            hFontTieuDe2.Boldweight = 100 * 10;
            hFontTieuDe2.FontName = "Times New Roman";
            hFontTieuDe2.Color = HSSFColor.BLACK.index;

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
            var styleTitle1 = workbook.CreateCellStyle();
            styleTitle1.SetFont(hFontTieuDe2);
            styleTitle1.Alignment = HorizontalAlignment.CENTER;

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


            Row rowC = null;
            //Khai báo row đầu tiên
            int firstRowNumber = 1;

            string rowtitle = "BẢNG LƯƠNG NHÂN VIÊN VIP";
            var titleCell = HSSFCellUtil.CreateCell(sheet.CreateRow(firstRowNumber), 3, rowtitle.ToUpper());
            titleCell.CellStyle = styleTitle;
            firstRowNumber = firstRowNumber + 2;
            string rowtitle1 = "Tháng: " + thang + " năm: " + nam;
            var titleCell1 = HSSFCellUtil.CreateCell(sheet.CreateRow(firstRowNumber), 3, rowtitle1);
            titleCell1.CellStyle = styleTitle1;

            firstRowNumber = firstRowNumber + 2;

            var list1 = new List<string>();
            list1.Add("STT");
            list1.Add("Mã nhân viên");
            list1.Add("Họ Tên");
            list1.Add("Họ Tên Không Dấu");
            list1.Add("Số CMND");
            list1.Add("Số Tài Khoản");
           
            list1.Add("Tổng lương");
            list1.Add("Lương Cơ Bản");
            list1.Add("Phụ Cấp Thâm Niên");
            list1.Add("Phụ Cấp Chức Vụ");
            
            list1.Add("Phụ cấp công trình");

            list1.Add("Phụ Cấp Thu Hút");

            list1.Add("Phụ Cấp Đặc Biệt");
            
            list1.Add("Ngày công chuẩn");
            list1.Add("Tổng Công");
            list1.Add("Lương Tháng");
            list1.Add("BHXH");
            list1.Add("Thuế TNCN");
            list1.Add("Truy Lãnh");
            list1.Add("Truy Thu");
            list1.Add("Đã Chuyển Đợt 1");
            list1.Add("Còn Lại Phải Chuyển");
            list1.Add("Ghi Chú");

            //Start row 13
            var headerRow = sheet.CreateRow(6);
            ReportHelperExcel.CreateHeaderRow(headerRow, 0, styleheadedColumnTable, list1);

            //Create header end

            var idRowStart = 6;
            var datas = vipContext.sp_BangLuongNhanVienVIP(nam, thang,null).ToList();
            //#region
            if (datas != null && datas.Count > 0)
            {
                var stt = 0;
                int dem = 0;
                //Giai đoạn
                foreach (var item1 in datas)
                {
                    dem = 0;
                    stt++;
                    idRowStart++;
                    rowC = sheet.CreateRow(idRowStart);
                    ReportHelperExcel.SetAlignment(rowC, dem++, stt.ToString(), hStyleConCenter);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.maNhanVien, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.tenNhanVien, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.tenKhongDau, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.soCMND, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.soTaiKhoanNganHang, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item1.tongLuong), hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item1.luongCoban), hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item1.phuCapThamNien), hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item1.phuCapChucVu), hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item1.phuCapCongTrinh), hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item1.phuCapThuHut), hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item1.phuCapDacBiet), hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item1.ngayCongChuan), hStyleConRight);
                    
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item1.ngayCong), hStyleConRight);
                    
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item1.luongThang), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item1.baoHiemXH), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item1.thueTNCN), hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item1.truyLanh), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item1.truyThu), hStyleConRight);
                   
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item1.luongDaChuyenDot1), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item1.conLaiPhaiChuyen), hStyleConRight);
                    
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item1.ghiChu), hStyleConRight);
                  
                }

                sheet.SetColumnWidth(0, 8 * 210);
                sheet.SetColumnWidth(1, 30 * 210);
                sheet.SetColumnWidth(2, 30 * 210);
                sheet.SetColumnWidth(3, 30 * 210);
                sheet.SetColumnWidth(4, 30 * 210);
                sheet.SetColumnWidth(5, 30 * 210);
                sheet.SetColumnWidth(6, 30 * 210);
                sheet.SetColumnWidth(7, 30 * 210);
                sheet.SetColumnWidth(8, 30 * 210);
                sheet.SetColumnWidth(9, 30 * 210);
                sheet.SetColumnWidth(10, 30 * 210);
                sheet.SetColumnWidth(11, 30 * 210);

                sheet.SetColumnWidth(12, 30 * 210);
                sheet.SetColumnWidth(13, 30 * 210);
                sheet.SetColumnWidth(14, 30 * 210);
                sheet.SetColumnWidth(15, 30 * 210);
                sheet.SetColumnWidth(16, 30 * 210);
                sheet.SetColumnWidth(17, 30 * 210);
                sheet.SetColumnWidth(18, 30 * 210);
                sheet.SetColumnWidth(19, 30 * 210);
                sheet.SetColumnWidth(20, 30 * 210);
                sheet.SetColumnWidth(21, 30 * 210);
                sheet.SetColumnWidth(22, 30 * 210);
                sheet.SetColumnWidth(23, 30 * 210);
                sheet.SetColumnWidth(24, 30 * 210);
                sheet.SetColumnWidth(25, 30 * 210);
            }
            else
            {

                sheet.SetColumnWidth(0, 8 * 210);
                sheet.SetColumnWidth(1, 30 * 210);
                sheet.SetColumnWidth(2, 30 * 210);
                sheet.SetColumnWidth(3, 30 * 210);
                sheet.SetColumnWidth(4, 30 * 210);
                sheet.SetColumnWidth(5, 30 * 210);
                sheet.SetColumnWidth(6, 30 * 210);
                sheet.SetColumnWidth(7, 30 * 210);
                sheet.SetColumnWidth(8, 30 * 210);
                sheet.SetColumnWidth(9, 30 * 210);
                sheet.SetColumnWidth(10, 30 * 210);
                sheet.SetColumnWidth(11, 30 * 210);

                sheet.SetColumnWidth(12, 30 * 210);
                sheet.SetColumnWidth(13, 30 * 210);
                sheet.SetColumnWidth(14, 30 * 210);
                sheet.SetColumnWidth(15, 30 * 210);
                sheet.SetColumnWidth(16, 30 * 210);
                sheet.SetColumnWidth(17, 30 * 210);
                sheet.SetColumnWidth(18, 30 * 210);
                sheet.SetColumnWidth(19, 30 * 210);
                sheet.SetColumnWidth(20, 30 * 210);
                sheet.SetColumnWidth(21, 30 * 210);
                sheet.SetColumnWidth(22, 30 * 210);
                sheet.SetColumnWidth(23, 30 * 210);
                sheet.SetColumnWidth(24, 30 * 210);
                sheet.SetColumnWidth(25, 30 * 210);

            }

            var stream = new MemoryStream();
            workbook.Write(stream);

            Response.ContentType = "application/vnd.ms-excel";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));
            Response.Clear();

            Response.BinaryWrite(stream.GetBuffer());
            Response.End();

        }

      

        public ActionResult LoadXemBangLuong(int nam, int thang, string qSearch,  int _page = 0)
        {
            #region Role user
            permission = GetPermission("XemBangLuong", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            string maNhanVien = GetUser().manv;
            //int page = _page == 0 ? 1 : _page;
            //int pIndex = page;
            int total = vipContext.sp_BangLuongNhanVienVIP(nam, thang, qSearch).Count();
            //PagingLoaderController("/BangLuongNVVIP/XemBangLuong/", total, page,"");
            ViewData["lsDanhSach"] = vipContext.sp_BangLuongNhanVienVIP(nam, thang, qSearch).ToList();
            var kqCheck = 0;
            var checkEx = vipContext.tbl_DuyetLuongVIPs.Where(t => t.nam == nam && t.thang == thang).FirstOrDefault();
            if (checkEx != null)
            {
                kqCheck = 1;
            }
            ViewData["kqCheck"] = kqCheck;
            var kqSend = 0;
            var checkSend = vipContext.tbl_SendBangLuongVIPs.Where(d => d.nam == nam && d.thang == thang).FirstOrDefault();
            if(checkSend != null){
                kqSend = 1;
            }
            ViewData["kqSend"] = kqSend;
            ViewBag.totalRow = total;
            return PartialView("_LoadXemBangLuong");
        }
        public ActionResult UpdateLuongNV(int thang, int nam)
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
                var result = new { kq = false };
                var checkEx = vipContext.tbl_DuyetLuongVIPs.Where(t => t.nam == nam && t.thang == thang).FirstOrDefault();
                if (checkEx != null)
                {
                    return Json(result, JsonRequestBehavior.AllowGet);
                }
                int month = Convert.ToInt32(thang);
                int year = Convert.ToInt32(nam);

                var lsNhanVien = vipContext.tbl_NhanVienVIPs.ToList().
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
                               Net = t.Net,
                               khongChamCong = t.khongChamCong ,
                               nganHang =t.nganHang ,
                               soTaiKhoan =t.soTaiKhoan 
                               
                           }).ToList();
                nhanSuContext.tbl_NhanVienVIPs.InsertAllOnSubmit(lsNhanVien);
                nhanSuContext.SubmitChanges();
                nhanSuContext.sp_TinhLuongVIP(month, year);

                vipContext.tbl_BangLuongVIPs.DeleteAllOnSubmit(vipContext.tbl_BangLuongVIPs.Where(d => d.thang == month && d.nam == year).ToList());
                vipContext.SubmitChanges();

                var lsLuong = nhanSuContext.tbl_BangLuongVIPs.Where(d => d.thang == month && d.nam == year).ToList()
                    .Select(t => new BatDongSan.Models.VIP.tbl_BangLuongVIP
                    {
                        maNhanVien = t.maNhanVien,
                        baoHiemXH = t.baoHiemXH,
                        baoHiemXHTV = t.baoHiemXHTV,
                        conLaiPhaiChuyen = t.conLaiPhaiChuyen,
                        email = t.email,
                        ghiChu = t.ghiChu,
                        luongCoban = t.luongCoban,
                        luongDaChuyenDot1 = t.luongDaChuyenDot1,
                        luongHopDongLaoDong = t.luongHopDongLaoDong,
                        luongThang = t.luongThang,
                        nam = t.nam,
                        ngayCong = t.ngayCong,
                        ngayCongChuan = t.ngayCongChuan,
                        ngayLuyKeThangTruoc = t.ngayLuyKeThangTruoc,
                        ngayNghiBu = t.ngayNghiBu,
                        ngayNghiLeTet = t.ngayNghiLeTet,
                        ngayNghiPhep = t.ngayNghiPhep,
                        ngayTinhLuong = t.ngayTinhLuong,
                        ngayVaoLam = t.ngayVaoLam,
                        phuCapChucVu = t.phuCapChucVu,
                        phuCapCongTrinh = t.phuCapCongTrinh,
                        phuCapDacBiet = t.phuCapDacBiet,
                        phuCapThamNien = t.phuCapThamNien,
                        phuCapThuHut = t.phuCapThuHut,
                        soCMND = t.soCMND,
                        soTaiKhoanNganHang = t.soTaiKhoanNganHang,
                        tenKhongDau = t.tenKhongDau,
                        tenNhanVien = t.tenNhanVien,
                        thang = t.thang,
                        thoiGianCongTac = t.thoiGianCongTac,
                        thueTNCN = t.thueTNCN,
                        thuTu = t.thuTu,
                        tongChiPhiTVPhaiTra = t.tongChiPhiTVPhaiTra,
                        tongLuong = t.tongLuong,
                        truyLanh = t.truyLanh,
                        truyThu = t.truyThu
                    }).ToList();

                vipContext.tbl_BangLuongVIPs.InsertAllOnSubmit(lsLuong);
                vipContext.SubmitChanges();
                nhanSuContext.tbl_BangLuongVIPs.DeleteAllOnSubmit(nhanSuContext.tbl_BangLuongVIPs.ToList());
                nhanSuContext.tbl_NhanVienVIPs.DeleteAllOnSubmit(nhanSuContext.tbl_NhanVienVIPs.ToList());
                nhanSuContext.SubmitChanges();
                //Check 
              
                //End check
                //var list = vipContext.sp_TinhLuongNhanVienVIP(thang, nam);
                result = new { kq = true };
                SaveActiveHistory("Get dữ liệu tính lương, công chuẩn: " + thang + " năm: " + nam);
                return Json(result, JsonRequestBehavior.AllowGet);

            }
            catch
            {
                return View();
            }

        }
        public ActionResult TinhLuongNVNew(int thang, int nam)
        {


            try
            {
                var result = new { kq = true };
                #region Role user
                permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
                if (!permission.HasValue)
                    result = new { kq = false };
                if (!permission.Value)
                    result = new { kq = false };
                #endregion
                vipContext.sp_NS_TinhLuongVIP(thang, nam);
                 result = new { kq = true };
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
                var checkEx = vipContext.tbl_DuyetLuongVIPs.Where(t => t.nam == nam && t.thang == thang).FirstOrDefault();
                if (checkEx != null)
                {
                    return Json(result, JsonRequestBehavior.AllowGet);
                }
                //End check
                // Insert Row 
                tbl_DuyetLuongVIP tblDuyetBL = new tbl_DuyetLuongVIP();
                tblDuyetBL.nam = nam;
                tblDuyetBL.thang = thang;
                tblDuyetBL.ngayDuyet = DateTime.Now;
                tblDuyetBL.nguoiDuyet = GetUser().manv;
                vipContext.tbl_DuyetLuongVIPs.InsertOnSubmit(tblDuyetBL);
                vipContext.SubmitChanges();
                // End Insert Row
                // Check Exist
                var list = vipContext.tbl_DuyetLuongVIPs.Where(t => t.nam == nam && t.thang == thang).FirstOrDefault();

                if (list != null)
                {
                    result = new { kq = true };
                }
                SaveActiveHistory("Duyệt lương nhân viên VIP tháng: " + thang + " năm: " + nam);
                return Json(result, JsonRequestBehavior.AllowGet);

            }
            catch
            {
                return View();
            }

        }
        public ActionResult SendBangLuongNV(int thang, int nam)
        {

            string qSearch = "";
            var list = vipContext.tbl_DuyetLuongVIPs.Where(t => t.nam == nam && t.thang == thang).FirstOrDefault();
            var result = new { kq = false };
            if (list == null)
            {
                return Json(result, JsonRequestBehavior.AllowGet);
            }

            var listSendMails = vipContext.sp_BangLuongNhanVienVIP(nam, thang, qSearch).ToList();
            foreach (var ds in listSendMails)
            {
               // var dsMauIn = nhanSuContext.GetTable<BatDongSan.Models.HeThong.Sys_PrintTemplate>().Where(d => d.maMauIn == "MIBLNVVIP").FirstOrDefault();
               // var phuCapPhatSinh = vipContext.tbl_NhanVienVIPs.Where(d => d.MaNhanVien == ds.maNhanVien).Select(d => d.PhuCapPhatSinh).FirstOrDefault() ?? 0;
                //string noiDung = string.Empty;
                ////Replace bang luong send mail
                //noiDung = dsMauIn.html.Replace("{$thang}", Convert.ToString(thang))
                //    .Replace("{$nam}", Convert.ToString(nam))
                //    .Replace("{$maNhanVien}", ds.maNhanVien)
                //    .Replace("{$HoVaTen}", ds.tenNhanVien)
                //    .Replace("{$tongLuongVaPhuCap}", String.Format("{0:###,##0}", ds.tongLuong))
                //    .Replace("{$luongCoBan}", String.Format("{0:###,##0}", ds.luongCoban))
                //    .Replace("{$phuCapThamNien}", String.Format("{0:###,##0}", ds.phuCapThamNien))
                //    .Replace("{$ngayCongTac}", String.Format("{0: dd/MM/yyyy}", ds.ngayVaoLam))
                //    .Replace("{$ngayCongTinhLuong}", Convert.ToString(ds.ngayTinhLuong))
                //    .Replace("{$thoiGianCongTac}", Convert.ToString(ds.thoiGianCongTac))
                //    .Replace("{$thanhTienPhuCapThamNien}", String.Format("{0:###,##0}", ds.phuCapThamNien))
                //    .Replace("{$phuCapChucVu}", String.Format("{0:###,##0}", ds.phuCapChucVu))
                //    .Replace("{$phuCapCongTrinh}", String.Format("{0:###,##0}", ds.phuCapCongTrinh))
                //    .Replace("{$phuCapThuHut}", String.Format("{0:###,##0}", ds.phuCapThuHut))
                //    .Replace("{$phuCapDacBiet}", String.Format("{0:###,##0}", ds.phuCapDacBiet))
                //    .Replace("{$phuCapPhatSinh}", String.Format("{0:###,##0}", phuCapPhatSinh))
                //    .Replace("{$cacKhoanGiamTru}", String.Format("{0:###,##0}", ((ds.baoHiemXH ?? 0) + (ds.thueTNCN ?? 0))))
                //    .Replace("{$khauTruBHXH}", String.Format("{0:###,##0}", ds.baoHiemXH))
                //    .Replace("{$thueTNCNThang}", String.Format("{0:###,##0}", ds.thueTNCN))
                //    .Replace("{$ngayCongChuan}", String.Format("{0:###,##0}", ds.ngayCongChuan))
                //    .Replace("{$ngayCongTinhLuong}", String.Format("{0:###,##0}", ds.ngayTinhLuong))
                //     .Replace("{$chiTietThucNhanKyNay}", String.Format("{0:###,##0}", ds.conLaiPhaiChuyen))
                //    .Replace("{$luongNgayCongERP}", String.Format("{0:###,##0}", ds.luongThang))
                //    .Replace("{$luongChuyenDot1}", String.Format("{0:###,##0}", ds.luongDaChuyenDot1))
                //    .Replace("{$truyThuKyTruoc}", String.Format("{0:###,##0}", ds.truyThu))
                //    .Replace("{$truyLanhKyTruoc}", String.Format("{0:###,##0}", ds.truyLanh));

                ////End
                string appPath = string.Format("{0}://{1}{2}{3}",
                 Request.Url.Scheme,
                 Request.Url.Host,
                  (Request.Url.Port == 80) ? string.Empty : ":" + Request.Url.Port,
                 "/BangLuongNVVIP/XemBangLuongNVVIP");
                // Code send mail
                MailHelper mailInit = new MailHelper(); // lay cac tham so trong webconfig
                System.Text.StringBuilder content = new System.Text.StringBuilder();

                content.Append("<h3>Email từ hệ thống nhân sự</h3>");

                content.Append("<p>Xin chào: " + ds.tenNhanVien + " !</p>");
                
                content.Append("<p>Vì lý do bảo mật, Anh/Chị vui lòng bấm vào link sau để xem bảng lương: <strong>" + appPath + "</strong>");
                content.Append("<p>Link trên chỉ hoạt động vào khung giờ: 9am đến 17pm, từ thứ 2 đến thứ 6.</p>");
                //Content
                //content.Append(noiDung);

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
                MailAddress toMail = new MailAddress(ds.email, ds.tenNhanVien); // goi den mail
                mailInit.ToMail = toMail;
                mailInit.Body = content.ToString();
                mailInit.SendMail();
                // End code send mail
            }
            result = new { kq = true };
            SaveActiveHistory("Send mail bảng lương nhân viên VIP: " + thang + " năm: " + nam);
            var tblSendMail = vipContext.tbl_SendBangLuongVIPs.Where(d => d.nam == nam && d.thang == thang).FirstOrDefault();
            if (tblSendMail != null)
            {
                vipContext.tbl_SendBangLuongVIPs.DeleteOnSubmit(tblSendMail);
                vipContext.SubmitChanges();
            }
            // Insert Row 
            tbl_SendBangLuongVIP tblDuyetBL = new tbl_SendBangLuongVIP();
            tblDuyetBL.nam = nam;
            tblDuyetBL.thang = thang;
            tblDuyetBL.ngaySend = DateTime.Now;
            tblDuyetBL.nguoiSend = GetUser().manv;
            vipContext.tbl_SendBangLuongVIPs.InsertOnSubmit(tblDuyetBL);
            vipContext.SubmitChanges();
            // End Insert Row
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


        public ActionResult XemBangLuongNVVIP()
        {
            #region Role user
            permission = GetPermission("XemBangLuongNVVIP", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            thang(DateTime.Now.Month);
            nam(DateTime.Now.Year);
            return View("");
        }
        public ActionResult LoadXemBangLuongNVVIP(int nam, string qSearch, int _page = 0)
        {
            #region Role user
            permission = GetPermission("XemBangLuongNVVIP", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            string maNhanVien = GetUser().manv;
            BangLuongNVVIPModel blVip = new BangLuongNVVIPModel();
            var lstThang = vipContext.tbl_DuyetLuongVIPs.Where(d => d.nam == nam).Select(d => d.thang).ToList();
            blVip.blVips = (from bl in vipContext.tbl_BangLuongVIPs
                     
                     where (bl.maNhanVien == maNhanVien && bl.nam == nam && lstThang.Contains(bl.thang))
                     select new BLNVVIP
                     {
                         maNhanVien = bl.maNhanVien,
                         tenNhanVien = bl.tenNhanVien,
                         thang = bl.thang,
                         tongLuong = bl.tongLuong,
                         ngayCongChuan = bl.ngayCongChuan,
                         ngayTinhLuong = bl.ngayTinhLuong,
                         thueTNCN = bl.thueTNCN,
                         baoHiemXH = bl.baoHiemXH,
                         conLaiPhaiChuyen = bl.conLaiPhaiChuyen

                     }).Distinct().ToList();
            
            return PartialView("_LoadXemBangLuongNVVIP",blVip);
        }
        public ActionResult ViewChiTietLuongNVVIP(string thang, string nam)
        {
            #region Role user
            permission = GetPermission("XemBangLuongNVVIP", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            string maNhanVien = GetUser().manv;
            var dsMauIn = nhanSuContext.GetTable<BatDongSan.Models.HeThong.Sys_PrintTemplate>().Where(d => d.maMauIn == "MIBLNVVIP").FirstOrDefault();
            string noiDung = string.Empty;
            var ds = vipContext.tbl_BangLuongVIPs.Where(t => t.maNhanVien == maNhanVien && t.thang.ToString() == thang && nam == t.nam.ToString()).FirstOrDefault();
            var phuCapPhatSinh = vipContext.tbl_NhanVienVIPs.Where(d => d.MaNhanVien == maNhanVien).Select(d => d.PhuCapPhatSinh).FirstOrDefault() ?? 0;
            ViewData["chiTiet"] = dsMauIn;
            if (dsMauIn != null)
            {
                noiDung = dsMauIn.html.Replace("{$thang}", thang)
                    .Replace("{$nam}", nam)
                    .Replace("{$maNhanVien}", ds.maNhanVien)
                    .Replace("{$HoVaTen}", ds.tenNhanVien)
                    .Replace("{$tongLuongVaPhuCap}", String.Format("{0:###,##0}", ds.tongLuong))
                    .Replace("{$luongCoBan}", String.Format("{0:###,##0}", ds.luongCoban))
                    .Replace("{$phuCapThamNien}", String.Format("{0:###,##0}", ds.phuCapThamNien))
                    .Replace("{$ngayCongTac}", String.Format("{0: dd/MM/yyyy}", ds.ngayVaoLam))
                    .Replace("{$ngayCongTinhLuong}", Convert.ToString(ds.ngayTinhLuong))
                    .Replace("{$thoiGianCongTac}", Convert.ToString(ds.thoiGianCongTac))
                    .Replace("{$thanhTienPhuCapThamNien}", String.Format("{0:###,##0}", ds.phuCapThamNien))
                    .Replace("{$phuCapChucVu}", String.Format("{0:###,##0}", ds.phuCapChucVu))
                    .Replace("{$phuCapCongTrinh}", String.Format("{0:###,##0}", ds.phuCapCongTrinh))
                    .Replace("{$phuCapThuHut}", String.Format("{0:###,##0}", ds.phuCapThuHut))
                    .Replace("{$phuCapDacBiet}", String.Format("{0:###,##0}", ds.phuCapDacBiet))
                    .Replace("{$phuCapPhatSinh}", String.Format("{0:###,##0}", phuCapPhatSinh))
                    .Replace("{$cacKhoanGiamTru}", String.Format("{0:###,##0}", ((ds.baoHiemXH ?? 0) + (ds.thueTNCN ?? 0))))
                    .Replace("{$khauTruBHXH}", String.Format("{0:###,##0}", ds.baoHiemXH))
                    .Replace("{$thueTNCNThang}", String.Format("{0:###,##0}", ds.thueTNCN))
                    .Replace("{$ngayCongChuan}", String.Format("{0:###,##0}", ds.ngayCongChuan))
                    .Replace("{$ngayCongTinhLuong}", String.Format("{0:###,##0}", ds.ngayTinhLuong))
                     .Replace("{$chiTietThucNhanKyNay}", String.Format("{0:###,##0}", ds.conLaiPhaiChuyen))
                    .Replace("{$luongNgayCongERP}", String.Format("{0:###,##0}", ds.luongThang))
                    .Replace("{$luongChuyenDot1}", String.Format("{0:###,##0}", ds.luongDaChuyenDot1))
                    .Replace("{$truyThuKyTruoc}", String.Format("{0:###,##0}", ds.truyThu))
                    .Replace("{$truyLanhKyTruoc}", String.Format("{0:###,##0}", ds.truyLanh));


            }
            ViewBag.NoiDung = noiDung;
            // return PartialView("_ViewChiTietLuong");
            return PartialView("_ViewChiTietLuongTemplateNVVIP");
        }

    }
}
