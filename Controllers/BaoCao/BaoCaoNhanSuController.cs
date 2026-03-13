using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper;
using BatDongSan.Helper.Common;
using BatDongSan.Helper.Utils;
using BatDongSan.Models.ChamCong;
using BatDongSan.Models.DanhMuc;
using BatDongSan.Models.NhanSu;
using BatDongSan.Utils;
using BatDongSan.Utils.Paging;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.UserModel.Contrib;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using Worldsoft.Mvc.Web.Util;
using System.Net.Mail;
using BatDongSan.Models.PhieuDeNghi;
using System.Configuration;
using NPOI.SS.Util;

namespace BatDongSan.Controllers.BaoCao
{
    public class BaoCaoNhanSuController : ApplicationController
    {
        private LinqNhanSuDataContext context = new LinqNhanSuDataContext();
        LinqPhieuDeNghiDataContext lqPhieuDN = new LinqPhieuDeNghiDataContext();
        private IList<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan> phongBans;

        private StringBuilder buildTree;
        private readonly string MCV = "BaoCaoNhanSu";
        private bool? permission;
        //
        // GET: /BaoCaoNhanSu/

        public ActionResult Index()
        {
            return View();
        }
        public ActionResult BCSinhNhatNV(int? page, int? pageSize, string searchString, int? thang, int? day)
        {
            #region Role user
            permission = GetPermission("BCSinhNhatNV", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            BuilThang(DateTime.Now.Month);
            thang = thang.HasValue ? thang : DateTime.Now.Month;

            return View("");
        }

        public ActionResult BCSinhNhatNVViewIndex(int? thang, int? day, int? pageSize, string searchString, int _page = 0)
        {
            try
            {
                #region Role user
                permission = GetPermission("BCSinhNhatNV", BangPhanQuyen.QuyenXem);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion


                context = new LinqNhanSuDataContext();
                ViewBag.isGet = "True";

                string maNhanVien = GetUser().manv;
                int page = _page == 0 ? 1 : _page;
                int pIndex = page;

                int total = context.sp_BC_NS_SinhNhatNhanVien(thang, day, searchString).Count();

                PagingLoaderController("/BaoCaoNhanSu/BCSinhNhatNV/", total, page, "?searchString=" + searchString + "&thang=" + thang + "&day=" + day);
                ViewData["lsDanhSach"] = context.sp_BC_NS_SinhNhatNhanVien(thang, day, searchString).Skip(start).Take(offset).ToList();

                return PartialView("ViewIndexSN");
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }

        }
        private void BuilThang(int thang)
        {
            Dictionary<int, int> dict = new Dictionary<int, int>();
            for (var i = 1; i <= 12; i++)
            {
                dict.Add(i, i);
            }
            ViewBag.Thangs = new SelectList(dict, "Key", "Value", thang);
        }
        public ActionResult BCDiTreVeSom(int? page, int? pageSize, string searchString, string tuNgay, string denNgay)
        {
            #region Role user
            permission = GetPermission("BCDiTreVeSom", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion


            return View("");
        }
        public ActionResult BCDiTreVeSomViewIndex(int? pageSize, string searchString, string tuNgay, string denNgay, int _page = 0)
        {
            try
            {
                #region Role user
                permission = GetPermission("BCDiTreVeSom", BangPhanQuyen.QuyenXem);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion

                DateTime? fromDate = null;
                DateTime? toDate = null;
                if (!String.IsNullOrEmpty(tuNgay))
                    fromDate = DateTime.ParseExact(tuNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                if (!String.IsNullOrEmpty(denNgay))
                    toDate = DateTime.ParseExact(denNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                context = new LinqNhanSuDataContext();
                ViewBag.isGet = "True";

                string maNhanVien = GetUser().manv;
                int page = _page == 0 ? 1 : _page;
                int pIndex = page;

                int total = context.sp_BC_NS_NhanVienDiTreVeSom(fromDate, toDate, searchString).Count();

                PagingLoaderController("/BaoCaoNhanSu/BCDiTreVeSom/", total, page, "?searchString=" + searchString + "&tuNgay=" + tuNgay + "&denNgay=" + denNgay);
                ViewData["lsDanhSach"] = context.sp_BC_NS_NhanVienDiTreVeSom(fromDate, toDate, searchString).Skip(start).Take(offset).ToList();

                return PartialView("ViewIndexDTVS");
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }

        }
        public ActionResult BCNhanVienNghiPhep(int? page, int? pageSize, string searchString, string tuNgay, string denNgay)
        {
            #region Role user
            permission = GetPermission("BCNhanVienNghiPhep", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            return View("BCNhanVienNghiPhep");
        }
        public ActionResult BCNhanVienNghiPhepViewIndex(int? pageSize, string searchString, string tuNgay, string denNgay, int _page = 0)
        {
            try
            {
                #region Role user
                permission = GetPermission("BCNhanVienNghiPhep", BangPhanQuyen.QuyenXem);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion

                DateTime? fromDate = null;
                DateTime? toDate = null;
                if (!String.IsNullOrEmpty(tuNgay))
                    fromDate = DateTime.ParseExact(tuNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                if (!String.IsNullOrEmpty(denNgay))
                    toDate = DateTime.ParseExact(denNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                context = new LinqNhanSuDataContext();
                ViewBag.isGet = "True";

                string maNhanVien = GetUser().manv;
                int page = _page == 0 ? 1 : _page;
                int pIndex = page;

                int total = context.sp_BC_NS_NhanVienNghiPhep(fromDate, toDate, searchString).Count();

                PagingLoaderController("/BaoCaoNhanSu/BCNhanVienNghiPhep/", total, page, "?searchString=" + searchString + "&tuNgay=" + tuNgay + "&denNgay=" + denNgay);
                ViewData["lsDanhSach"] = context.sp_BC_NS_NhanVienNghiPhep(fromDate, toDate, searchString).Skip(start).Take(offset).ToList();

                return PartialView("ViewIndexNP");
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }

        }
        public ActionResult BCTQNghiPhep()
        {
            #region Role user
            permission = GetPermission("BCTQNghiPhep", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            buildTree = new StringBuilder();
            phongBans = context.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan>().ToList();
            buildTree = TreePhongBanStyle.BuildTreeDepartment(phongBans);
            ViewBag.PhongBans = buildTree.ToString();
            nam(DateTime.Now.Year);
            return View("BCTQNghiPhep");
        }
        public ActionResult BCTQNghiPhepViewIndex(int? pageSize, string searchString, string maPhongBan, int nam, int _page = 0)
        {
            try
            {
                #region Role user
                permission = GetPermission("BCTQNghiPhep", BangPhanQuyen.QuyenXem);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion
                context = new LinqNhanSuDataContext();
                ViewBag.isGet = "True";

                string maNhanVien = GetUser().manv;
                int page = _page == 0 ? 1 : _page;
                int pIndex = page;

                int total = context.sp_BC_NghiPhep_Index(maPhongBan,searchString,nam).Count();

                PagingLoaderController("/BaoCaoNhanSu/BCTQNghiPhep/", total, page, "?searchString=" + searchString + "&maPhongBan=" + maPhongBan + "&nam=" + nam);
                ViewData["lsDanhSach"] = context.sp_BC_NghiPhep_Index(maPhongBan, searchString, nam).Skip(start).Take(offset).ToList();

                return PartialView("ViewIndexNPTQ");
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }

        }
        #region Xuat File XuatFileBangTongHopCongThang
        public void XuatFileBangTongHopCongThang(string searchString, string maPhongBan, int nam)
        {
            try
            {
                var filename = "";
                var virtualPath = HttpRuntime.AppDomainAppVirtualPath;

                var fileStream = new FileStream(System.Web.HttpContext.Current.Server.MapPath(virtualPath + @"\Content\Report\ReportTemplateTHCT.xls"), FileMode.Open, FileAccess.Read);

                var workbook = new HSSFWorkbook(fileStream, true);

                filename += "BaoCaoTongQuanNghiPhep_" +nam + ".xls";

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



                var sheet = workbook.CreateSheet("BaoCaoTongQuanNghiPhep");

                //Khai báo row đầu tiên
                int firstRowNumber = 3;

                string cellTenCty = "TỔNG CÔNG TY XDCTGT 6 - CÔNG TY CỔ PHẦN";
                var titleCellCty = HSSFCellUtil.CreateCell(sheet.CreateRow(0), 0, cellTenCty.ToUpper());
                titleCellCty.CellStyle = styleTitle;

                string cellTenCacBanDH = "CÁC BAN ĐIỀU HÀNH DỰ ÁN";
                var titleCellTenCacBanDH = HSSFCellUtil.CreateCell(sheet.CreateRow(1), 1, cellTenCacBanDH.ToUpper());
                titleCellTenCacBanDH.CellStyle = styleTitle;

                string cellTitleMain = "Báo cáo tổng quan nghỉ phép " + nam;
                var titleCellTitleMain = HSSFCellUtil.CreateCell(sheet.CreateRow(2), 5, cellTitleMain.ToUpper());
                titleCellTitleMain.CellStyle = styleTitle;

                firstRowNumber++;
                var list1 = new List<string>();
                list1.Add("STT");
                list1.Add("Họ tên");
                list1.Add("Mã nhân viên");
                list1.Add("Phòng ban");
                list1.Add("Số ngày phép năm hiện tại");
                list1.Add("Ngày nghỉ");
                list1.Add("");
                list1.Add("");
                list1.Add("");
                list1.Add("");
                list1.Add("");
                list1.Add("");
                list1.Add("Phép năm còn lại");
                var list2 = new List<string>();
                list2.Add("STT");
                list2.Add("Họ tên");
                list2.Add("Mã nhân viên");
                list2.Add("Phòng ban");
                list2.Add("Số ngày phép năm hiện tại");
                list2.Add("Phép năm");
                list2.Add("Phép cưới");
                list2.Add("Phép tang");
                list2.Add("Nghỉ bù");
                list2.Add("Khám thai sinh con");
                list2.Add("Nghỉ bệnh");
                list2.Add("Nghỉ không lương");
                list2.Add("Phép năm còn lại");

                var idRowStart = firstRowNumber; // bat dau o dong thu 4
                var headerRow = sheet.CreateRow(idRowStart);
                int rowend = idRowStart;
                ReportHelperExcel.CreateHeaderRow(headerRow, 0, styleheadedColumnTable, list1);
                idRowStart++;
                var headerRow1 = sheet.CreateRow(idRowStart);
                ReportHelperExcel.CreateHeaderRow(headerRow1, 0, styleheadedColumnTable, list2);
                

                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 12, 12));
                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow.RowNum, 5, 11));

                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 4, 4));
                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 3, 3));
                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 2, 2));
                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 1, 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 0, 0));

                sheet.SetColumnWidth(0, 5 * 210);
                sheet.SetColumnWidth(1, 25 * 210);
                sheet.SetColumnWidth(2, 15 * 210);
                sheet.SetColumnWidth(3, 30 * 210);
                sheet.SetColumnWidth(4, 15 * 210);
                sheet.SetColumnWidth(5, 15 * 210);
                sheet.SetColumnWidth(6, 15 * 210);
                sheet.SetColumnWidth(7, 15 * 210);
                sheet.SetColumnWidth(8, 15 * 210);
                sheet.SetColumnWidth(9, 15 * 210);
                sheet.SetColumnWidth(10, 15 * 210);
                sheet.SetColumnWidth(11, 15 * 210);

                sheet.SetColumnWidth(12, 15 * 250);
                var data = context.sp_BC_NghiPhep_Index(maPhongBan, searchString, nam).ToList();
                var stt = 0;
                int dem = 0;

                foreach (var item1 in data)
                {
                    dem = 0;

                    stt++;
                    idRowStart++;

                    rowC = sheet.CreateRow(idRowStart);
                    ReportHelperExcel.SetAlignment(rowC, dem++, stt.ToString(), hStyleConCenter);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.hoTen, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.maNhanVien, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.tenPhongBan, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, Convert.ToString(item1.soNgayPhepNamHienTai), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, Convert.ToString(item1.phepNam), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, Convert.ToString(item1.phepCuoi), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, Convert.ToString(item1.phepTang), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, Convert.ToString(item1.nghiBu), hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowC, dem++, Convert.ToString(item1.khamThaiSinhCon), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, Convert.ToString(item1.nghiBenh), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, Convert.ToString(item1.nghiKhongLuong), hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowC, dem++, Convert.ToString(item1.phepNamConLai), hStyleConRight);
                }


                idRowStart = idRowStart + 2;
                var date = DateTime.Now.Day;
                string cellFooterNgayLap = "Tp.Hồ Chí Minh, ngày " + date +" năm " + nam;
                var titleCellFooterNgayLap = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 8, cellFooterNgayLap);
                titleCellFooterNgayLap.CellStyle = styleTitleItalic;


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
        public ActionResult BCHopDongSapHetHan(int? page, int? pageSize, string searchString, int? soNgayHetHan)
        {
            #region Role user
            permission = GetPermission("BCHopDongSapHetHan", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion


            return View("");
        }
        public ActionResult BCHopDongSapHetHanViewIndex(int? soNgayHetHan, string searchString, int _page = 0)
        {
            try
            {
                #region Role user
                permission = GetPermission("BCHopDongSapHetHan", BangPhanQuyen.QuyenXem);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion


                context = new LinqNhanSuDataContext();
                ViewBag.isGet = "True";

                string maNhanVien = GetUser().manv;
                int page = _page == 0 ? 1 : _page;
                int pIndex = page;

                int total = context.sp_BC_NS_HopDongSapHetHan(soNgayHetHan, searchString).Count();

                PagingLoaderController("/BaoCaoNhanSu/BCHopDongSapHetHan/", total, page, "?searchString=" + searchString + "&soNgayHetHan=" + soNgayHetHan);
                ViewData["lsDanhSach"] = context.sp_BC_NS_HopDongSapHetHan(soNgayHetHan, searchString).Skip(start).Take(offset).ToList();

                return PartialView("ViewIndexHDHH");
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }

        }
        public ActionResult BCTangCaTrongNam(int? page, int? pageSize, string searchString)
        {
            #region Role user
            permission = GetPermission("BCTangCaTrongNam", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion



            return View("");
        }
        public ActionResult BCTangCaTrongNamViewIndex(string searchString, int _page = 0)
        {
            try
            {
                #region Role user
                permission = GetPermission("BCTangCaTrongNam", BangPhanQuyen.QuyenXem);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion


                context = new LinqNhanSuDataContext();
                ViewBag.isGet = "True";

                string maNhanVien = GetUser().manv;
                int page = _page == 0 ? 1 : _page;
                int pIndex = page;

                int total = lqPhieuDN.sp_NS_TongSoGioTangCa(searchString).Count();

                PagingLoaderController("/BaoCaoNhanSu/BCTangCaTrongNam/", total, page, "?searchString=" + searchString);
                ViewData["lsDanhSach"] = lqPhieuDN.sp_NS_TongSoGioTangCa(searchString).Skip(start).Take(offset).ToList();

                return PartialView("ViewIndexTangCaTrongNam");
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }

        }
        public ActionResult BieuDoLuong()
        {
            thang(DateTime.Now.Month);
            nam(DateTime.Now.Year);
            return View();
        }
        public ActionResult GetListBieuDoLuong(int? thangFrom, int? thangTo, int? namFrom, int? namTo)
        {

            #region Role user
            permission = GetPermission("BieuDoLuong", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            try
            {

                var list = context.sp_BC_NS_BieuDoLuongThang(thangFrom, namFrom, thangTo, namTo).ToList();
                var result = new { kq = list };
                return Json(result, JsonRequestBehavior.AllowGet);





            }
            catch
            {
                var result = new { kq = false };
                return Json(result, JsonRequestBehavior.AllowGet);
            }

        }
        private void thang(int value)
        {
            Dictionary<int, string> dics = new Dictionary<int, string>();
            for (int i = 0; i < 13; i++)
            {
                dics[i] = i.ToString();
            }
            ViewData["lstThangFrom"] = new SelectList(dics, "Key", "Value", value);
            ViewData["lstThangTo"] = new SelectList(dics, "Key", "Value", value);

            
        }
        private void nam(int value)
        {
            Dictionary<int, string> dics = new Dictionary<int, string>();
            for (int i = 2015; i < 2031; i++)
            {
                dics[i] = i.ToString();
            }
            ViewData["lstNamFrom"] = new SelectList(dics, "Key", "Value", value);
            ViewData["lstNamTo"] = new SelectList(dics, "Key", "Value", value);
            ViewData["lstNam"] = new SelectList(dics, "Key", "Value", value);

        }
        // check mail trong thang da send
        public ActionResult CheckUpdateMailSN(int thang, string type)
        {


            try
            {
                #region Role user
                permission = GetPermission("BCSinhNhatNV", BangPhanQuyen.QuyenXem);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion
                //Check 
                var result = new { kq = false };

              
                if (type == "month")
                {
                     var checkEx = context.tbl_TrangThaiSinhNhats.Where(t => t.nam == DateTime.Now.Year && t.thang == thang && t.day == null).FirstOrDefault();
                     if (checkEx != null)
                     {
                         return Json(result, JsonRequestBehavior.AllowGet);
                     }
                }
                else {
                     var checkEx = context.tbl_TrangThaiSinhNhats.Where(t => t.nam == DateTime.Now.Year && t.thang == thang && t.day == DateTime.Now.Day).FirstOrDefault();
                     if (checkEx != null)
                     {
                         return Json(result, JsonRequestBehavior.AllowGet);
                     }
                }
                    
                //End check
                // Insert Row 
                tbl_TrangThaiSinhNhat tblDuyetBL = new tbl_TrangThaiSinhNhat();
                tblDuyetBL.nam = DateTime.Now.Year;
                tblDuyetBL.thang = thang;
                if (type == "month")
                {
                    tblDuyetBL.day = null;
                }
                else {
                    tblDuyetBL.day = DateTime.Now.Day;
                }
                tblDuyetBL.ngayDuyet = DateTime.Now;
                tblDuyetBL.nguoiDuyet = GetUser().manv;
                context.tbl_TrangThaiSinhNhats.InsertOnSubmit(tblDuyetBL);
                context.SubmitChanges();
                // End Insert Row
                // Check Exist

                
                result = new { kq = true };
                
                return Json(result, JsonRequestBehavior.AllowGet);

            }
            catch
            {
                var result = new { kq = false };
                return Json(result, JsonRequestBehavior.AllowGet);
            }

        }
        // end check mail
        public ActionResult SendMailSNThang(int? thang)
        {
            #region Role user
            permission = GetPermission("BCSinhNhatNV", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            string qSearch = "";

            var listSendMails = context.sp_BC_NS_SinhNhatNhanVien(thang,null, qSearch).ToList();
            foreach (var item in listSendMails)
            {
                // Code send mail
                MailHelper mailInit = new MailHelper(); // lay cac tham so trong webconfig
                System.Text.StringBuilder content = new System.Text.StringBuilder();


                //Content 
                content.Append("<div style=\"width: 493px; margin: 16px auto; line-height: 1.7; font-size: 16px; box-shadow: -3px -3px  #4b6580; padding: 10px; color: green; border: 3px dotted #0c7932; border-radius: 7px;\">");
                content.Append("<p>Xin chào: <span style=\"font-size: 25px; color: #fb4d0b;\">" + item.hoTen + "</span>");
                content.Append("</p>");
                content.Append("<p style=\"font-style: italic; font-size: 30px; color: #3eaf3e; text-align: center;\">Chúc mừng <span style=\"font-size: 30px; text-transform: uppercase; color: #fb4d0b;\">Sinh nhật </span><span style=\"font-size: 37px; color: #d21839; display:block; clear:both; float:none;\">" + thang + "/" + DateTime.Now.Year + "</span>");
                content.Append("</p>");
                content.Append("<p style=\"font-style: italic;\">Thanks and Regards!</p>");
                content.Append("<p style=\"font-style: italic;\">Email từ hệ thống nhân sự</p>");
                content.Append("</div>");
                //End content
                //Send only email is @thuanviet.com.vn
                string[] array01 = item.email.ToLower().Split('@');
                string string2 = ConfigurationManager.AppSettings["OnlySend"]; //get domain from config files
                string[] array1 = string2.Split(',');
                // bool EmailofThuanViet;
                //EmailofThuanViet = array1.Contains(array01[1]);
                // if (emailNV == "" || emailNV == null || EmailofThuanViet == false)
                // {
                //    return false;
                // }
                MailAddress toMail = new MailAddress(item.email, item.hoTen); // goi den mail
                mailInit.ToMail = toMail;
                mailInit.Body = content.ToString();
                mailInit.SendMail();
                // End code send mail
            }
            var result = new { kq = true };
            return Json(result, JsonRequestBehavior.AllowGet);
        }

        public ActionResult SendMailSNNgay(int? thang)
        {
            try
            {
                #region Role user
                permission = GetPermission("BCSinhNhatNV", BangPhanQuyen.QuyenXem);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion
                string qSearch = "";

                var listSendMails = context.sp_BC_NS_SinhNhatNhanVien(thang, DateTime.Now.Day, qSearch).ToList();
                foreach (var item in listSendMails)
                {
                    // Code send mail
                    MailHelper mailInit = new MailHelper(); // lay cac tham so trong webconfig
                    System.Text.StringBuilder content = new System.Text.StringBuilder();


                    //Content 
                    content.Append("<div style=\"width: 493px; margin: 16px auto; line-height: 1.7; font-size: 16px; box-shadow: -3px -3px  #4b6580; padding: 10px; color: green; border: 3px dotted #0c7932; border-radius: 7px;\">");
                    content.Append("<p>Xin chào: <span style=\"font-size: 25px; color: #fb4d0b;\">" + item.hoTen + "</span>");
                    content.Append("</p>");
                    content.Append("<p style=\"font-style: italic; font-size: 30px; color: #3eaf3e; text-align: center;\">Chúc mừng <span style=\"font-size: 30px; text-transform: uppercase; color: #fb4d0b;\">Sinh nhật </span><span style=\"font-size: 37px; color: #d21839; display:block; clear:both; float:none;\">" + DateTime.Now.Day + "/" + thang + "</span>");
                    content.Append("</p>");
                    content.Append("<p style=\"font-style: italic;\">Thanks and Regards!</p>");
                    content.Append("<p style=\"font-style: italic;\">Email từ hệ thống nhân sự</p>");
                    content.Append("</div>");
                    //End content
                    //Send only email is @thuanviet.com.vn
                    string[] array01 = item.email.ToLower().Split('@');
                    string string2 = ConfigurationManager.AppSettings["OnlySend"]; //get domain from config files
                    string[] array1 = string2.Split(',');
                    // bool EmailofThuanViet;
                    //EmailofThuanViet = array1.Contains(array01[1]);
                    // if (emailNV == "" || emailNV == null || EmailofThuanViet == false)
                    // {
                    //    return false;
                    // }
                    MailAddress toMail = new MailAddress(item.email, item.hoTen); // goi den mail
                    mailInit.ToMail = toMail;
                    mailInit.Body = content.ToString();
                    mailInit.SendMail();
                    // End code send mail
                }

                var result = new { kq = true, day = DateTime.Now.Day };
                return Json(result, JsonRequestBehavior.AllowGet);

            }
            catch{
                var result = new { kq = false, day = DateTime.Now.Day };
                return Json(result, JsonRequestBehavior.AllowGet);

            }
        }

    }
}
