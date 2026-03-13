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
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using NPOI.HSSF.UserModel.Contrib;
using NPOI.SS.Util;
using Worldsoft.Mvc.Web.Util;
using BatDongSan.Models.DanhMuc;
using BatDongSan.Models.DBChamCong;
using BatDongSan.Models.ERP;
using BatDongSan.Models.HeThong;
using BatDongSan.Models.PhieuDeNghi;
namespace BatDongSan.Controllers.TinhCong
{
    public class BangChamCongTongHopController : ApplicationController
    {
        private LinqDanhMucDataContext context = new LinqDanhMucDataContext();
        private LinqNhanSuDataContext nhanSuContext = new LinqNhanSuDataContext();
        private IList<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan> phongBans;
        private StringBuilder buildTree;
        private readonly string MCV = "ChamCongAM";
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
            thang(DateTime.Now.Month);
            nam(DateTime.Now.Year);
            return View("");

        }
        public ActionResult XemBangChamCongChiTiet()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            //  DayDuLieuChamCongVe();

            buildTree = new StringBuilder();
            phongBans = nhanSuContext.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan>().ToList();
            buildTree = TreePhongBanAjax.BuildTreeDepartment(phongBans);
            ViewBag.PhongBans = buildTree.ToString();
            thang(DateTime.Now.Month);
            nam(DateTime.Now.Year);

            TuNgay(1);
            DenNgay(10);
            return View("");
        }
        public ActionResult LoadBangChamCongChiTiet(string qSearch, int thang, int nam, int _page = 0)
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
            int total = nhanSuContext.sp_NS_BangChamCongChiTiet(thang, nam, qSearch, "").Count();
            PagingLoaderController("/BangChamCongTongHop/LoadBangChamCongChiTiet/", total, page, "?qsearch=" + qSearch);
            ViewData["lsDanhSach"] = nhanSuContext.sp_NS_BangChamCongChiTiet(thang, nam, qSearch, "").Skip(start).Take(offset).ToList();

            ViewData["qSearch"] = qSearch;
            return PartialView("_LoadBangChamCongChiTiet");
        }

        #region Xuat File Bang cham cong tong hop
        public void XuatFileBangChamCongTH(int thang, int nam)
        {

            try
            {

                var soNgay = DateTime.DaysInMonth(nam, thang);
                var filename = "";
                var virtualPath = HttpRuntime.AppDomainAppVirtualPath;

                var fileStream = new FileStream(System.Web.HttpContext.Current.Server.MapPath(virtualPath + @"\Content\Report\ReportTemplatecc.xls"), FileMode.Open, FileAccess.Read);

                var workbook = new HSSFWorkbook(fileStream, true);
                filename += "BangChamCongTongHop_" + nam + "_" + thang + ".xls";

                #region format style excel cell
                /*style title start*/
                //tạo font cho các title
                //font tiêu đề 
                HSSFFont hFontTieuDe = (HSSFFont)workbook.CreateFont();
                hFontTieuDe.FontHeightInPoints = 18;
                hFontTieuDe.Boldweight = 100 * 10;
                hFontTieuDe.FontName = "Times New Roman";
                hFontTieuDe.Color = HSSFColor.BLACK.index;
                //font tiêu đề Yellow
                HSSFFont hFontTieuDeYellow = (HSSFFont)workbook.CreateFont();
                hFontTieuDeYellow.FontHeightInPoints = 18;
                hFontTieuDeYellow.Boldweight = 100 * 10;
                hFontTieuDeYellow.FontName = "Times New Roman";
                hFontTieuDeYellow.Color = HSSFColor.BLACK.index;

                //font chữ bình thường
                HSSFFont hFontNommal = (HSSFFont)workbook.CreateFont();
                hFontNommal.Color = HSSFColor.BLACK.index;
                hFontNommal.FontName = "Times New Roman";

                //font Cty
                HSSFFont hFontNommalCTY = (HSSFFont)workbook.CreateFont();
                hFontNommalCTY.FontHeightInPoints = 18;
                hFontNommalCTY.Boldweight = 100 * 10;
                hFontNommalCTY.FontName = "Times New Roman";
                hFontNommalCTY.Color = HSSFColor.BLACK.index;

                HSSFFont hFontNommalTieuDe = (HSSFFont)workbook.CreateFont();
                hFontNommalTieuDe.FontHeightInPoints = 18;
                hFontNommalTieuDe.FontName = "Times New Roman";
                hFontNommalTieuDe.Color = HSSFColor.BLACK.index;

                //font tiêu đề 
                HSSFFont hFontTongGiaTriHT = (HSSFFont)workbook.CreateFont();
                hFontTongGiaTriHT.FontHeightInPoints = 11;
                hFontTongGiaTriHT.Boldweight = (short)FontBoldWeight.BOLD;
                hFontTongGiaTriHT.FontName = "Times New Roman";
                hFontTongGiaTriHT.Color = HSSFColor.BLACK.index;

                //font thông tin bảng tính
                HSSFFont hFontTT = (HSSFFont)workbook.CreateFont();
                hFontTT.IsItalic = true;
                hFontTT.Boldweight = (short)FontBoldWeight.NORMAL;
                hFontTT.Color = HSSFColor.BLACK.index;
                hFontTT.FontName = "Times New Roman";
                hFontTieuDe.FontHeightInPoints = 11;
                HSSFFont hFontTT13 = (HSSFFont)workbook.CreateFont();
                hFontTT13.IsItalic = false;
                hFontTT13.Boldweight = (short)FontBoldWeight.BOLD;
                hFontTT13.Color = HSSFColor.BLACK.index;
                hFontTT13.FontName = "Times New Roman";
                hFontTT13.FontHeightInPoints = 10;
                HSSFFont hFontTT26 = (HSSFFont)workbook.CreateFont();
                hFontTT26.IsItalic = false;
                hFontTT26.Boldweight = (short)FontBoldWeight.NORMAL;
                hFontTT26.Color = HSSFColor.BLACK.index;
                hFontTT26.FontName = "Times New Roman";
                hFontTT26.FontHeightInPoints = 12;

                //font chứ hoa đậm
                HSSFFont hFontNommalUpper = (HSSFFont)workbook.CreateFont();
                hFontNommalUpper.Boldweight = (short)FontBoldWeight.BOLD;
                hFontNommalUpper.Color = HSSFColor.BLACK.index;
                hFontNommalUpper.FontName = "Times New Roman";



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
                var styleTitleYellow = workbook.CreateCellStyle();
                styleTitleYellow.SetFont(hFontTieuDe);
                styleTitleYellow.Alignment = HorizontalAlignment.JUSTIFY;
                styleTitleYellow.FillBackgroundColor = HSSFColor.YELLOW.index;
                styleTitleYellow.FillForegroundColor = HSSFColor.YELLOW.index;
                styleTitleYellow.FillPattern = FillPatternType.SOLID_FOREGROUND;
                var styleTitle2 = workbook.CreateCellStyle();
                styleTitle2.SetFont(hFontTT);
                styleTitle2.Alignment = HorizontalAlignment.LEFT;
                var styleTitle13 = workbook.CreateCellStyle();
                styleTitle13.SetFont(hFontTT13);
                styleTitle13.Alignment = HorizontalAlignment.LEFT;
                var styleTitle26 = workbook.CreateCellStyle();
                styleTitle26.SetFont(hFontTT26);
                styleTitle26.Alignment = HorizontalAlignment.LEFT;


                var styleTitleCTY = workbook.CreateCellStyle();
                styleTitleCTY.SetFont(hFontNommalCTY);
                styleTitleCTY.Alignment = HorizontalAlignment.LEFT;
                var styleTitleTieuDe = workbook.CreateCellStyle();
                styleTitleTieuDe.SetFont(hFontNommalTieuDe);
                styleTitleTieuDe.Alignment = HorizontalAlignment.LEFT;

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



                var hStyleConLeftBack = (HSSFCellStyle)workbook.CreateCellStyle();
                hStyleConLeftBack.SetFont(hFontNommal);
                hStyleConLeftBack.VerticalAlignment = VerticalAlignment.TOP;
                hStyleConLeftBack.Alignment = HorizontalAlignment.LEFT;
                hStyleConLeftBack.WrapText = true;
                hStyleConLeftBack.BorderBottom = CellBorderType.THIN;
                hStyleConLeftBack.BorderLeft = CellBorderType.THIN;
                hStyleConLeftBack.BorderRight = CellBorderType.THIN;
                hStyleConLeftBack.BorderTop = CellBorderType.THIN;
                hStyleConLeftBack.FillBackgroundColor = HSSFColor.GREY_25_PERCENT.index;
                hStyleConLeftBack.FillForegroundColor = HSSFColor.GREY_25_PERCENT.index;
                hStyleConLeftBack.FillPattern = FillPatternType.SOLID_FOREGROUND;

                var hStyleConCenterBack = (HSSFCellStyle)workbook.CreateCellStyle();
                hStyleConCenterBack.SetFont(hFontNommal);
                hStyleConCenterBack.VerticalAlignment = VerticalAlignment.TOP;
                hStyleConCenterBack.Alignment = HorizontalAlignment.CENTER;
                hStyleConCenterBack.WrapText = true;
                hStyleConCenterBack.BorderBottom = CellBorderType.THIN;
                hStyleConCenterBack.BorderLeft = CellBorderType.THIN;
                hStyleConCenterBack.BorderRight = CellBorderType.THIN;
                hStyleConCenterBack.BorderTop = CellBorderType.THIN;
                hStyleConCenterBack.FillBackgroundColor = HSSFColor.GREY_25_PERCENT.index;
                hStyleConCenterBack.FillForegroundColor = HSSFColor.GREY_25_PERCENT.index;
                hStyleConCenterBack.FillPattern = FillPatternType.SOLID_FOREGROUND;

                var hStyleConCenterYellowBack = (HSSFCellStyle)workbook.CreateCellStyle();
                hStyleConCenterYellowBack.SetFont(hFontNommal);
                hStyleConCenterYellowBack.VerticalAlignment = VerticalAlignment.TOP;
                hStyleConCenterYellowBack.Alignment = HorizontalAlignment.CENTER;
                hStyleConCenterYellowBack.WrapText = true;
                hStyleConCenterYellowBack.BorderBottom = CellBorderType.THIN;
                hStyleConCenterYellowBack.BorderLeft = CellBorderType.THIN;
                hStyleConCenterYellowBack.BorderRight = CellBorderType.THIN;
                hStyleConCenterYellowBack.BorderTop = CellBorderType.THIN;
                hStyleConCenterYellowBack.FillBackgroundColor = HSSFColor.YELLOW.index;
                hStyleConCenterYellowBack.FillForegroundColor = HSSFColor.YELLOW.index;
                hStyleConCenterYellowBack.FillPattern = FillPatternType.SOLID_FOREGROUND;


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


                string cellTitleMain = "BẢNG CHẤM CÔNG THÁNG " + thang + "/" + nam;
                var sheet = workbook.CreateSheet("thang_" + thang + "_nam_" + nam);
                workbook.ActiveSheetIndex = 1;


                //Khai báo row đầu tiên
                int firstRowNumber = 0;

                string cellTenCty = "CƠ QUAN TỔNG CÔNG TY XDCTGT 6";
                //var titleCellCty = HSSFCellUtil.CreateCell(sheet.CreateRow(1), 0, cellTenCty.ToUpper());
                //titleCellCty.CellStyle = styleTitle;




                string nguoiChamCong = "Phụ trách đơn vị";
                string phuTrachDonVi = "Người theo dõi chấm công";
                string ngayThangNam = "Ngày   tháng   năm " + nam;
                string nguoiDuyetCong = "Người duyệt công";
                //var titleCellTitleMain = HSSFCellUtil.CreateCell(sheet.CreateRow(2), 5, cellTitleMain.ToUpper());
                //titleCellTitleMain.CellStyle = styleTitle;
                var ListTH = nhanSuContext.sp_NS_BangChamCongChiTiet_PhongBan(thang, nam, null, null).ToList();
                firstRowNumber = 0;
                var idRowStart = firstRowNumber;

                var DMNghiLe = context.tbl_DM_NghiLes.Select(d => d.ngayNghiLe).ToList();

                var countNV = 0;
                if (ListTH.Count > 0)
                {
                    // Khoi tao nhan vien/
                    var MaNVBanDau = ListTH[0].maNhanVien;
                    var thuTuParentNhanVien = 0;
                    var FlashNhanVien = 0;
                    // end khoi tao nhan vien
                    int j = 0;
                    var HSMTBanDau = ListTH[0].maPhongBan;
                    var thuTuParent = 0;
                    var Flash = 0;
                    var FlashFooter = 0;
                    var maPb = 0;
                    var flasRow = 0;
                    var demSTT = 0;
                    var flasRowFooter = 0;
                    var thuTuParentFooter = 0;
                    var demNVPhongBan = 0;
                    var grpNhanVien = 0;
                    foreach (var item in ListTH)
                    {

                        if (HSMTBanDau == item.maPhongBan)
                        {
                            Flash = 0;
                            FlashFooter = 0;
                        }
                        else
                        {
                            HSMTBanDau = item.maPhongBan;
                            Flash = 1;
                            FlashFooter = 1;
                            j = 1;
                        }
                        if (Flash == 1 || thuTuParent == 0)
                        {
                            grpNhanVien = ListTH.Where(d => d.maPhongBan == item.maPhongBan).ToList().Count();
                            demNVPhongBan = 0;
                            if (flasRow == 1)
                            {
                                firstRowNumber = firstRowNumber + 8;
                            }
                            flasRow = 1;
                            thuTuParent = 1;
                            countNV = 0;
                            Flash = 0;
                            idRowStart = firstRowNumber + 2;
                            var headerRowCty = sheet.CreateRow(idRowStart);
                            idRowStart++;
                            var headerRowTieuDe = sheet.CreateRow(idRowStart);
                            idRowStart++;
                            var headerRow0 = sheet.CreateRow(idRowStart);
                            int rowend = idRowStart;
                            maPb = maPb + 1;
                            string cellTitleMainCTY = Convert.ToString(cellTenCty);
                            var titleCellTitleMainCTY = HSSFCellUtil.CreateCell(headerRowCty, 1, cellTitleMainCTY.ToUpper());
                            titleCellTitleMainCTY.CellStyle = styleTitle;
                            string cellTitleMainTieuDe = Convert.ToString(cellTitleMain);
                            var titleCellTitleMainTieuDe = HSSFCellUtil.CreateCell(headerRowTieuDe, 5, cellTitleMainTieuDe.ToUpper());
                            titleCellTitleMainTieuDe.CellStyle = styleTitle;
                            var tenPB = maPb + ". " + item.tenPhongBan;

                            string cellTitleMain2 = Convert.ToString(tenPB);
                            var titleCellTitleMain2 = HSSFCellUtil.CreateCell(headerRow0, 0, cellTitleMain2.ToUpper());
                            titleCellTitleMain2.CellStyle = styleTitleYellow;
                            sheet.AddMergedRegion(new CellRangeAddress(headerRow0.RowNum, headerRow0.RowNum, 0, 10));
                            sheet.SetColumnWidth(headerRow0.RowNum, 15 * 510);
                            idRowStart++;
                            var headerRow21 = sheet.CreateRow(idRowStart + 1);
                            var list1 = new List<string>();
                            list1.Add("STT");
                            list1.Add("Họ và tên");
                            list1.Add("Ngày công trong tháng");
                            for (var i = 2; i <= soNgay; i++)
                            {
                                list1.Add("");

                            }
                            var ngayGop = soNgay;

                            list1.Add("Công tác");
                            list1.Add("Chấm công");
                            list1.Add("Nghỉ Lễ");
                            list1.Add("Nghỉ Phép");
                            list1.Add("Nghỉ Ốm");
                            list1.Add("Nghỉ Không lương");
                            list1.Add("Nghỉ Việc riêng");
                            list1.Add("Tổng công");
                            list1.Add("Ghi chú");


                            var list2 = new List<string>();
                            list2.Add("STT");
                            list2.Add("Họ và tên");
                            //list2.Add("Ngày công trong tháng");
                            for (var i = 1; i <= soNgay; i++)
                            {
                                list2.Add(Convert.ToString(i));




                            }


                            list2.Add("Công tác");
                            list2.Add("Chấm công");
                            list2.Add("Nghỉ Lễ");
                            list2.Add("Nghỉ Phép");
                            list2.Add("Nghỉ Ốm");
                            list2.Add("Nghỉ Không lương");
                            list2.Add("Nghỉ Việc riêng");
                            list2.Add("Tổng công");
                            list2.Add("Ghi chú");


                            var headerRow = sheet.CreateRow(idRowStart);
                            ReportHelperExcel.CreateHeaderRow(headerRow, 0, styleheadedColumnTable, list1);
                            //var titleTC2 = HSSFCellUtil.GetCell(headerRow, ngayGop + 2);
                            //titleTC2.CellStyle = styleTitleYellow;
                            //var titleTC3 = HSSFCellUtil.GetCell(headerRow, ngayGop + 3);
                            //titleTC3.CellStyle = styleTitleYellow;
                            //var titleTC4 = HSSFCellUtil.GetCell(headerRow, ngayGop + 4);
                            //titleTC4.CellStyle = styleTitleYellow;
                            idRowStart++;
                            var headerRow1 = sheet.CreateRow(idRowStart);
                            ReportHelperExcel.CreateHeaderRow(headerRow1, 0, styleheadedColumnTable, list2);

                            sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, soNgay + 10, soNgay + 10));
                            sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, soNgay + 9, soNgay + 9));
                            sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, soNgay + 8, soNgay + 8));
                            sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, soNgay + 7, soNgay + 7));
                            sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, soNgay + 6, soNgay + 6));
                            sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, soNgay + 5, soNgay + 5));
                            sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, soNgay + 4, soNgay + 4));
                            sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, soNgay + 3, soNgay + 3));
                            sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, soNgay + 2, soNgay + 2));

                            sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow.RowNum, 2, (soNgay + 1)));
                            sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 1, 1));
                            sheet.AddMergedRegion(new CellRangeAddress(headerRow.RowNum, headerRow1.RowNum, 0, 0));



                            sheet.SetColumnWidth(0, 5 * 210);
                            sheet.SetColumnWidth(1, 30 * 210);

                            for (var i = 2; i <= soNgay + 1; i++)
                            {
                                sheet.SetColumnWidth(i, 10 * 110);


                            }


                            sheet.SetColumnWidth(soNgay + 10, 15 * 210);
                            sheet.SetColumnWidth(soNgay + 9, 15 * 210);
                            sheet.SetColumnWidth(soNgay + 8, 15 * 210);
                            sheet.SetColumnWidth(soNgay + 7, 15 * 210);
                            sheet.SetColumnWidth(soNgay + 6, 15 * 210);
                            sheet.SetColumnWidth(soNgay + 5, 15 * 210);

                            sheet.SetColumnWidth(soNgay + 2, 15 * 210);
                            sheet.SetColumnWidth(soNgay + 3, 15 * 210);
                            sheet.SetColumnWidth(soNgay + 4, 15 * 210);

                        }
                        if (Flash == 0)
                        {
                            demNVPhongBan++;
                            // Khoi tao nhan vien/
                            if (MaNVBanDau == item.maNhanVien)
                            {
                                FlashNhanVien = 0;
                            }
                            else
                            {
                                MaNVBanDau = item.maNhanVien;
                                FlashNhanVien = 1;

                            }
                            if (FlashNhanVien == 1 || thuTuParentNhanVien == 0)
                            {

                                countNV++;
                                firstRowNumber++;
                                thuTuParentNhanVien = 1;
                                FlashNhanVien = 0;
                                // end khoi tao nhan vien

                                var stt = 0;
                                int dem = 0;
                                double? sumLuongLe = 0;


                                dem = 0;

                                stt++;
                                idRowStart++;

                                rowC = sheet.CreateRow(idRowStart);
                                ReportHelperExcel.SetAlignment(rowC, dem++, countNV.ToString(), hStyleConCenter);
                                ReportHelperExcel.SetAlignment(rowC, dem++, item.hoTen, hStyleConLeft);
                                // For thu trong tuan cho tung nhan vien
                                var listNgayCongNV = ListTH.Where(d => d.maNhanVien == item.maNhanVien).ToList();
                                decimal tongNgayLeNhanVien = 0;
                                //decimal tongNgayPhepNhanVien = 0;
                                //decimal tongNgayNghiKhongLuong = 0;
                                decimal tongGioQuet = 0;
                                decimal tongNgayNghiPhepNamTheoQuiDinh = 0;
                                decimal tongNgayNghiPhepOmCoGiayBenhVien = 0;
                                decimal tongNgayNghiPhepRieng = 0;
                                decimal tongNgayNghiPhepKhongLuong = 0;
                                decimal tongNgayDiCongTac = 0;

                                for (var ngay = 1; ngay <= soNgay; ngay++)
                                {
                                    int ngayLe = 0;

                                    // Check ngay le
                                    foreach (var itemNgay in DMNghiLe)
                                    {
                                        if (ngay == itemNgay.Value.Date.Day && thang == itemNgay.Value.Date.Month && nam == itemNgay.Value.Date.Year)
                                        {
                                            ngayLe = 1;
                                        }
                                    }



                                    if (ngayLe == 1)
                                    {
                                        ++tongNgayLeNhanVien;
                                        // if la ngay le thi ko can check cong
                                        ReportHelperExcel.SetAlignment(rowC, dem++, "L", hStyleConCenterBack);
                                    }
                                    else
                                    {
                                        // Check ngay trong thang co di lam khong
                                        //select maLoaiNghiPhep from tbl_DM_LoaiNghiPhep where tinhCong = 1 and trangThai = 1
                                        //Có thuộc trong nhóm tính chấm công hay không?
                                        var loaiNghi = listNgayCongNV.Where(d => d.ngayTrongThang == ngay).Select(d => d.loai).FirstOrDefault();

                                        if (!string.IsNullOrEmpty(loaiNghi))
                                        {
                                            //string tenLoaiNghiPhep = context.tbl_DM_LoaiNghiPheps.Where(d => d.tenLoaiNghiPhep == loaiNghi && (d.trangThai ?? false)).Select(d=>d.).FirstOrDefault() ? ?string.Empty;
                                            //Phép năm theo qui định
                                            if (loaiNghi == "Phép năm theo qui định")
                                            {
                                                ++tongNgayNghiPhepNamTheoQuiDinh;
                                                ReportHelperExcel.SetAlignment(rowC, dem++, "P", hStyleConCenter);
                                            }
                                            else if (loaiNghi == "Nghỉ bệnh có giấy bệnh viện")//Nghỉ bệnh có giấy bệnh viện
                                            {
                                                ++tongNgayNghiPhepOmCoGiayBenhVien;
                                                ReportHelperExcel.SetAlignment(rowC, dem++, "O", hStyleConCenter);
                                            }
                                            else if (loaiNghi == "Nghỉ không lương")
                                            {
                                                ++tongNgayNghiPhepKhongLuong;
                                                ReportHelperExcel.SetAlignment(rowC, dem++, "K", hStyleConCenter);
                                            }
                                            else if (loaiNghi == "Chấm công")
                                            {
                                                tongGioQuet += Convert.ToDecimal(item.soGioCong ?? 0);
                                                ReportHelperExcel.SetAlignment(rowC, dem++, "X", hStyleConCenter);
                                            }
                                            else if (loaiNghi == "Công tác")
                                            {
                                                ++tongNgayDiCongTac;
                                                ReportHelperExcel.SetAlignment(rowC, dem++, "C", hStyleConCenter);
                                            }
                                            else
                                            {
                                                ++tongNgayNghiPhepRieng;
                                                ReportHelperExcel.SetAlignment(rowC, dem++, "R", hStyleConCenter);
                                            }
                                            ////Có tính công thì phun ra data
                                            //var nghiPhepCoKhongCoLuong = context.tbl_DM_LoaiNghiPheps.Where(d => d.tenLoaiNghiPhep == loaiNghi && (d.trangThai ?? false)).Select(d=>d.tinhCong).FirstOrDefault();
                                            ////Xuất phát từ danh mục ngày nghĩ
                                            //if(nghiPhepCoKhongCoLuong != null)
                                            //{
                                            //    if ((nghiPhepCoKhongCoLuong ?? false))
                                            //    {
                                            //        ++tongNgayPhepNhanVien;
                                            //        ReportHelperExcel.SetAlignment(rowC, dem++, "P", hStyleConCenter);
                                            //    }
                                            //    else
                                            //    {
                                            //        ++tongNgayNghiKhongLuong;
                                            //        ReportHelperExcel.SetAlignment(rowC, dem++, "R", hStyleConCenter);
                                            //    }
                                            //}
                                            //else
                                            //{
                                            //    tongGioQuet += Convert.ToDecimal(item.soGioCong ?? 0);
                                            //    ReportHelperExcel.SetAlignment(rowC, dem++, "X", hStyleConCenter);
                                            //}
                                        }
                                        else
                                        {
                                            //Trường hợp không cần chấm công
                                            string ngayCanCheck = ngay + "/" + thang + "/" + nam;
                                            DateTime ngayCK = DateTime.ParseExact(ngayCanCheck, "d/M/yyyy", CultureInfo.InvariantCulture);

                                            // Check có phải là thứ 7, cn khong
                                            if (ngayCK.DayOfWeek == DayOfWeek.Saturday || ngayCK.DayOfWeek == DayOfWeek.Sunday)
                                            {
                                                ReportHelperExcel.SetAlignment(rowC, dem++, "", hStyleConLeftBack);
                                            }
                                            else
                                            {
                                                ReportHelperExcel.SetAlignment(rowC, dem++, "0", hStyleConCenter);
                                            }
                                        }
                                    }
                                    // End for thu trong tuan cho tung nhan vien


                                }
                                //var getLoaiNghi = nhanSuContext.sp_NS_BangTongHopCongThang(thang, nam, "", "", 0, "").Where(d => d.maNhanVien == item.maNhanVien).FirstOrDefault();
                                //if (getLoaiNghi != null)
                                {
                                    //list2.Add("Tổng công");
                                    //list2.Add("Nghỉ Lễ");
                                    //list2.Add("Nghỉ Phép");
                                    //list2.Add("Nghỉ Ốm");
                                    //list2.Add("Nghỉ Không lương");
                                    //list2.Add("Nghỉ Việc riêng");

                                    ReportHelperExcel.SetAlignment(rowC, dem++, tongNgayDiCongTac.ToString(), hStyleConRight);
                                    ReportHelperExcel.SetAlignment(rowC, dem++, (Math.Round(tongGioQuet / 8, 2, MidpointRounding.AwayFromZero)).ToString(), hStyleConRight);
                                    ReportHelperExcel.SetAlignment(rowC, dem++, tongNgayLeNhanVien.ToString(), hStyleConRight);
                                    ReportHelperExcel.SetAlignment(rowC, dem++, tongNgayNghiPhepNamTheoQuiDinh.ToString(), hStyleConRight);
                                    ReportHelperExcel.SetAlignment(rowC, dem++, tongNgayNghiPhepOmCoGiayBenhVien.ToString(), hStyleConRight);
                                    ReportHelperExcel.SetAlignment(rowC, dem++, tongNgayNghiPhepKhongLuong.ToString(), hStyleConRight);
                                    ReportHelperExcel.SetAlignment(rowC, dem++, tongNgayNghiPhepRieng.ToString(), hStyleConRight);
                                    ReportHelperExcel.SetAlignment(rowC, dem++, (
                                            tongNgayDiCongTac
                                            + tongNgayLeNhanVien
                                            + tongNgayNghiPhepNamTheoQuiDinh
                                            + tongNgayNghiPhepOmCoGiayBenhVien
                                            + tongNgayNghiPhepRieng
                                            + tongNgayNghiPhepKhongLuong + Math.Round(tongGioQuet / 8, 2, MidpointRounding.AwayFromZero)).ToString(), hStyleConRight);

                                }
                                //else
                                //{
                                //    ReportHelperExcel.SetAlignment(rowC, dem++, "", hStyleConRight);
                                //    ReportHelperExcel.SetAlignment(rowC, dem++, "", hStyleConRight);
                                //    ReportHelperExcel.SetAlignment(rowC, dem++, "", hStyleConRight);
                                //    ReportHelperExcel.SetAlignment(rowC, dem++, "", hStyleConRight);
                                //}

                                ReportHelperExcel.SetAlignment(rowC, dem++, item.ghiChu, hStyleConCenter);
                            }


                            if (demNVPhongBan == grpNhanVien)
                            {
                                idRowStart = idRowStart + 1;
                                var headerRowNCC = sheet.CreateRow(idRowStart);
                                string cellTitleNgChamCong = Convert.ToString(nguoiChamCong);
                                var titleCellNgChamCong = HSSFCellUtil.CreateCell(headerRowNCC, 1, cellTitleNgChamCong);
                                titleCellNgChamCong.CellStyle = styleTitle13;
                                string cellPhuTrachDonVi = Convert.ToString(phuTrachDonVi);
                                var titleCellPhuTrachDonVi = HSSFCellUtil.CreateCell(headerRowNCC, 10, cellPhuTrachDonVi);
                                titleCellPhuTrachDonVi.CellStyle = styleTitle13;
                                string cellngayThangNam = Convert.ToString(ngayThangNam);
                                var titleCellngayThangNam = HSSFCellUtil.CreateCell(headerRowNCC, 22, cellngayThangNam);
                                titleCellngayThangNam.CellStyle = styleTitle2;
                                idRowStart++;
                                var headerRowNDC = sheet.CreateRow(idRowStart);
                                string cellnguoiDuyetCong = Convert.ToString(nguoiDuyetCong);
                                var titleCellnguoiDuyetCong = HSSFCellUtil.CreateCell(headerRowNDC, 23, cellnguoiDuyetCong);
                                titleCellnguoiDuyetCong.CellStyle = styleTitle13;
                                idRowStart = idRowStart + 2;
                            }
                        }
                    }
                }




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
        #region Xuat File XuatFileBangTongHopCongThang
        public void XuatFileBangTongHopCongThang(string qSearch, string maPhongBan, int thang, int nam, int _page = 0)
        {
            try
            {
                var filename = "";
                var virtualPath = HttpRuntime.AppDomainAppVirtualPath;

                var fileStream = new FileStream(System.Web.HttpContext.Current.Server.MapPath(virtualPath + @"\Content\Report\ReportTemplateTHCT.xls"), FileMode.Open, FileAccess.Read);

                var workbook = new HSSFWorkbook(fileStream, true);

                filename += "BangTongHopCongThang_" + thang + "_" + nam + ".xls";

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



                var sheet = workbook.CreateSheet("BangTongHopCongThang");
                workbook.ActiveSheetIndex = 1;
                //Khai báo row đầu tiên
                int firstRowNumber = 3;

                string cellTenCty = "TỔNG CÔNG TY XDCTGT 6 - CÔNG TY CỔ PHẦN";
                var titleCellCty = HSSFCellUtil.CreateCell(sheet.CreateRow(0), 0, cellTenCty.ToUpper());
                titleCellCty.CellStyle = styleTitle;

                string cellTenCacBanDH = "CÁC BAN ĐIỀU HÀNH DỰ ÁN";
                var titleCellTenCacBanDH = HSSFCellUtil.CreateCell(sheet.CreateRow(1), 1, cellTenCacBanDH.ToUpper());
                titleCellTenCacBanDH.CellStyle = styleTitle;

                string cellTitleMain = "BẢNG TỔNG HỢP CÔNG THÁNG " + thang + "/" + nam;
                var titleCellTitleMain = HSSFCellUtil.CreateCell(sheet.CreateRow(2), 5, cellTitleMain.ToUpper());
                titleCellTitleMain.CellStyle = styleTitle;

                firstRowNumber++;

                var list1 = new List<string>();
                list1.Add("STT");
                list1.Add("Họ tên");
                list1.Add("Mã nhân viên");
                list1.Add("Mã chấm công");
                list1.Add("Phòng ban");
                list1.Add("Chức vụ");
                list1.Add("Công chuẩn");
                list1.Add("Số ngày quét");
                list1.Add("Công chờ việc");
                list1.Add("Công tác");
                list1.Add("Nghỉ phép có lương");
                list1.Add("Nghỉ phép không lương");
                list1.Add("Nghỉ lễ");
                list1.Add("Tăng ca");
                list1.Add("Lũy kế tháng trước");
                list1.Add("Tổng cộng");
                list1.Add("Ghi chú");

                var idRowStart = firstRowNumber; // bat dau o dong thu 4
                var headerRow = sheet.CreateRow(idRowStart);
                int rowend = idRowStart;
                ReportHelperExcel.CreateHeaderRow(headerRow, 0, styleheadedColumnTable, list1);
                //idRowStart++;


                var data = nhanSuContext.sp_NS_BangTongHopCongThang(thang, nam, qSearch, "", 0, maPhongBan).ToList();
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
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.maChamCong, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.phongBan, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.chucVu, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.congChuan, hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.ngayQuet, hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.soNgayCongChoViec, hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.congTac, hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.nghiPhep, hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.nghiPhepKhongLuong, hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.nghiLe, hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowC, dem++, ((item1.tangCaChuNhat ?? 0) + (item1.tangCaLe ?? 0) + (item1.tangCaThuong ?? 0)), hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.luyKeThangTruoc, hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.tongCong, hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.ghiChu, hStyleConLeft);
                }


                idRowStart = idRowStart + 2;
                rowC = sheet.CreateRow(idRowStart);

                var date = DateTime.Now.Day;
                string cellFooterNgayLap = "Tp.Hồ Chí Minh, ngày " + date + " tháng " + thang + " năm " + nam;
                var titleCellFooterNgayLap = HSSFCellUtil.CreateCell(rowC, 8, cellFooterNgayLap);
                titleCellFooterNgayLap.CellStyle = styleTitleItalic;

                ++idRowStart;
                rowC = sheet.CreateRow(idRowStart);
                string cellFooterNguoiSoatCong = "NGƯỜI SOÁT CÔNG";
                var titleCellFooterNguoiSoatCong = HSSFCellUtil.CreateCell(rowC, 3, cellFooterNguoiSoatCong);
                titleCellFooterNguoiSoatCong.CellStyle = styleTitleItalic;

                string cellFooterGDHCNS = "GIÁM ĐỐC HÀNH CHÍNH - NHÂN SỰ";
                var titleCellFooterGDHCNS = HSSFCellUtil.CreateCell(rowC, 8, cellFooterGDHCNS);
                titleCellFooterGDHCNS.CellStyle = styleTitleItalic;

                //idRowStart = idRowStart + 2;
                //string cellFooterPTC = "PHÒNG TỔ CHỨC CB-LĐ";
                //var titleCellFooterPTC = HSSFCellUtil.CreateCell(sheet.CreateRow(idRowStart), 1, cellFooterPTC);
                //titleCellFooterPTC.CellStyle = styleTitle;

                //string cellFooterKT = "PHÒNG TÀI CHÍNH KẾ TOÁN";
                //var titleCellFooterKT = HSSFCellUtil.CreateCell(sheet.GetRow(idRowStart), 7, cellFooterKT);
                //titleCellFooterKT.CellStyle = styleTitle;

                //string cellFooterTGD = "TỔNG GIÁM ĐỐC";
                //var titleCellFooterTGD = HSSFCellUtil.CreateCell(sheet.GetRow(idRowStart), 14, cellFooterTGD);
                //titleCellFooterTGD.CellStyle = styleTitle;
                sheet.SetColumnWidth(0, 5 * 210);
                sheet.SetColumnWidth(1, 30 * 210);
                sheet.SetColumnWidth(2, 15 * 210);
                sheet.SetColumnWidth(3, 15 * 210);
                sheet.SetColumnWidth(4, 60 * 210);
                sheet.SetColumnWidth(5, 20 * 210);
                sheet.SetColumnWidth(6, 15 * 210);
                sheet.SetColumnWidth(7, 15 * 210);
                sheet.SetColumnWidth(8, 15 * 210);
                sheet.SetColumnWidth(9, 15 * 210);
                sheet.SetColumnWidth(10, 15 * 210);
                sheet.SetColumnWidth(11, 15 * 210);
                sheet.SetColumnWidth(12, 15 * 210);
                sheet.SetColumnWidth(13, 15 * 210);
                sheet.SetColumnWidth(14, 15 * 210);
                sheet.SetColumnWidth(15, 30 * 210);
                sheet.SetColumnWidth(16, 30 * 210);

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
        public ActionResult XemBangChamCongTongHop()
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
        public ActionResult LoadBangChamCongTongHop(string qSearch, string maPhongBan, int thang, int nam, int _page = 0)
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
            int total = nhanSuContext.sp_NS_BangTongHopCongThang(thang, nam, qSearch, "", 0, maPhongBan).Count();
            PagingLoaderController("/BangChamCongTongHop/LoadBangChamCongTongHop/", total, page, "?qsearch=" + qSearch);
            ViewData["lsDanhSach"] = nhanSuContext.sp_NS_BangTongHopCongThang(thang, nam, qSearch, "", 0, maPhongBan).Skip(start).Take(offset).ToList();

            ViewData["qSearch"] = qSearch;
            return PartialView("_LoadBangChamCongTongHop");
        }

        public ActionResult LoadBangDanhSachPhieuTangCaDuyet(string qSearch, string maPhongBan, int thang, int nam, int _page = 0)
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

            var dataResults = nhanSuContext.sp_NS_PhieuDangKyTangCa_Duyet(thang, nam, qSearch, null, maPhongBan).ToList();

            //int total = dataResults.Count();
            //PagingLoaderController(string.Empty, total, page, string.Empty);

            ViewData["lsDanhSach"] = dataResults;
            ViewData["qSearch"] = qSearch;

            return PartialView("_LoadBangDanhSachPhieuTangCaDuyet");
        }


        public ActionResult ImportDuLieuVanTay(int thang, int nam, int TuNgay, int DenNgay)
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

                TuNgay = 18;
                DenNgay = 18;
                thang = DateTime.Now.Month;
                nam = DateTime.Now.Year;

                DBChamCongDataContext linqChamCong = new DBChamCongDataContext();
                var lstMayChamCong = linqChamCong.checkinouts.Select(d => d.SN).Distinct().ToList();
                LinqNhanSuDataContext linqNhanSu = new LinqNhanSuDataContext();
                var lstVanTay_MayChamCong = (from p in linqChamCong.userinfos
                                             join q in linqChamCong.checkinouts on p.userid equals q.userid
                                             where (q.checktime.Value.Month == thang && q.checktime.Value.Year == nam && q.checktime.Value.Day >= TuNgay && q.checktime.Value.Day <= DenNgay && lstMayChamCong.Contains(q.SN))
                                             select new
                                             {
                                                 userid = p.userid,
                                                 maChamCong = p.badgenumber,
                                                 maMayChamCong = q.SN,
                                                 idQuet = q.id,
                                                 checktime = q.checktime


                                             }).ToList();

                int soLuotImport = 0;
                foreach (var item in lstVanTay_MayChamCong)
                {
                    var checkTblChamCong_NhanSu = linqNhanSu.tbl_NS_ChamCongs.Where(d => d.idQuet == item.idQuet).FirstOrDefault();
                    if (checkTblChamCong_NhanSu == null)
                    {
                        tbl_NS_ChamCong tblChamCong = new tbl_NS_ChamCong();
                        tblChamCong.checkTime = item.checktime;
                        tblChamCong.maMayChamCong = item.maMayChamCong;
                        tblChamCong.idQuet = item.idQuet;
                        tblChamCong.maChamCong = item.maChamCong;
                        linqNhanSu.tbl_NS_ChamCongs.InsertOnSubmit(tblChamCong);
                        linqNhanSu.SubmitChanges();
                        soLuotImport++;
                    }
                }

                SaveActiveHistory("Import dữ liệu vân tay: " + thang + " năm: " + nam + ". Import được: " + soLuotImport);
                var result = new { kq = true, soLuotImport = soLuotImport };

                return Json(result, JsonRequestBehavior.AllowGet);

            }
            catch
            {
                SaveActiveHistory("Import dữ liệu vân tay: " + thang + " năm: " + nam + " thật bại.");
                var result = new { kq = false };
                return Json(result, JsonRequestBehavior.AllowGet);
            }

        }
        public ActionResult UpdateBangChamCongTH(int thang, int nam)
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

                var list = nhanSuContext.sp_NS_BangTongHopCongThang(thang, nam, null, null, 1, null);


                var result = new { kq = true };
                SaveActiveHistory("Tính công tháng: " + thang + " năm: " + nam);
                return Json(result, JsonRequestBehavior.AllowGet);

            }
            catch
            {
                return View();
            }

        }
        public ActionResult UpdateBangChamCongChiTiet(int thang, int nam)
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
                SaveActiveHistory("Tính công chi tiết: " + thang + " năm: " + nam);
                var list = nhanSuContext.sp_Ns_CapNhatBangCongChiTiet(thang, nam);


                var result = new { kq = true };

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
            ViewData["thangtc"] = new SelectList(dics, "Key", "Value", value);
            ViewData["thangImport"] = new SelectList(dics, "Key", "Value", value);
            ViewData["thang_Tab1"] = new SelectList(dics, "Key", "Value", value);
        }
        private void nam(int value)
        {
            Dictionary<int, string> dics = new Dictionary<int, string>();
            for (int i = 2015; i < 2031; i++)
            {
                dics[i] = i.ToString();
            }
            ViewData["nam"] = new SelectList(dics, "Key", "Value", value);
            ViewData["namtc"] = new SelectList(dics, "Key", "Value", value);
            ViewData["namImport"] = new SelectList(dics, "Key", "Value", value);
            ViewData["nam_Tab1"] = new SelectList(dics, "Key", "Value", value);
        }
        private void TuNgay(int value)
        {
            Dictionary<int, string> dics = new Dictionary<int, string>();
            for (int i = 1; i <= 31; i++)
            {
                dics[i] = i.ToString();
            }
            ViewData["TuNgay"] = new SelectList(dics, "Key", "Value", value);
        }
        private void DenNgay(int value)
        {
            Dictionary<int, string> dics = new Dictionary<int, string>();
            for (int i = 1; i <= 31; i++)
            {
                dics[i] = i.ToString();
            }
            ViewData["DenNgay"] = new SelectList(dics, "Key", "Value", value);
        }

        public ActionResult SaveDuyetTangCa(string soPhieu, string maNhanVien, string ngayTangCa, string duyetGioVao, string duyetGioRa, string duyetSoGio, double? heSoTangCa)
        {
            try
            {
                #region Role user
                permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
                if (!permission.HasValue)
                    return Json("LogIn");
                if (!permission.Value)
                    return Json("AccessDenied");
                #endregion

                LinqPhieuDeNghiDataContext lqPhieuDN = new LinqPhieuDeNghiDataContext();

                var data = lqPhieuDN.tbl_NS_PhieuTangCa_DSNhanViens.Where(d => d.soPhieu == soPhieu && d.maNhanVien == maNhanVien).FirstOrDefault();
                if (data != null)
                {
                    data.ngayCapNhat = DateTime.Now;
                    data.nguoiCapNhat = GetUser().manv;
                    //string date = String.Format("{0:MM/dd/yyyy}", ngayTangCa);
                    data.duyetGioVao = DateTime.ParseExact(ngayTangCa + " " + duyetGioVao, "dd/MM/yyyy HH:mm", CultureInfo.InstalledUICulture);
                    data.duyetGioRa = DateTime.ParseExact(ngayTangCa + " " + duyetGioRa, "dd/MM/yyyy HH:mm", CultureInfo.InstalledUICulture);
                    data.heSoTangCa = heSoTangCa ?? 0;
                    data.duyetSoGio = Convert.ToDouble(!string.IsNullOrEmpty(duyetSoGio) ? duyetSoGio : "0");

                    lqPhieuDN.SubmitChanges();

                    return Json(string.Empty);
                }
                else
                {
                    return Json("Không có nhân viên nào được duyệt.");
                }
            }
            catch
            {
                return Json(ConstantVariale.MESSAGE_ERROR);
            }

        }

        public void DayDuLieuChamCongVe()
        {
            int TuNgay = 18;
            int DenNgay = 18;
            int thang = DateTime.Now.Month;
            int nam = DateTime.Now.Year;

            DBChamCongDataContext linqChamCong = new DBChamCongDataContext();
            LinqNhanSuDataContext linqNhanSu = new LinqNhanSuDataContext();
            var lstVanTay_MayChamCong = linqChamCong.sp_LayChamCongConThieuVeERP(string.Empty, string.Empty, string.Empty, TuNgay, DenNgay, thang, nam).ToList();

            int soLuotImport = 0;

            foreach (var item in lstVanTay_MayChamCong)
            {
                var checkTblChamCong_NhanSu = linqNhanSu.tbl_NS_ChamCongs.Where(d => d.idQuet == item.idQuet && d.checkTime.Value.Year == nam && d.checkTime.Value.Month == thang).FirstOrDefault();

                if (checkTblChamCong_NhanSu == null)
                {
                    tbl_NS_ChamCong tblChamCong = new tbl_NS_ChamCong();

                    tblChamCong.checkTime = Convert.ToDateTime("2019-09-18 08:25:02.000");
                    tblChamCong.maMayChamCong = item.SN;
                    tblChamCong.idQuet = item.idQuet;
                    tblChamCong.maChamCong = item.badgenumber;
                    linqNhanSu.tbl_NS_ChamCongs.InsertOnSubmit(tblChamCong);
                    linqNhanSu.SubmitChanges();
                    soLuotImport++;
                }
            }

            SaveActiveHistory("Import dữ liệu vân tay: " + thang + " năm: " + nam + ". Import được: " + soLuotImport);
            var result = new { kq = true, soLuotImport = soLuotImport };
        }

        #region Import và download file mẫu công tháng
        public void XuatFileBangTongHopCongThangMau(string qSearch, string maPhongBan, int thang, int nam, int _page = 0)
        {
            try
            {
                var filename = "";
                var virtualPath = HttpRuntime.AppDomainAppVirtualPath;

                var fileStream = new FileStream(System.Web.HttpContext.Current.Server.MapPath(virtualPath + @"\Content\Report\ReportTemplateTHCT.xls"), FileMode.Open, FileAccess.Read);

                var workbook = new HSSFWorkbook(fileStream, true);

                filename += "BangTongHopCongThang_" + thang + "_" + nam + ".xls";

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



                var sheet = workbook.CreateSheet("BangTongHopCongThang");
                workbook.ActiveSheetIndex = 1;
                //Khai báo row đầu tiên
                int firstRowNumber = 0;


                var list1 = new List<string>();
                list1.Add("STT");
                list1.Add("Họ tên");
                list1.Add("Mã nhân viên");
                list1.Add("Mã chấm công");
                list1.Add("Phòng ban");
                list1.Add("Chức vụ");
                list1.Add("Công chuẩn");
                list1.Add("Số ngày quét");
                list1.Add("Công chờ việc");
                list1.Add("Công tác");
                list1.Add("Nghỉ phép có lương");
                list1.Add("Nghỉ phép không lương");
                list1.Add("Nghỉ lễ");
                list1.Add("Tăng ca");
                list1.Add("Lũy kế tháng trước");
                list1.Add("Tổng cộng");
                list1.Add("Ghi chú");

                var idRowStart = firstRowNumber; // bat dau o dong thu 4
                var headerRow = sheet.CreateRow(idRowStart);
                int rowend = idRowStart;
                ReportHelperExcel.CreateHeaderRow(headerRow, 0, styleheadedColumnTable, list1);
                //idRowStart++;


                var data = nhanSuContext.sp_NS_BangTongHopCongThang(thang, nam, qSearch, "", 0, maPhongBan).ToList();
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
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.maChamCong, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.phongBan, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.chucVu, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.congChuan, hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.ngayQuet, hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.soNgayCongChoViec, hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.congTac, hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.nghiPhep, hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.nghiPhepKhongLuong, hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.nghiLe, hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowC, dem++, ((item1.tangCaChuNhat ?? 0) + (item1.tangCaLe ?? 0) + (item1.tangCaThuong ?? 0)), hStyleConRight);

                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.luyKeThangTruoc, hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.tongCong, hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.ghiChu, hStyleConLeft);
                }

                sheet.SetColumnWidth(0, 5 * 210);
                sheet.SetColumnWidth(1, 30 * 210);
                sheet.SetColumnWidth(2, 15 * 210);
                sheet.SetColumnWidth(3, 15 * 210);
                sheet.SetColumnWidth(4, 60 * 210);
                sheet.SetColumnWidth(5, 20 * 210);
                sheet.SetColumnWidth(6, 15 * 210);
                sheet.SetColumnWidth(7, 15 * 210);
                sheet.SetColumnWidth(8, 15 * 210);
                sheet.SetColumnWidth(9, 15 * 210);
                sheet.SetColumnWidth(10, 15 * 210);
                sheet.SetColumnWidth(11, 15 * 210);
                sheet.SetColumnWidth(12, 15 * 210);
                sheet.SetColumnWidth(13, 15 * 210);
                sheet.SetColumnWidth(14, 15 * 210);
                sheet.SetColumnWidth(15, 30 * 210);
                sheet.SetColumnWidth(16, 30 * 210);
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

        public ActionResult ImportFileTC(int? thang, int? nam)
        {
            try
            {
                string[] supportedFiles = { ".xlsx", ".xls" };
                HttpPostedFileBase File;
                File = Request.Files[0];
                if (File.ContentLength > 0)
                {
                    string extension = Path.GetExtension(File.FileName);
                    bool exist = Array.Exists(supportedFiles, element => element == extension);
                    if (exist == false)
                    {
                        return Json(new { success = false });
                    }
                    else
                    {
                        var date = DateTime.Now.ToString("yyyyMMdd-HHMMss");
                        string savedLocation = "/UploadFiles/NhanVien/";
                        Directory.CreateDirectory(savedLocation);
                        var filePath = Server.MapPath(savedLocation);
                        string fileName = date.ToString() + File.FileName;
                        string savedFileName = Path.Combine(filePath, fileName);
                        File.SaveAs(savedFileName);

                        ExcelDataProcessing excelDataProcessor = new ExcelDataProcessing(savedFileName);
                        DataTable dt = excelDataProcessor.GetDataTableWorkSheet("BangTongHopCongThang");
                        List<tbl_NS_BangTongHopCongThang> listTHC = new List<tbl_NS_BangTongHopCongThang>();
                        tbl_NS_BangTongHopCongThang congTH;
                        foreach (DataRow row in dt.Rows)
                        {

                            if (!String.IsNullOrEmpty(row["Mã nhân viên"].ToString()))
                            {
                                var tongHopCong = nhanSuContext.tbl_NS_BangTongHopCongThangs.Where(d => d.thang == thang && d.nam == nam && d.maNhanVien.Trim().ToLower() == row["Mã nhân viên"].ToString().Trim().ToLower()).FirstOrDefault();
                                if (tongHopCong != null)
                                {
                                    tongHopCong.congChuan = string.IsNullOrEmpty(row["Công chuẩn"].ToString()) ? 0 : Convert.ToDouble(row["Công chuẩn"].ToString());
                                    tongHopCong.ngayQuet = string.IsNullOrEmpty(row["Số ngày quét"].ToString()) ? 0 : Convert.ToDouble(row["Số ngày quét"].ToString());
                                    tongHopCong.congTac = string.IsNullOrEmpty(row["Công tác"].ToString()) ? 0 : Convert.ToDouble(row["Công tác"].ToString());
                                    tongHopCong.nghiPhep = string.IsNullOrEmpty(row["Nghỉ phép có lương"].ToString()) ? 0 : Convert.ToDouble(row["Nghỉ phép có lương"].ToString());
                                    tongHopCong.nghiPhepKhongLuong = string.IsNullOrEmpty(row["Nghỉ phép không lương"].ToString()) ? 0 : Convert.ToDouble(row["Nghỉ phép không lương"].ToString());
                                    tongHopCong.nghiLe = string.IsNullOrEmpty(row["Nghỉ lễ"].ToString()) ? 0 : Convert.ToDouble(row["Nghỉ lễ"].ToString());
                                    tongHopCong.luyKeThangTruoc = string.IsNullOrEmpty(row["Lũy kế tháng trước"].ToString()) ? 0 : Convert.ToDouble(row["Lũy kế tháng trước"].ToString());
                                    tongHopCong.tongCong = string.IsNullOrEmpty(row["Tổng cộng"].ToString()) ? 0 : Math.Round(Convert.ToDouble(row["Tổng cộng"].ToString()), 2, MidpointRounding.ToEven);
                                    tongHopCong.ngayQuet = string.IsNullOrEmpty(row["Tổng cộng"].ToString()) ? 0 : Math.Round(Convert.ToDouble(row["Tổng cộng"].ToString()), 2, MidpointRounding.ToEven);
                                    nhanSuContext.SubmitChanges();
                                }
                                else
                                {
                                    congTH = new tbl_NS_BangTongHopCongThang();
                                    congTH.thang = thang ?? DateTime.Now.Month;
                                    congTH.nam = nam ?? DateTime.Now.Year;
                                    congTH.congChuan = string.IsNullOrEmpty(row["Công chuẩn"].ToString()) ? 0 : Convert.ToDouble(row["Công chuẩn"].ToString());
                                    congTH.ngayQuet = string.IsNullOrEmpty(row["Số ngày quét"].ToString()) ? 0 : Convert.ToDouble(row["Số ngày quét"].ToString());
                                    congTH.soNgayCongChoViec = string.IsNullOrEmpty(row["Công chờ việc"].ToString()) ? 0 : Convert.ToDouble(row["Công chờ việc"].ToString());
                                    congTH.congTac = string.IsNullOrEmpty(row["Công tác"].ToString()) ? 0 : Convert.ToDouble(row["Công tác"].ToString());
                                    congTH.nghiPhep = string.IsNullOrEmpty(row["Nghỉ phép có lương"].ToString()) ? 0 : Convert.ToDouble(row["Nghỉ phép có lương"].ToString());
                                    congTH.nghiPhepKhongLuong = string.IsNullOrEmpty(row["Nghỉ phép không lương"].ToString()) ? 0 : Convert.ToDouble(row["Nghỉ phép không lương"].ToString());
                                    congTH.nghiLe = string.IsNullOrEmpty(row["Nghỉ lễ"].ToString()) ? 0 : Convert.ToDouble(row["Nghỉ lễ"].ToString());
                                    congTH.luyKeThangTruoc = string.IsNullOrEmpty(row["Lũy kế tháng trước"].ToString()) ? 0 : Convert.ToDouble(row["Lũy kế tháng trước"].ToString());
                                    congTH.tongCong = string.IsNullOrEmpty(row["Tổng cộng"].ToString()) ? 0 : Math.Round(Convert.ToDouble(row["Tổng cộng"].ToString()), 2, MidpointRounding.ToEven);
                                    congTH.ngayQuet = string.IsNullOrEmpty(row["Tổng cộng"].ToString()) ? 0 : Math.Round(Convert.ToDouble(row["Tổng cộng"].ToString()), 2, MidpointRounding.ToEven);
                                    congTH.ghiChu = row["Ghi chú"].ToString().Trim();
                                    congTH.hoTen = row["Họ tên"].ToString().Trim();
                                    congTH.maNhanVien = row["Mã nhân viên"].ToString().Trim();
                                    congTH.maChamCong = row["Mã chấm công"].ToString().Trim();
                                    congTH.phongBan = row["Phòng ban"].ToString().Trim();
                                    congTH.chucVu = row["Chức vụ"].ToString().Trim();
                                    listTHC.Add(congTH);
                                }
                            }
                        }
                        if (listTHC != null && listTHC.Count > 0)
                        {
                            nhanSuContext.tbl_NS_BangTongHopCongThangs.InsertAllOnSubmit(listTHC);
                            nhanSuContext.SubmitChanges();
                        }
                        // System.IO.File.Delete(Server.MapPath("/UploadFiles/NhanVien/" + fileName));
                    }
                }
                SaveActiveHistory("Import file điều chỉnh công tổng cộng tháng " + thang + " năm " + nam);
                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                return Json(new { success = false });
            }
        }
        #endregion

        #region Điều chỉnh công
        public JsonResult DieuChinhCong(string maNhanVienDieuChinh, int thangDieuChinh, int namDieuChinh, float? congDieuChinh)
        {
            try
            {
                var capNhatCongDC = nhanSuContext.tbl_NS_BangTongHopCongThangs.Where(d => d.maNhanVien == maNhanVienDieuChinh && d.thang == thangDieuChinh && d.nam == namDieuChinh).FirstOrDefault();
                capNhatCongDC.tongCong = Math.Round(congDieuChinh ?? 0, 2, MidpointRounding.ToEven);
                capNhatCongDC.ngayQuet = Math.Round(congDieuChinh ?? 0, 2, MidpointRounding.ToEven);
                nhanSuContext.SubmitChanges();
                return Json(string.Empty);
            }
            catch
            {
                return Json("Error");
            }
        }
        #endregion

    }
}
