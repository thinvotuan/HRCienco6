using BatDongSan.Helper.Common;
using BatDongSan.Helper.Utils;
using BatDongSan.Models.NhanSu;
using BatDongSan.Models.QLSuatAn;
using BatDongSan.Models.QLSuatAn.LuuVet;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.UserModel.Contrib;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using Worldsoft.Mvc.Web.Util;

namespace BatDongSan.Controllers.QLSuatAn
{
    public class BaoCaoThongKeSuatAnTheoNhanVienController : ApplicationController
    {
        //
        // GET: /BaoCaoThongKeSuatAnTheoNhanVien/
        private bool? permission;
        private string MCV = "BCTKSATNV";
        private int rowNumerInPage = 50;

        QLSuatAnDataContext suatAnContext = new QLSuatAnDataContext();
        QLSALuuVetDataContext context = new QLSALuuVetDataContext();

        public ActionResult Index(string maCongTy, string tenPhongBan, string maCuaHang, string tuNgay, string denNgay, string qSearch, int _page = 0)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            int page = (_page == 0 ? 1 : _page);

            DateTime? fromDate = null;
            DateTime? toDate = null;
            if (!String.IsNullOrEmpty(tuNgay))
                fromDate = DateTime.ParseExact(tuNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            if (!String.IsNullOrEmpty(denNgay))
                toDate = DateTime.ParseExact(denNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);


            var objectDatas = context.sp_SA_BC_ThongKeSuatAnTungNhanVien(System.Configuration.ConfigurationManager.AppSettings["CongTy"].ToString(), maCuaHang, fromDate, toDate, qSearch, null, null).ToList();


            var nhanViens = suatAnContext.sp_SA_DangKySuatAn_CongChuanNhanVien_Cienco6((toDate ?? DateTime.Now).Month, (toDate ?? DateTime.Now).Year)                
                .ToList();

            if (nhanViens != null && nhanViens.Count() > 0 && !string.IsNullOrEmpty(qSearch))
            {
                nhanViens = nhanViens.Where(d => d.maNhanVien.Contains(qSearch)
                                                || d.tenNhanVien.Contains(qSearch) || d.tenPhongBan.Contains(qSearch)).ToList();
            }

            if (!string.IsNullOrEmpty(tenPhongBan))
            {
                nhanViens = nhanViens.Where(d => d.tenPhongBan == tenPhongBan).ToList();
            }

            ViewBag.luuVetData = objectDatas;
            ViewBag.suatAnData = nhanViens;

            if (Request.IsAjaxRequest())
            {
                return PartialView("ViewIndex");
            }

            CuaHangs(string.Empty);

            StringBuilder buildTree = new StringBuilder();
            var phongBans = suatAnContext.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan>().ToList();
            buildTree = TreePhongBanStyle.BuildTreeDepartment(phongBans);

            ViewBag.PhongBans = buildTree.ToString();

            return View();
        }


        private void CongTys(string value)
        {
            Dictionary<string, string> dics = new Dictionary<string, string>();

            dics[""] = "Tất cả";
            dics["ThuanViet"] = "ThuanViet";
            dics["TVWindow"] = "TVWindow";
            dics["NewCity"] = "NewCity";
            dics["Cienco6"] = "Cienco6";

            ViewBag.CongTys = new SelectList(dics, "Key", "Value", value);
        }


        private void CuaHangs(string value)
        {
            Dictionary<string, string> dics = new Dictionary<string, string>();

            dics[""] = "Tất cả";

            var cuaHangs = context.tbl_SA_CuaHangs.Select(d => new { maCuaHang = d.maCuaHang, tenCuaHang = d.tenCuaHang }).ToList();

            foreach (var item in cuaHangs)
            {
                dics[item.maCuaHang] = item.tenCuaHang;
            }

            ViewBag.CuaHangs = new SelectList(dics, "Key", "Value", value);
        }


        public void XuatFileNhanVien(string maCongTy, string tenPhongBan, string maCuaHang, string tuNgay, string denNgay, string qSearch)
        {
            var filename = "";
            var virtualPath = HttpRuntime.AppDomainAppVirtualPath;
            var fileStream = new FileStream(System.Web.HttpContext.Current.Server.MapPath(virtualPath + @"\Content\Report\ReportTemplate.xls"), FileMode.Open, FileAccess.Read);

            var workbook = new HSSFWorkbook(fileStream, true);
            filename += "NhanVien" + string.Format("{0:dd/MM/yyyy}", DateTime.Now) + ".xls";


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


            DateTime? fromDate = null;
            DateTime? toDate = null;
            if (!String.IsNullOrEmpty(tuNgay))
                fromDate = DateTime.ParseExact(tuNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            if (!String.IsNullOrEmpty(denNgay))
                toDate = DateTime.ParseExact(denNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);

            Row rowC = null;
            //Khai báo row đầu tiên
            int firstRowNumber = 1;

            string rowtitle = "BẢNG KÊ SUẤT ĂN";
            var titleCell = HSSFCellUtil.CreateCell(sheet.CreateRow(firstRowNumber), 2, rowtitle.ToUpper());
            titleCell.CellStyle = styleTitle;

            firstRowNumber = firstRowNumber + 1;

            string rowtitle_NhaHang = "NHÀ HÀNG : " + context.tbl_SA_CuaHangs.Where(d=>d.maCuaHang == maCuaHang).Select(d=>d.tenCuaHang).FirstOrDefault() ?? "TẤT CẢ";
            var titleCell_NhaHang = HSSFCellUtil.CreateCell(sheet.CreateRow(firstRowNumber), 2, rowtitle_NhaHang.ToUpper());
            titleCell_NhaHang.CellStyle = styleTitle;

            firstRowNumber = firstRowNumber + 1;
            string rowtitle1 = "Từ Ngày : " + string.Format("{0:dd/MM/yyyy}", fromDate) + "- Đến ngày : " + string.Format("{0:dd/MM/yyyy}", toDate);
            var titleCell1 = HSSFCellUtil.CreateCell(sheet.CreateRow(firstRowNumber), 2, rowtitle1);
            titleCell1.CellStyle = styleTitle;

            firstRowNumber = firstRowNumber + 2;

            var list1 = new List<string>();
            list1.Add("STT");
            list1.Add("Mã nhân viên");
            list1.Add("Tên nhân viên");
            list1.Add("Tên phòng ban");
            list1.Add("Công chuẩn gần đây");
            list1.Add("Số phiếu đăng ký");

            //Start row 13
            var headerRow = sheet.CreateRow(6);
            ReportHelperExcel.CreateHeaderRow(headerRow, 0, styleheadedColumnTable, list1);

            //Create header end

            var idRowStart = 6;

            var luuVetData = context.sp_SA_BC_ThongKeSuatAnTungNhanVien(System.Configuration.ConfigurationManager.AppSettings["CongTy"].ToString(), maCuaHang, fromDate, toDate, qSearch, null, null).ToList();

            var nhanViens = suatAnContext.sp_SA_DangKySuatAn_CongChuanNhanVien_Cienco6((toDate ?? DateTime.Now).Month, (toDate ?? DateTime.Now).Year).ToList();

            if (nhanViens != null && nhanViens.Count() > 0 && !string.IsNullOrEmpty(qSearch))
            {
                nhanViens = nhanViens.Where(d => d.maNhanVien.Contains(qSearch)
                                                || d.tenNhanVien.Contains(qSearch) || d.tenPhongBan.Contains(qSearch)).ToList();
            }

            if (!string.IsNullOrEmpty(tenPhongBan))
            {
                nhanViens = nhanViens.Where(d => d.tenPhongBan == tenPhongBan).ToList();
            }

            //#region
            if (nhanViens != null && nhanViens.Count > 0)
            {
                var stt = 0;
                int dem = 0;
                int tongSoSuat = 0;
                BatDongSan.Models.QLSuatAn.LuuVet.sp_SA_BC_ThongKeSuatAnTungNhanVienResult luuVet = null;
                //Giai đoạn
                foreach (var item1 in nhanViens)
                {
                    if (luuVetData != null && luuVetData.Count() > 0)
                    {
                        luuVet = luuVetData.Where(d => d.maNhanVien == item1.maNhanVien).FirstOrDefault();
                    }

                    if (luuVet == null)
                    {
                        luuVet = new BatDongSan.Models.QLSuatAn.LuuVet.sp_SA_BC_ThongKeSuatAnTungNhanVienResult();
                    }

                    dem = 0;
                    stt++;
                    idRowStart++;
                    rowC = sheet.CreateRow(idRowStart);
                    ReportHelperExcel.SetAlignment(rowC, dem++, stt.ToString(), hStyleConCenter);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.maNhanVien, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.tenNhanVien, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item1.tenPhongBan, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, (item1.congChuan ?? 0).ToString(), hStyleConCenter);
                    ReportHelperExcel.SetAlignment(rowC, dem++, (luuVet.soPhieuDangKyTuNgayDenNgay ?? 0).ToString(), hStyleConCenter);

                    tongSoSuat += luuVet.soPhieuDangKyTuNgayDenNgay ?? 0;
                    //ReportHelperExcel.SetAlignment(rowC, dem++, item1.tenMonAn, hStyleConLeft);
                }

                dem = 0;
                idRowStart++;
                rowC = sheet.CreateRow(idRowStart);
                ReportHelperExcel.SetAlignment(rowC, dem++, string.Empty, hStyleConCenter);
                ReportHelperExcel.SetAlignment(rowC, dem++, string.Empty, hStyleConLeft);
                ReportHelperExcel.SetAlignment(rowC, dem++, string.Empty, hStyleConLeft);
                ReportHelperExcel.SetAlignment(rowC, dem++, string.Empty, hStyleConLeft);
                ReportHelperExcel.SetAlignment(rowC, dem++, string.Empty, hStyleConCenter);
                ReportHelperExcel.SetAlignment(rowC, dem++, tongSoSuat.ToString(), hStyleConCenter);

                sheet.SetColumnWidth(0, 8 * 210);
                sheet.SetColumnWidth(1, 30 * 210);
                sheet.SetColumnWidth(2, 30 * 210);
                sheet.SetColumnWidth(3, 30 * 210);
                sheet.SetColumnWidth(2, 30 * 210);
                sheet.SetColumnWidth(3, 30 * 210);
            }
            else
            {
                sheet.SetColumnWidth(0, 8 * 210);
                sheet.SetColumnWidth(1, 30 * 210);
                sheet.SetColumnWidth(2, 30 * 210);
                sheet.SetColumnWidth(3, 30 * 210);
                sheet.SetColumnWidth(2, 30 * 210);
                sheet.SetColumnWidth(3, 30 * 210);
            }

            var stream = new MemoryStream();
            workbook.Write(stream);

            Response.ContentType = "application/vnd.ms-excel";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));
            Response.Clear();

            Response.BinaryWrite(stream.GetBuffer());
            Response.End();

        }
    }
}
