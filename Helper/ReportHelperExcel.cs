using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;

namespace Worldsoft.Mvc.Web.Util
{
    public static class ReportHelperExcel
    {
        /// <summary>
        /// Thiết lập thông tin cho file báo cáo
        /// </summary>
        public static void InitializeWorkbook(HSSFWorkbook workbook, string title)
        {
            var documentSummaryInformation = PropertySetFactory.CreateDocumentSummaryInformation();

            documentSummaryInformation.Company = "ThuanViet Corporation";
            workbook.DocumentSummaryInformation = documentSummaryInformation;

            var summaryInformation = PropertySetFactory.CreateSummaryInformation();

            summaryInformation.Subject = title;
            workbook.SummaryInformation = summaryInformation;
        }

        //[2017-11-14]TrungIT - Modify Start
        /// <summary>
        /// Thiết lập và canh giá trị của cell
        /// </summary>
        //public static Cell SetAlignment(Row row, int rowId, string value, CellStyle cellStyle)
        //{
        //    var cell = row.CreateCell(rowId);

        //    if (!String.IsNullOrEmpty(value))
        //    {
        //        cell.SetCellValue(value.Trim());
        //    }
        //    else
        //    {
        //        cell.SetCellValue(String.Empty);
        //    }
        //    if (cellStyle != null)
        //    {
        //        cell.CellStyle = cellStyle;
        //    }

        //    return cell;
        //}

        public static Cell SetAlignment(Row row, int rowId, object value, CellStyle cellStyle)
        {
            var cell = row.CreateCell(rowId);

            if (value is byte || value is ushort || value is short ||
                 value is uint || value is int || value is ulong || value is long)
            {
                cell.SetCellValue((Int64)value);
            }

            else if (value is float || value is double)
            {
                cell.SetCellValue((double)value);
            }
            else
            {
                // strings and anything else should be text
                if (!String.IsNullOrEmpty((string)value))
                {
                    cell.SetCellValue(((string)value).Trim());
                }
                else
                {
                    cell.SetCellValue(String.Empty);
                }
            }

            if (cellStyle != null)
            {
                cell.CellStyle = cellStyle;
            }

            return cell;
        }
        //[2017-11-14 ]TrungIT - Modify End
        private static string GetDefaultDataFormat(object value, ref Cell cell)
        {
            if (value == null)
            {
                cell.SetCellValue((string)value);
                return "General";
            }

            else if (value is DateTime)
            {
                cell.SetCellValue((DateTime)value);
                return "m/d/yy h:mm";
            }

            else if (value is bool)
            {
                cell.SetCellValue((bool)value);
                return "[=0]\"Yes\";[=1]\"No\"";
            }

            else if (value is byte || value is ushort || value is short ||
                 value is uint || value is int || value is ulong || value is long)
            {
                cell.SetCellValue((Int64)value);
                return "0";
            }

            else if (value is float || value is double)
            {
                cell.SetCellValue((double)value);
                return "#,##0.0000";
            }
            else
            {
                // strings and anything else should be text
                cell.SetCellValue((string)value);
                return "text";
            }
        }

        private static readonly Dictionary<string, short> _cellStyleCache = new Dictionary<string, short>();

        private static short GetCellStyleForFormat(Workbook workbook, string dataFormat)
        {
            if (!_cellStyleCache.ContainsKey(dataFormat))
            {
                var newDataFormat = workbook.CreateDataFormat();
                var style = workbook.CreateCellStyle();

                style.DataFormat = newDataFormat.GetFormat(dataFormat);

                _cellStyleCache[dataFormat] = style.DataFormat;
            }

            return _cellStyleCache[dataFormat];
        }

        public static Cell SetAlignment(Row row, int rowId, object value, CellStyle cellStyle, Workbook workbook)
        {
            var cell = row.CreateCell(rowId);

            if (cellStyle != null)
            {
                cell.CellStyle = cellStyle;
            }


            if (value == null)
            {
                cell.SetCellValue(string.Empty);
                return cell;
            }

            //cell.SetCellValue(value);

            // get format from the column definition ("m/d", "##.###", etc.), or use the default
            string dataFormat = GetDefaultDataFormat(value, ref cell);

            // find/create cell style
            cell.CellStyle.DataFormat = GetCellStyleForFormat(workbook, dataFormat);

            return cell;
        }
        //[2017-11-14 ]TrungIT - Modify Start
        //public static Cell SetAlignment(Row row, int rowId, string value, CellStyle cellStyle, Font fontName)
        //{
        //    var cell = row.CreateCell(rowId);
        //    cellStyle.SetFont(fontName);
        //    if (!String.IsNullOrEmpty(value))
        //    {
        //        cell.SetCellValue(value.Trim());
        //    }
        //    else
        //    {
        //        cell.SetCellValue(String.Empty);
        //    }
        //    if (cellStyle != null)
        //    {
        //        cell.CellStyle = cellStyle;
        //    }

        //    return cell;
        //}

        public static Cell SetAlignment(Row row, int rowId, object value, CellStyle cellStyle, Font fontName)
        {
            var cell = row.CreateCell(rowId);
            cellStyle.SetFont(fontName);

            if (value is byte || value is ushort || value is short ||
                 value is uint || value is int || value is ulong || value is long)
            {
                cell.SetCellValue((Int64)value);
            }

            else if (value is float || value is double)
            {
                cell.SetCellValue((double)value);
            }
            else
            {
                // strings and anything else should be text
                if (!String.IsNullOrEmpty((string)value))
                {
                    cell.SetCellValue(((string)value).Trim());
                }
                else
                {
                    cell.SetCellValue(String.Empty);
                }
            }

            if (cellStyle != null)
            {
                cell.CellStyle = cellStyle;
            }

            return cell;
        }
        //[2017-11-14 ]TrungIT - Modify End
        public static void CreateHeaderRow(Row row, int cellId, CellStyle cellStyle, IEnumerable<string> list, Font fontName)
        {
            var i = cellId;

            cellStyle.Alignment = HorizontalAlignment.CENTER;
            cellStyle.VerticalAlignment = VerticalAlignment.CENTER;
            cellStyle.FillForegroundColor = HSSFColor.DARK_BLUE.index;
            cellStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
            cellStyle.SetFont(fontName);
            foreach (var item in list)
            {
                var cell = row.CreateCell(i);
                cell.SetCellValue(item);
                cell.CellStyle = cellStyle;
                i++;
            }
        }

        public static void CreateHeaderRow(Row row, int cellId, CellStyle cellStyle, IEnumerable<string> list)
        {
            var i = cellId;

            cellStyle.Alignment = HorizontalAlignment.CENTER;
            cellStyle.VerticalAlignment = VerticalAlignment.CENTER;
            cellStyle.FillForegroundColor = HSSFColor.WHITE.index;
            cellStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;

            foreach (var item in list)
            {
                var cell = row.CreateCell(i);
                cell.SetCellValue(item);
                cell.CellStyle = cellStyle;
                i++;
            }
        }

        /// <summary>
        /// son.le - 2011.07.14: tạo cell và insert dữ liệu
        /// </summary>
        /// <param name="row"></param>
        /// <param name="cellId"></param>
        /// <param name="cellStyle"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static Cell CreateCell(this Row row, int cellId, CellStyle cellStyle, string value)
        {
            var cell = row.CreateCell(cellId);

            if (!String.IsNullOrEmpty(value))
                cell.SetCellValue(value);
            cell.RemoveCellComment();

            if (cellStyle != null)
            {
                cell.CellStyle = cellStyle;
            }

            return cell;
        }

        /// <summary>
        /// son.le - 2011.07.14: tạo dòng header cho báo cáo
        /// </summary>
        /// <param name="list"></param>
        public static void CreateRowHeader(this Sheet sheet, int rowId, int cellStartId, IEnumerable<string> list)
        {
            if (sheet == null) return;
            var headerRow = sheet.CreateRow(rowId);
            headerRow.Height = (short)(100 * 5.2);

            //định dạng cho dòng header
            var headerFont = sheet.Workbook.CreateFont();
            headerFont.Color = HSSFColor.WHITE.index;
            headerFont.Boldweight = 100 * 10;

            var cellStyle = sheet.Workbook.CreateCellStyle();

            cellStyle.SetFont(headerFont);

            cellStyle.VerticalAlignment = VerticalAlignment.CENTER;
            cellStyle.Alignment = HorizontalAlignment.CENTER;
            cellStyle.FillForegroundColor = HSSFColor.DARK_BLUE.index;
            cellStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
            cellStyle.WrapText = true;

            cellStyle.BorderBottom = CellBorderType.THIN;
            cellStyle.BorderLeft = CellBorderType.THIN;
            cellStyle.BorderRight = CellBorderType.THIN;
            cellStyle.BorderTop = CellBorderType.THIN;

            cellStyle.BottomBorderColor = HSSFColor.DARK_BLUE.index;
            cellStyle.LeftBorderColor = HSSFColor.DARK_BLUE.index;
            cellStyle.RightBorderColor = HSSFColor.DARK_BLUE.index;
            cellStyle.TopBorderColor = HSSFColor.DARK_BLUE.index;

            var col = cellStartId - 1;
            foreach (var item in list)
            {
                var cell = headerRow.CreateCell(++col);
                cell.SetCellValue(item);
                cell.CellStyle = cellStyle;
            }
        }

        public static CellStyle CreateCellStyle(this Sheet sheet, char? align)
        {
            if (sheet == null) return null;
            var cellStyle = sheet.Workbook.CreateCellStyle();

            cellStyle.BorderBottom = CellBorderType.THIN;
            cellStyle.BorderLeft = CellBorderType.THIN;
            cellStyle.BorderRight = CellBorderType.THIN;
            cellStyle.BorderTop = CellBorderType.THIN;
            cellStyle.BottomBorderColor = HSSFColor.BLACK.index;
            cellStyle.LeftBorderColor = HSSFColor.BLACK.index;
            cellStyle.RightBorderColor = HSSFColor.BLACK.index;
            cellStyle.TopBorderColor = HSSFColor.BLACK.index;

            cellStyle.VerticalAlignment = VerticalAlignment.CENTER;
            cellStyle.WrapText = true;

            //TẠO ALIGN
            if (align != null)
            {
                switch (align)
                {
                    case 'L':
                        cellStyle.Alignment = HorizontalAlignment.LEFT;
                        break;
                    case 'R':
                        cellStyle.Alignment = HorizontalAlignment.RIGHT;
                        break;
                    case 'C':
                        cellStyle.Alignment = HorizontalAlignment.CENTER;
                        break;
                }

            }

            return cellStyle;
        }
        //Định dạng style dòng tổng trong danh sách
        public static CellStyle CreateFinalCellStyle(this Sheet sheet)
        {
            if (sheet == null) return null;
            var rowFinalStyle = sheet.Workbook.CreateCellStyle();

            rowFinalStyle.BorderBottom = CellBorderType.THIN;
            rowFinalStyle.BorderLeft = CellBorderType.THIN;
            rowFinalStyle.BorderRight = CellBorderType.THIN;
            rowFinalStyle.BorderTop = CellBorderType.THIN;

            rowFinalStyle.BottomBorderColor = HSSFColor.BLACK.index;
            rowFinalStyle.LeftBorderColor = HSSFColor.BLACK.index;
            rowFinalStyle.RightBorderColor = HSSFColor.BLACK.index;
            rowFinalStyle.TopBorderColor = HSSFColor.BLACK.index;
            rowFinalStyle.Alignment = HorizontalAlignment.RIGHT;


            //ĐỊNH DẠNG FONT CHO DÒNG TỔNG GIÁ TRỊ, DIỆN TÍCH
            var fontFinalRow = sheet.Workbook.CreateFont();

            fontFinalRow.FontHeightInPoints = 10;
            fontFinalRow.Boldweight = 100 * 10;
            fontFinalRow.FontName = "Times New Roman";
            fontFinalRow.IsItalic = true;
            fontFinalRow.Color = HSSFColor.BLACK.index;

            //ĐỊNH DẠNG FONT CHO DÒNG CUỐI
            rowFinalStyle.SetFont(fontFinalRow);
            return rowFinalStyle;
        }
    }
}