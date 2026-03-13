using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Utils;
using BatDongSan.Models.NhanSu;
using BatDongSan.Models.PhieuDeNghi;
using BatDongSan.Models.HeThong;
using BatDongSan.Models.DanhMuc;
using System.Text;
using BatDongSan.Utils.Paging;
using BatDongSan.Helper.Common;
using System.Globalization;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel.Contrib;
using Worldsoft.Mvc.Web.Util;

namespace BatDongSan.Controllers.NhanSu
{
    public class PhepNamNhanVienController : ApplicationController
    {
        LinqPhieuDeNghiDataContext lqPhieuDN = new LinqPhieuDeNghiDataContext();
        LinqNhanSuDataContext linqNS = new LinqNhanSuDataContext();
        LinqDanhMucDataContext linqDM = new LinqDanhMucDataContext();

        tbl_DN_NghiPhepNhanVien phepNam;
        IList<tbl_DN_NghiPhepNhanVien> listPhepNam;
        PhepNamNhanVien PhepNamModel;
        List<DanhSachNhanVien> ListNhanVienChon;
        private StringBuilder buildTree = null;
        private IList<tbl_DM_PhongBan> phongBans;
        public const string taskIDSystem = "PhepNamNhanVien";
        public bool? permission;
        //
        // GET: /PhepNamNhanVien/        

        public ActionResult Index(int? page, int? pageSize, string searchString, string maPhongBan, int? Nam)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            
            if (Nam == null)
            {
                Nam = DateTime.Now.Year;
            }
            int currentPageIndex = page.HasValue ? page.Value : 1;
            pageSize = pageSize ?? 30;
            IList<sp_PhepNamNhanVien_IndexResult> list;
            using (linqNS = new LinqNhanSuDataContext())
            {

                list = linqNS.sp_PhepNamNhanVien_Index(Nam, maPhongBan, searchString).ToList();
                ViewBag.Count = list.Count();
                ViewBag.searchString = searchString;
                ViewBag.maPhongBan = maPhongBan;
                ViewBag.Nam = Nam;
                TempData["Params"] = searchString + "," + maPhongBan + "," + Nam;
            }
            if (Request.IsAjaxRequest())
            {
                return PartialView("ViewIndex", list.ToPagedList(currentPageIndex, 30));
            }
            else
            {
                //
                buildTree = new StringBuilder();
                phongBans = linqDM.tbl_DM_PhongBans.ToList();
                buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
                ViewBag.PhongBans = buildTree.ToString();
                //
                BuildNamLamViec(DateTime.Now.Year, 20, DateTime.Now.Year);

                var namTonDauTien = lqPhieuDN.tbl_DN_NghiPhepNhanViens.Where(d => (d.isDauKy ?? false)).Select(d => d.nam).FirstOrDefault();
                if (namTonDauTien == 0)
                {
                    namTonDauTien = DateTime.Now.Year + 1;
                }
                else
                {
                    ViewBag.namTon = namTonDauTien.ToString();
                    namTonDauTien = namTonDauTien + 1;
                }
                BuildNamTonDauKy(namTonDauTien, 20, null);
            }
            return View(list.ToPagedList(currentPageIndex, 30));
        }


        public ActionResult Index_Tab2(int? page, int? pageSize, string searchString, string maPhongBan, string ngayKiemTra)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            
            int currentPageIndex = page.HasValue ? page.Value : 1;
            pageSize = pageSize ?? 100;
            DateTime? dtDenNgay;

            if (!string.IsNullOrEmpty(ngayKiemTra))
            {
                dtDenNgay = DateTime.ParseExact(ngayKiemTra, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            }
            else
            {
                dtDenNgay = DateTime.Now;
            }


            List<sp_PhepTonNamResult> list = new List<sp_PhepTonNamResult>();

            using (linqNS = new LinqNhanSuDataContext())
            {

                list = linqNS.sp_PhepTonNam(dtDenNgay, searchString, maPhongBan).ToList();
                PagingLoaderController(string.Empty, list.Count(), currentPageIndex, string.Empty);
            }

            ViewBag.NgayKiemTra = dtDenNgay;

            ViewData["lsDanhSach"] = list.Skip(start).Take(offset).ToList();

            return PartialView("ViewIndex_Tab2");
        }


        public ActionResult Index_Tab3(int? page, int? pageSize, string searchString, string maPhongBan)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            int currentPageIndex = page.HasValue ? page.Value : 1;
            pageSize = pageSize ?? 100;


            List<sp_PhepNamNhanVien_IndexResult> list = new List<sp_PhepNamNhanVien_IndexResult>();

            using (linqNS = new LinqNhanSuDataContext())
            {
                int namTon = lqPhieuDN.tbl_DN_NghiPhepNhanViens.Where(d => (d.isDauKy ?? false)).Select(d => d.nam).FirstOrDefault();
                if (namTon > 0)
                {
                    list = linqNS.sp_PhepNamNhanVien_Index(namTon, maPhongBan, searchString).ToList();
                }
                PagingLoaderController(string.Empty, list.Count(), currentPageIndex, string.Empty);
            }

            ViewData["lsDanhSach"] = list.Skip(start).Take(offset).ToList();

            return PartialView("ViewIndex_Tab3");
        }


        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            PhepNamModel = new PhepNamNhanVien();
            PhepNamModel.ngayLap = DateTime.Now;
            PhepNamModel.Nam = DateTime.Now.Year;
            PhepNamModel.NhanVien = new BatDongSan.Models.NhanSu.NhanVienModel()
            {
                maNhanVien = GetUser().manv,
                hoVaTen = HoVaTen(GetUser().manv)
            };
            var namTonDauTien = lqPhieuDN.tbl_DN_NghiPhepNhanViens.Where(d => (d.isDauKy ?? false)).Select(d => d.nam).FirstOrDefault();
            if (namTonDauTien == 0)
            {
                namTonDauTien = DateTime.Now.Year + 1;
            }

            ViewBag.namTon = namTonDauTien.ToString();

            //BuildNamLamViec(namTonDauTien, 1, namTonDauTien);

            return View(PhepNamModel);
        }

        public ActionResult DanhSachNhanVienChon(FormCollection coll, string[] MaNV, int Nam)
        {
            try
            {
                string[] maNhanVienDaAdd = coll.GetValues("maNhanVien");
                string[] splitStr = MaNV;

                ListNhanVienChon = new List<DanhSachNhanVien>();
                if (maNhanVienDaAdd != null && maNhanVienDaAdd.Length > 0)
                {
                    for (int i = 0; i < maNhanVienDaAdd.Length; i++)
                    {

                        var ListNhanVien = linqNS.sp_PhepNamCuaNV(maNhanVienDaAdd[i], Nam).FirstOrDefault();
                        DanhSachNhanVien DanhSach = new DanhSachNhanVien();
                        DanhSach.maPhongBan = ListNhanVien.DepartmentId;
                        DanhSach.tenPhongBan = ListNhanVien.TenPhongBan;

                        DanhSach.maChucVu = ListNhanVien.CapBacQuanLyId;
                        DanhSach.tenChucVu = ListNhanVien.CapBacQuanLy;

                        DanhSach.tenChiNhanh = ListNhanVien.MaBoPhan;

                        DanhSach.maNhanVien = ListNhanVien.maNV;
                        DanhSach.tenNhanVien = HoVaTen(ListNhanVien.maNV);
                        DanhSach.soNgayPhepTonNamTruoc = (decimal)ListNhanVien.PhepTon;
                        DanhSach.soNgayPhepTonDuocDuyet = (decimal)ListNhanVien.PhepTon;
                        DanhSach.soNgayPhepQuiDinh = ListNhanVien.NgayPhepQuiDinh;
                        DanhSach.soNgayPhepThamNien = ListNhanVien.NgayPhepThamNien;
                        DanhSach.tongSoNgayPhep = DanhSach.soNgayPhepTonDuocDuyet + DanhSach.soNgayPhepQuiDinh + DanhSach.soNgayPhepThamNien;
                        DanhSach.soNgayPhepDacBiet = 0;
                        ListNhanVienChon.Add(DanhSach);
                    }
                }

                if (splitStr != null && splitStr.Length > 0)
                {
                    for (int j = 0; j < splitStr.Length; j++)
                    {
                        if (!ListNhanVienChon.Select(d => d.maNhanVien).Contains(splitStr[j]))
                        {
                            var ListNhanVien = linqNS.sp_PhepNamCuaNV(splitStr[j], Nam).FirstOrDefault();
                            DanhSachNhanVien DanhSach = new DanhSachNhanVien();
                            DanhSach.maPhongBan = ListNhanVien.DepartmentId;
                            DanhSach.tenPhongBan = ListNhanVien.TenPhongBan;

                            DanhSach.tenChiNhanh = ListNhanVien.MaBoPhan;

                            DanhSach.maChucVu = ListNhanVien.CapBacQuanLyId;
                            DanhSach.tenChucVu = ListNhanVien.CapBacQuanLy;

                            DanhSach.maNhanVien = ListNhanVien.maNV;
                            DanhSach.tenNhanVien = HoVaTen(ListNhanVien.maNV);
                            DanhSach.soNgayPhepTonNamTruoc = (decimal?)ListNhanVien.PhepTon ?? 0;
                            DanhSach.soNgayPhepTonDuocDuyet = (decimal?)ListNhanVien.PhepTon ?? 0;
                            DanhSach.soNgayPhepQuiDinh = ListNhanVien.NgayPhepQuiDinh;
                            DanhSach.soNgayPhepThamNien = ListNhanVien.NgayPhepThamNien;
                            DanhSach.tongSoNgayPhep = DanhSach.soNgayPhepTonDuocDuyet + DanhSach.soNgayPhepQuiDinh + DanhSach.soNgayPhepThamNien;
                            //ListNhanVien.TongSoNgayPhepNam;
                            DanhSach.soNgayPhepDacBiet = 0;
                            ListNhanVienChon.Add(DanhSach);
                        }
                    }
                }
                return PartialView("_LoadDanhSachNhanVien", ListNhanVienChon);
            }
            catch
            {

                return View("error");
            }
        }
        [HttpPost]
        public ActionResult Create(FormCollection col)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            try
            {

                string[] maNhanVien = col.GetValues("maNhanVien");
                string[] _pepTon = col.GetValues("soNgayPhepTonDuocDuyet");
                int year = int.Parse(col.Get("Nam"));
                string employeeIdCurrent = "";
                int bienDem = 0;
                sp_PhepNamCuaNVResult info = new sp_PhepNamCuaNVResult();
                List<tbl_DN_NghiPhepNhanVien> checkInset = new List<tbl_DN_NghiPhepNhanVien>();
                List<tbl_DN_NghiPhepNhanVien> phieuDN = new List<tbl_DN_NghiPhepNhanVien>();
                if (maNhanVien.Count() > 0)
                {
                    for (int i = 0; i < maNhanVien.Count(); i++)
                    {
                        bienDem = i + 1;
                        employeeIdCurrent = maNhanVien[i];
                        info = linqNS.sp_PhepNamCuaNV(employeeIdCurrent, year).SingleOrDefault();
                        if (info != null)
                        {
                            tbl_DN_NghiPhepNhanVien linkQ = new tbl_DN_NghiPhepNhanVien();
                            linkQ.maNhanVien = maNhanVien[i];
                            linkQ.nam = year;

                            linkQ.soNgayPhepDacBiet = Convert.ToDecimal(col.GetValues("soNgayPhepDacBiet")[i]);
                            linkQ.maPhongBan = info.DepartmentId;
                            linkQ.capBacQuanLy = info.CapBacQuanLyId;
                            linkQ.soNgayPhepConLaiNamCu = (decimal)info.PhepTon;
                            linkQ.soNgayPhepTonDuocDuyet = decimal.Parse(_pepTon[i]);
                            linkQ.soNgayPhepQuiDinh = info.NgayPhepQuiDinh;
                            linkQ.soNgayPhepThamNien = info.NgayPhepThamNien;
                            linkQ.soNgayPhepNamHienTai = info.NgayPhepQuiDinh + info.NgayPhepThamNien + linkQ.soNgayPhepTonDuocDuyet + linkQ.soNgayPhepDacBiet;
                            linkQ.ngayLap = DateTime.Now;
                            linkQ.nguoiLap = GetUser().manv;
                            checkInset = lqPhieuDN.tbl_DN_NghiPhepNhanViens.Where(d => d.maNhanVien == linkQ.maNhanVien && d.nam == linkQ.nam).ToList();
                            if (checkInset.Count == 0)
                            {
                                phieuDN.Add(linkQ);
                                // lqPhieuDN.tbl_DN_NghiPhepNhanViens.InsertOnSubmit(linkQ);
                                // lqPhieuDN.SubmitChanges();
                            }
                        }
                    }
                    if (phieuDN.Count > 0)
                    {
                        lqPhieuDN.tbl_DN_NghiPhepNhanViens.InsertAllOnSubmit(phieuDN);
                        lqPhieuDN.SubmitChanges();
                    }
                }
                return RedirectToAction("Index");
            }
            catch
            {
                return View("error");
            }
        }

        public ActionResult Edit(string maNhanVien, int Nam)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            PhepNamModel = new PhepNamNhanVien();
            phepNam = lqPhieuDN.tbl_DN_NghiPhepNhanViens.Where(d => d.maNhanVien == maNhanVien && d.nam == Nam).FirstOrDefault();
            if (phepNam != null)
            {
                PhepNamModel.Nam = Nam;
                PhepNamModel.ngayLap = DateTime.Now;
                PhepNamModel.NhanVien = new BatDongSan.Models.NhanSu.NhanVienModel
                {
                    maNhanVien = GetUser().manv,
                    hoVaTen = HoVaTen(GetUser().manv)
                };
                PhepNamModel.tenChucVu = lqPhieuDN.GetTable<vw_NS_DanhSachNhanVienTheoPhongBan>().Where(d => d.maNhanVien == phepNam.maNhanVien).Select(d => d.TenChucDanh).FirstOrDefault() ?? string.Empty;
                //lqPhieuDN.GetTable<Sys_CapBacQuanLy>().Where(d => d.maCapBac == phepNam.capBacQuanLy).Select(d => d.tenCapBac).FirstOrDefault();

                PhepNamModel.soNgayPhepTonNamTruoc = phepNam.soNgayPhepConLaiNamCu;
                PhepNamModel.soNgayPhepTonDuocDuyet = phepNam.soNgayPhepTonDuocDuyet;
                PhepNamModel.soNgayPhepQuiDinh = phepNam.soNgayPhepQuiDinh;
                PhepNamModel.soNgayPhepThamNien = phepNam.soNgayPhepThamNien;
                PhepNamModel.tongSoNgayPhep = phepNam.soNgayPhepNamHienTai;
                PhepNamModel.maNhanVien = maNhanVien;
                PhepNamModel.tenNhanVien = HoVaTen(maNhanVien);
                PhepNamModel.soNgayPhepDacBiet = phepNam.soNgayPhepDacBiet ?? 0;
                PhepNamModel.Nam = Nam;
                return View(PhepNamModel);
            }
            else
            {
                return RedirectToAction("index");
            }

        }
        [HttpPost]
        public ActionResult Edit(FormCollection col)
        {
            try
            {
                #region Role user
                permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenSua);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion
                string maNhanVien = col.Get("MaNhanVien");
                int year = int.Parse(col.Get("name"));

                var PhepNam = lqPhieuDN.tbl_DN_NghiPhepNhanViens.Where(d => d.maNhanVien == maNhanVien && d.nam == year).FirstOrDefault();
                PhepNam.soNgayPhepTonDuocDuyet = decimal.Parse(col.Get("soNgayPhepTonDuocDuyet"));
                PhepNam.soNgayPhepQuiDinh = decimal.Parse(col.Get("soNgayPhepNamQuiDinh"));
                PhepNam.soNgayPhepThamNien = decimal.Parse(col.Get("soNgayPhepThamNien"));
                PhepNam.soNgayPhepNamHienTai = decimal.Parse(col.Get("soNgayPhepHienTai"));
                PhepNam.soNgayPhepDacBiet = Convert.ToDecimal(col.Get("soNgayPhepDacBiet"));
                PhepNam.nguoiLap = GetUser().manv;
                PhepNam.ngayLap = DateTime.Now;
                lqPhieuDN.SubmitChanges();
                return RedirectToAction("Edit", new { maNhanVien = maNhanVien, Nam = PhepNam.nam });
            }
            catch
            {
                return View("error");
            }
        }
        public ActionResult ChonNhanVien()
        {
            var buildTree = new StringBuilder();
            var phongBans = linqDM.tbl_DM_PhongBans.ToList();
            buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
            ViewBag.NVPB = buildTree.ToString();
            return PartialView("_ChonNhanVien");
        }
        public ActionResult LoadNhanVien(int? page, string searchString, string maPhongBan, int Nam)
        {
            var NhanVienDaLapPhep = lqPhieuDN.tbl_DN_NghiPhepNhanViens.Where(d => d.nam == Nam).Select(v => v.maNhanVien.Trim()).ToList();

            IList<sp_PB_DanhSachNhanVienResult> phongBan1s;
            phongBan1s = linqDM.sp_PB_DanhSachNhanVien(searchString, maPhongBan).Where(d => !NhanVienDaLapPhep.Contains(d.maNhanVien.Trim())).ToList();
            ViewBag.isGet = "True";
            int currentPageIndex = page.HasValue ? page.Value : 1;
            ViewBag.Count = phongBan1s.Count();
            ViewBag.Search = searchString;
            ViewBag.MaPhongBan = maPhongBan;
            ViewBag.Nam = Nam;
            return PartialView("_LoadChonNhanVien", phongBan1s.ToPagedList(currentPageIndex, 10));
        }


        public void BuildNamLamViec(int namHienTai, int limmit, int selected)
        {
            Dictionary<int, int> dics = new Dictionary<int, int>();
            for (int i = namHienTai - limmit; i <= namHienTai + limmit; i++)
            {
                dics.Add(i, i);
            }
            ViewBag.namLamViec = new SelectList(dics, "Key", "Value", selected);
        }


        public void BuildNamTonDauKy(int namTonDauTien, int limmit, int? selected)
        {
            Dictionary<int, int> dics = new Dictionary<int, int>();

            for (int i = namTonDauTien; i <= namTonDauTien + limmit; i++)
            {
                dics.Add(i, i);
            }

            ViewBag.namTonDauKy = new SelectList(dics, "Key", "Value", selected);
        }

        public string HoVaTen(string MaNV)
        {

            return linqNS.tbl_NS_NhanViens.Where(d => d.maNhanVien == MaNV).Select(d => d.ho + " " + d.ten).FirstOrDefault();
        }

        //public ActionResult ResetPhepTon(int? nam)
        //{
        //    try
        //    {
        //        linqNS.sp_ReSetPhepTon(nam);
        //        return Json(string.Empty);
        //    }
        //    catch
        //    {
        //        return View();
        //    }
        //}

        public ActionResult ResetPhepTon(string ngayReset)
        {
            try
            {
                if (string.IsNullOrEmpty(ngayReset))
                {
                    return Json("Vui lòng chọn ngày cần Reset.!");
                }

                linqNS.sp_ReSetPhepTon(DateTime.ParseExact(ngayReset, "dd/MM/yyyy", CultureInfo.InvariantCulture), GetUser().manv);
                return Json(string.Empty);
            }
            catch
            {
                return View();
            }
        }


        public ActionResult CapNhatPhepTon(int nam)
        {
            try
            {
                #region Role user
                permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenSua);
                if (!permission.HasValue)
                    return Json("LogIn");
                if (!permission.Value)
                    return Json("AccessDenied");
                #endregion

                //linqNS.sp_NS_KetChuyenTonPhep(nam, GetUser().manv);
                var dauKy = linqNS.GetTable<tbl_DN_NghiPhepNhanVien>().Where(d => d.nam == nam).FirstOrDefault();
                if (dauKy == null)
                {
                    linqNS.sp_PhepNamNhanVien_TuDongInsert_KhiBatDauQuaNamMoi(nam, 1);
                }
                else
                {
                    if ((dauKy.isDauKy ?? false) == true)
                    {
                        return Json("-1");
                    }
                    else
                    {
                        linqNS.sp_PhepNamNhanVien_TuDongInsert_KhiBatDauQuaNamMoi(nam, 1);
                    }
                }
                return Json(string.Empty);
            }
            catch
            {
                return Json("Có lỗi xảy ra trong quá trình xử lý, Vui lòng liên hệ bộ phận IT.");
            }
        }
        

        public void XuatFilePhepNam(string ngayKiemTra, string qSearch, string maPhongBan)
        {

            try
            {
                //#region Role user
                //permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenIn);
                //if (!permission.HasValue)
                //    return;
                //if (!permission.Value)
                //    return;
                //#endregion

                DateTime ngayCanTinh = DateTime.ParseExact(ngayKiemTra, "dd/MM/yyyy", CultureInfo.InvariantCulture);

                var filename = "";
                var virtualPath = HttpRuntime.AppDomainAppVirtualPath;

                var fileStream = new FileStream(System.Web.HttpContext.Current.Server.MapPath(virtualPath + @"\Content\Report\ReportTemplate.xls"), FileMode.Open, FileAccess.Read);

                var workbook = new HSSFWorkbook(fileStream, true);
                filename += "PhepNamNhanVien_" + ngayKiemTra + ".xls";


                var sheet = workbook.GetSheet("danhsachnhanvien");


                #region format style excel
                /*style title start*/
                //tạo font cho các title
                //font tiêu đề 
                HSSFFont hFontTieuDe = (HSSFFont)workbook.CreateFont();
                hFontTieuDe.FontHeightInPoints = 18;
                hFontTieuDe.Boldweight = 100 * 10;
                hFontTieuDe.FontName = "Times New Roman";
                hFontTieuDe.Color = HSSFColor.BLUE.index;

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

                Row rowC = null;
                //Khai báo row đầu tiên
                int firstRowNumber = 1;

                string cellTenCty = "CÔNG TY CỔ PHẦN TV.WINDOW";
                var titleCellCty = HSSFCellUtil.CreateCell(sheet.CreateRow(0), 0, cellTenCty.ToUpper());
                titleCellCty.CellStyle = styleTitle;

                //string cellEnd1 = "Bảng Lương Tháng " + thang + " năm " + nam;
                //var titleCellEnd1 = HSSFCellUtil.CreateCell(sheet.GetRow(0), 43, cellEnd1.ToUpper());
                //titleCellEnd1.CellStyle = styleTitle;

                string rowtitle = "Bảng tính nghĩ phép " + " thống kê đến ngày " + ngayKiemTra;
                var titleCell = HSSFCellUtil.CreateCell(sheet.CreateRow(firstRowNumber), 0, rowtitle.ToUpper());
                titleCell.CellStyle = styleTitle;

                //string cellEnd2 = "";
                //var titleCellEnd2 = HSSFCellUtil.CreateCell(sheet.GetRow(1), 43, cellEnd2.ToUpper());
                //titleCellEnd2.CellStyle = styleTitle;

                ++firstRowNumber;

                var list1 = new List<string>();
                list1.Add("STT");
                list1.Add("Mã nhân viên");
                list1.Add("Họ tên");
                list1.Add("Phòng ban");
                list1.Add("Phép tồn " + (ngayCanTinh.Year - 1));
                list1.Add("Phép năm " + ngayCanTinh.Year + " theo qui định");
                list1.Add("Tổng phép đầu năm " + ngayCanTinh.Year);
                list1.Add("Phép năm " + ngayCanTinh.Year + " đã nghĩ đến thời điểm " + ngayKiemTra);
                list1.Add("Phép nghỉ " + (ngayCanTinh.Year - 1));
                list1.Add("Phép nghỉ " + (ngayCanTinh.Year));
                list1.Add("Số phép tồn đến ngày " + ngayKiemTra);
                //list1.Add("Số ngày phép được tính lại năm " + (ngayCanTinh.Year));
                //list1.Add("Số ngày phép năm còn lại");

                var idRowStart = 2; // bat dau o dong thu 4
                var headerRow = sheet.CreateRow(idRowStart);
                int rowend = idRowStart;
                ReportHelperExcel.CreateHeaderRow(headerRow, 0, styleheadedColumnTable, list1);

                var data = linqNS.sp_PhepTonNam(ngayCanTinh, qSearch, maPhongBan).ToList();

                var stt = 0;
                int dem = 0;

                foreach (var item in data)
                {
                    dem = 0;

                    idRowStart++;
                    rowC = sheet.CreateRow(idRowStart);
                    ReportHelperExcel.SetAlignment(rowC, dem++, (++stt).ToString(), hStyleConCenter);

                    ReportHelperExcel.SetAlignment(rowC, dem++, item.maNhanVien, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item.tenNhanVien, hStyleConLeft);

                    ReportHelperExcel.SetAlignment(rowC, dem++, item.phongBan, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item.tonNhapCu), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item.quiDinhNam_Moi), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item.tongNam_Moi), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item.soPhepDaNghiNam_Moi), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item.tonNam_Cu), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item.thongKenghi_namMoi), hStyleConRight);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item.soPhepTon), hStyleConRight);
                    //ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.##}", item.soNgayPhepNamHienTai), hStyleConRight);
                }

                sheet.SetColumnWidth(0, 8 * 210);
                sheet.SetColumnWidth(1, 30 * 210);
                sheet.SetColumnWidth(2, 30 * 310);
                sheet.SetColumnWidth(3, 30 * 310);
                sheet.SetColumnWidth(4, 20 * 210);
                sheet.SetColumnWidth(5, 20 * 210);
                sheet.SetColumnWidth(6, 20 * 210);
                sheet.SetColumnWidth(7, 20 * 210);
                sheet.SetColumnWidth(8, 20 * 210);
                sheet.SetColumnWidth(9, 20 * 210);
                sheet.SetColumnWidth(10, 20 * 210);
                //sheet.SetColumnWidth(11, 20 * 210);
                //sheet.SetColumnWidth(10, 8 * 210);

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

        public ActionResult UpdateNgayPhepConLaiNhanVien(string maNhanVien, int nam, decimal soNgayPhepNamHienTai)
        {
            try
            {
                #region Role user
                permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenSua);
                if (!permission.HasValue)
                    return Json("LogIn");
                if (!permission.Value)
                    return Json("AccessDenied");
                #endregion
                var objPhepConLaiNhanVien = lqPhieuDN.tbl_DN_NghiPhepNhanViens.Where(d => d.maNhanVien == maNhanVien && d.nam == nam).FirstOrDefault();
                if (objPhepConLaiNhanVien != null)
                {
                    //Ngày phép bổ sung của nhân viên đó
                    objPhepConLaiNhanVien.soNgayPhepBoSung = (float)soNgayPhepNamHienTai;
                }

                lqPhieuDN.SubmitChanges();
                SaveActiveHistory("Cập nhật tổng số ngày phép nhân viên: " + maNhanVien + " số ngày: " + soNgayPhepNamHienTai + " năm: " + nam);
                return Json(string.Empty);
            }
            catch
            {
                return Json("ERROR");
            }
        }
    }
}
