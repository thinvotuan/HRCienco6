using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Helper.Utils;
using BatDongSan.Models.DanhMuc;
using BatDongSan.Models.HeThong;
using BatDongSan.Models.NhanSu;
using BatDongSan.Utils.Paging;

namespace BatDongSan.Controllers.NhanSu
{
    public class LuongChoViecController : ApplicationController
    {
        LinqDanhMucDataContext linqDanhMuc = new LinqDanhMucDataContext();
        LinqHeThongDataContext lqHT = new LinqHeThongDataContext();
        LinqNhanSuDataContext context = new LinqNhanSuDataContext();
        IList<sp_NS_LuongChoViec_IndexResult> luongChoViecs;
        HopDongLaoDongModel model;
        tbl_NS_LuongChoViec choViec;
        LinqDanhMucDataContext linqDM = new LinqDanhMucDataContext();
        IList<tbl_DM_PhongBan> phongBans;
        StringBuilder buildTree = null;
        readonly string MCV = "LuongChoViec";
        bool? permission;

        public ActionResult Index(int? page, int? pageSize, string loaiHopDong, string searchString, string maPhongBan)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            BindDataTrangThai(MCV);
            buildTree = new StringBuilder();
            phongBans = linqDM.tbl_DM_PhongBans.ToList();
            buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
            ViewBag.PhongBans = buildTree.ToString();
            //
            int currentPageIndex = page.HasValue ? page.Value : 1;
            pageSize = pageSize ?? 30;

            var loaiHopDongs = context.tbl_DM_LoaiHopDongLaoDongs.ToList();
            loaiHopDongs.Insert(0, new BatDongSan.Models.NhanSu.tbl_DM_LoaiHopDongLaoDong { maLoaiHopDong = "", tenLoaiHopDong = "[Tất cả loại hợp đồng]" });
            ViewBag.LoaiHopDongs = new SelectList(loaiHopDongs, "maLoaiHopDong", "tenLoaiHopDong");
            luongChoViecs = context.sp_NS_LuongChoViec_Index(null, loaiHopDong, maPhongBan, searchString, currentPageIndex, 20).OrderByDescending(d => d.maLuongChoViec).ToList();

            int? tongSoDong = 0;
            try
            {
                ViewBag.Count = luongChoViecs[0].tongSoDong;
                tongSoDong = luongChoViecs[0].tongSoDong;
            }
            catch
            {
                ViewBag.Count = 0;
            }
            if (Request.IsAjaxRequest())
            {
                return PartialView("PartialIndex", luongChoViecs.ToPagedList(currentPageIndex, 20, true, tongSoDong));
            }
            ViewBag.searchString = searchString;
            ViewBag.loaiHopDong = loaiHopDong;
            ViewBag.maPhongBan = maPhongBan;
            return View(luongChoViecs.ToPagedList(currentPageIndex, 20, true, tongSoDong));
        }

        #region Create, Edit, Details

        // GET: /HopDongLaoDong/Create

        public ActionResult Create()
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
                thang(DateTime.Now.Month);
                nam(DateTime.Now.Year);
                model = new HopDongLaoDongModel();
                model.maLuongChoViec = IdGenerator();
                model.ngayLap = DateTime.Now;
                model.tyLeLuongChoViec = 0;
                model.congChoViec = 0;
                model.luongChoViec = 0;
                return View(model);
            }
            catch
            {
                return View("Error");
            }
        }

        // GET: /HopDongLaoDong/Create

        public ActionResult Edit(string id)
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
                choViec = context.tbl_NS_LuongChoViecs.Where(d => d.maLuongChoViec == id).FirstOrDefault();
                SetModelData(choViec.soHopDong);
                SetDataModelChoViec(choViec);
                thang((int?) choViec.thang??DateTime.Now.Month);
                nam((int?) choViec.nam??DateTime.Now.Year);
                return View(model);
            }
            catch
            {
                return View("Error");
            }
        }

        public ActionResult Details(string id)
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
                choViec = context.tbl_NS_LuongChoViecs.Where(d => d.maLuongChoViec == id).FirstOrDefault();
                SetModelData(choViec.soHopDong);
                SetDataModelChoViec(choViec);
                return View(model);
            }
            catch
            {
                return View("Error");
            }
        }
        //
        // POST: /HopDongLaoDong/Create

        [HttpPost]
        public ActionResult Create(FormCollection collection)
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
                choViec = new tbl_NS_LuongChoViec();
                GetDataFromView(collection, true);
                context.tbl_NS_LuongChoViecs.InsertOnSubmit(choViec);
                context.SubmitChanges();
                return RedirectToAction("Edit", new { id = choViec.maLuongChoViec });
            }
            catch
            {
                return View("Error");
            }
        }

        [HttpPost]
        public ActionResult Edit(FormCollection collection)
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
                choViec = new tbl_NS_LuongChoViec();
                GetDataFromView(collection, false);
                context.SubmitChanges();
                return RedirectToAction("Edit", new { id = choViec.maLuongChoViec });
            }
            catch
            {
                return View("Error");
            }
        }

        public void GetDataFromView(FormCollection collection, bool isCreate)
        {
            if (isCreate == true)
            {
                choViec.maLuongChoViec = IdGenerator();
                choViec.maNhanVien = collection.Get("maNhanVien");
            }
            else
            {
                choViec = context.tbl_NS_LuongChoViecs.Where(d => d.maLuongChoViec == collection["maLuongChoViec"]).FirstOrDefault();
                choViec.maLuongChoViec = collection.Get("maLuongChoViec");
            }

            choViec.ngayLap = DateTime.Now;
            choViec.nguoiLap = GetUser().manv;
            choViec.luongDongBaoHiem = string.IsNullOrEmpty(collection["luongDongBH"]) ? 0 : Convert.ToDecimal(collection["luongDongBH"]);
            choViec.khoanBoSungLuong = string.IsNullOrEmpty(collection["khoanBoSungLuong"]) ? 0 : Convert.ToDecimal(collection["khoanBoSungLuong"]);
            choViec.tongLuong = string.IsNullOrEmpty(collection["tongLuong"]) ? 0 : Convert.ToDecimal(collection["tongLuong"]);
            choViec.soHopDong = collection["soHopDong"];
            choViec.soPhuLuc = collection["soPhuLuc"];
            choViec.tyLe = Convert.ToDouble(collection["tyLeLuongChoViec"].Replace('%', ' '));
            choViec.luong = string.IsNullOrEmpty(collection["luongThoaThuan"]) ? 0 : Convert.ToDecimal(collection["luongThoaThuan"]);
            choViec.phuCapLuong = string.IsNullOrEmpty(collection["phuCapLuong"]) ? 0 : Convert.ToDecimal(collection["phuCapLuong"]);
            choViec.luongChoViec = string.IsNullOrEmpty(collection["luongChoViec"]) ? 0 : Convert.ToDecimal(collection["luongChoViec"]);
            choViec.congChoViec = string.IsNullOrEmpty(collection["congChoViec"]) ? 0 : Convert.ToInt32(collection["congChoViec"]);
            choViec.thang = string.IsNullOrEmpty(collection["thang"]) ? 0 : Convert.ToInt32(collection["thang"]);
            choViec.nam = string.IsNullOrEmpty(collection["nam"]) ? 0 : Convert.ToInt32(collection["nam"]);
            choViec.ghiChu = collection.Get("ghiChuMain");
        }


        //
        // POST: /HopDongLaoDong/Delete/5

        [HttpPost]
        public ActionResult Delete(string id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXoa);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            try
            {
                choViec = context.tbl_NS_LuongChoViecs.Where(s => s.maLuongChoViec == id).FirstOrDefault();
                context.tbl_NS_LuongChoViecs.DeleteOnSubmit(choViec);
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View("Error");
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
        }

        public void SetDataModelChoViec(tbl_NS_LuongChoViec choViec)
        {
            model.maLuongChoViec = choViec.maLuongChoViec;
            model.ngayLap = choViec.ngayLap;
            model.ghiChu = choViec.ghiChu;
            model.tyLeLuongChoViec = (double)choViec.tyLe;
            model.congChoViec = (double)choViec.congChoViec;
            model.luongChoViec = (decimal?)choViec.luongChoViec ?? 0;
            model.tenHopDong = context.tbl_NS_HopDongLaoDongs.Where(d => d.soHopDong == choViec.soHopDong).Select(d => d.tenHopDong).FirstOrDefault() ?? string.Empty;
            model.soPhuLuc = choViec.soPhuLuc;
            model.NguoiLap = new NhanVienModel(choViec.nguoiLap, HoVaTen(choViec.nguoiLap));
            model.NhanVien = new NhanVienModel(choViec.maNhanVien, HoVaTen(choViec.maNhanVien));
            model.luongThoaThuan = choViec.luong;
        }

        #endregion

        public string IdGenerator()
        {
            return GenerateUtil.CheckLetter("LCV", GetMax());
        }

        public string GetMax()
        {
            return context.tbl_NS_LuongChoViecs.OrderByDescending(d=>d.ngayLap).Select(d => d.maLuongChoViec).FirstOrDefault()??String.Empty;
        }

        #region  Load danh sách nhân viên

        public ActionResult NhanVienDuyet()
        {
            try
            {
                buildTree = new StringBuilder();
                phongBans = linqDanhMuc.tbl_DM_PhongBans.ToList();
                buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
                ViewBag.NVPB = buildTree.ToString();
                return PartialView("_NhanVienPhongBan");
            }
            catch
            {
                return View("Error");
            }
        }

        public ActionResult LoadNhanVienDuyet(int? page, string searchString, string maPhongBan)
        {
            try
            {
                IList<sp_PB_DanhSachNhanVienResult> phongBan1s;
                phongBan1s = linqDanhMuc.sp_PB_DanhSachNhanVien(searchString, maPhongBan).ToList();
                ViewBag.isGet = "True";
                int currentPageIndex = page.HasValue ? page.Value : 1;
                ViewBag.Count = phongBan1s.Count();
                ViewBag.Search = searchString;
                ViewBag.MaPhongBan = maPhongBan;
                return PartialView("LoadNhanVienDuyet", phongBan1s.ToPagedList(currentPageIndex, 10));
            }
            catch
            {
                return View("Error");
            }
        }
        #endregion

        #region Thông tin  thỏa thuận mức lương khi ký hợp đồng của nhân viên đó


        public JsonResult ThongTinNhanVien(string maNhanVien)
        {
            try
            {
                model = new HopDongLaoDongModel();
                string maHopDong = context.tbl_NS_HopDongLaoDongs.Where(d => d.maNhanVien == maNhanVien).Select(d => d.soHopDong).FirstOrDefault() ?? string.Empty;
                model = SetModelData(maHopDong);
                return Json(model);
            }
            catch
            {
                return Json("False");
            }
        }

        public HopDongLaoDongModel SetModelData(string maHopDong)
        {
            //Kiểm tra hợp đồng đó có phụ lục hay chưa nếu có thì lấy phụ lục mới nhất và phụ lục đó đã duyệt
            DMNguoiDuyetController nd = new DMNguoiDuyetController();
            string soPhuLuc = string.Empty;
            var dsPhuLuc = context.tbl_NS_PhuLucHopDongs.Where(d => d.soHopDong == maHopDong).Select(d => new PhuLucHopDongModel
            {
                soPhuLuc = d.soPhuLuc,
                Duyet = nd.GetDetailByMaPhieuTheoQuiTrinhDong(d.soPhuLuc, d.maQuiTrinhDuyet ?? 0),
            }).ToList();
            if (dsPhuLuc != null && dsPhuLuc.Count() > 0)
            {
                soPhuLuc = dsPhuLuc.OrderByDescending(d => d.soPhuLuc).Where(d => d.Duyet.maBuocDuyet == "DTN" || d.Duyet.maBuocDuyet == "DD").Select(d => d.soPhuLuc).FirstOrDefault() ?? string.Empty;
            }
            if (soPhuLuc != "")
            {
                model = (from hd in context.tbl_NS_PhuLucHopDongs
                         where hd.soPhuLuc == soPhuLuc
                         select new HopDongLaoDongModel
                         {
                             soHopDong = hd.soHopDong,
                             tenHopDong = context.tbl_NS_HopDongLaoDongs.Where(d => d.soHopDong == maHopDong).Select(d => d.tenHopDong).FirstOrDefault() ?? string.Empty,
                             soPhuLuc = soPhuLuc,
                             luongDongBH = hd.luongDongBaoHiem ?? 0,
                             khoanBoSungLuong = hd.khoanBoSungLuong ?? 0,
                             phuCapLuong = (double?)hd.phuCapLuong ?? 0,
                             tongLuong = hd.tongLuong ?? 0,
                             luongHopDong = hd.luong,
                         }).FirstOrDefault();
            }
            else
            {
                model = (from hd in context.tbl_NS_HopDongLaoDongs
                         where hd.soHopDong == maHopDong
                         select new HopDongLaoDongModel
                         {
                             soHopDong = hd.soHopDong,
                             tenHopDong = hd.tenHopDong,
                             soPhuLuc = string.Empty,
                             luongDongBH = hd.luongDongBaoHiem ?? 0,
                             khoanBoSungLuong = hd.khoanBoSungLuong ?? 0,
                             phuCapLuong = (double?)hd.phuCapLuong ?? 0,
                             tongLuong = hd.tongLuong ?? 0,
                             luongHopDong = hd.luongThoaThuan,
                         }).FirstOrDefault();
            }
            return model;
        }
        #endregion
    }
}
