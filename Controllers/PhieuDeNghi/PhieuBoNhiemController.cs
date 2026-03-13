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

namespace BatDongSan.Controllers.PhieuDeNghi
{
    public class PhieuBoNhiemController : ApplicationController
    {
        LinqHeThongDataContext lqHT = new LinqHeThongDataContext();
        LinqPhieuDeNghiDataContext lqPhieuDN = new LinqPhieuDeNghiDataContext();
        LinqNhanSuDataContext linqNS = new LinqNhanSuDataContext();
        LinqDanhMucDataContext linqDM = new LinqDanhMucDataContext();
        tbl_NS_PhieuDieuChuyen tblPhieuDeNghi;
        IList<tbl_NS_PhieuDieuChuyen> tblPhieuDeNghis;
        PhieuDieuChuyen PhieuDeNghiModel;
        public const string taskIDSystem = "PhieuBoNhiem";//REGWORKVOTE
        public bool? permission;
        //
        // GET: /PhieuCongTac/

        public ActionResult Index(int? page, string searchString, string tuNgay, string denNgay, string trangThai)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            BindDataTrangThai(taskIDSystem);
            var userName = GetUser().manv;
            DateTime? fromDate = null;
            DateTime? toDate = null;
            if (!String.IsNullOrEmpty(tuNgay))
                fromDate = DateTime.ParseExact(tuNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            if (!String.IsNullOrEmpty(denNgay))
                toDate = DateTime.ParseExact(denNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);

            ViewBag.isGet = "True";
            var tblPhieuDeNghis = lqPhieuDN.sp_NS_PhieuDangKyDieuChuyen_Index(trangThai, fromDate, toDate, searchString, GetUser().manv, "PhieuBoNhiem").ToList();
            int currentPageIndex = page.HasValue ? page.Value : 1;
            ViewBag.Count = tblPhieuDeNghis.Count();
            ViewBag.Search = searchString;
            ViewBag.tuNgay = tuNgay;
            ViewBag.denNgay = tuNgay;
            ViewBag.trangThai = trangThai;
            return View(tblPhieuDeNghis.ToPagedList(currentPageIndex, 50));

        }
        public ActionResult ViewIndex(int? page, string qSearch, string tuNgay, string denNgay, string trangThai)
        {
            var userName = GetUser().manv;
            DateTime? fromDate = null;
            DateTime? toDate = null;
            if (!String.IsNullOrEmpty(tuNgay))
                fromDate = String.IsNullOrEmpty(tuNgay) ? (DateTime?)null : DateTime.ParseExact(tuNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            if (!String.IsNullOrEmpty(denNgay))
            {
                toDate = String.IsNullOrEmpty(denNgay) ? (DateTime?)null : DateTime.ParseExact(denNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                toDate = toDate.Value.AddDays(1);
            }

            ViewBag.isGet = "True";
            var tblPhieuDeNghis = lqPhieuDN.sp_NS_PhieuDangKyDieuChuyen_Index(trangThai, fromDate, toDate, qSearch, GetUser().manv, "PhieuBoNhiem").ToList();
            int currentPageIndex = page.HasValue ? page.Value : 1;
            ViewBag.Count = tblPhieuDeNghis.Count();
            ViewBag.Search = qSearch;
            ViewBag.tuNgay = tuNgay;
            ViewBag.denNgay = denNgay;
            ViewBag.trangThai = trangThai;
            return PartialView("ViewIndex", tblPhieuDeNghis.ToPagedList(currentPageIndex, 50));

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

            PhieuDeNghiModel = new PhieuDieuChuyen();
            PhieuDeNghiModel.maPhieu = IdGenerator();// GenerateUtil.CheckLetter("PBN", GetMax());
            PhieuDeNghiModel.ngayLap = DateTime.Now;
            PhieuDeNghiModel.loaiDieuChuyen = "PhieuBoNhiem";
            PhieuDeNghiModel.NhanVienLapPhieu = new BatDongSan.Models.NhanSu.NhanVienModel
            {
                maNhanVien = GetUser().manv,
                hoVaTen = HoVaTen(GetUser().manv)
            };
            var listPhanCas = linqDM.tbl_NS_PhanCas;
            ViewBag.ListPhanCas = new SelectList(listPhanCas, "maPhanCa", "tenPhanCa");
            buitlChucDanh(string.Empty);
            buitlBoPhanTinhLuong(string.Empty);
            //buitlCoSo(string.Empty);
            return View(PhieuDeNghiModel);
        }

        [HttpPost]
        [ValidateInput(false)]
        public ActionResult Create(FormCollection coll)
        {
            try
            {
                tblPhieuDeNghi = new tbl_NS_PhieuDieuChuyen();
                tblPhieuDeNghi.soPhieu = IdGenerator();// GenerateUtil.CheckLetter("PBN", GetMax());
                tblPhieuDeNghi.nguoiLap = GetUser().manv;
                tblPhieuDeNghi.ngayLap = DateTime.Now;

                tblPhieuDeNghi.maNhanVien = coll.Get("NhanVien.maNhanVien");
                tblPhieuDeNghi.ngayChuyen = String.IsNullOrEmpty(coll.Get("ngayChuyen")) ? (DateTime?)null : DateTime.ParseExact(coll.Get("ngayChuyen"), "dd/MM/yyyy", CultureInfo.InvariantCulture);

                tblPhieuDeNghi.maPhongBanMoi = coll.Get("maPhongBanCu");
                tblPhieuDeNghi.maChucDanhMoi = coll.Get("maChucDanh");

                tblPhieuDeNghi.boPhanTinhLuongMoi = coll.Get("maBoPhanTinhLuong");
                tblPhieuDeNghi.maPhanCaMoi =!String.IsNullOrEmpty(coll.Get("maPhanCaMoi")) ? coll.Get("maPhanCaMoi"):string.Empty;
                tblPhieuDeNghi.maPhanCaCu = !String.IsNullOrEmpty(coll.Get("maPhanCaCu")) ? coll.Get("maPhanCaCu") : string.Empty; 
                tblPhieuDeNghi.maPhongBanCu = coll.Get("maPhongBanCu");
                tblPhieuDeNghi.maChucDanhCu = coll.Get("maChucDanhCu");
                tblPhieuDeNghi.boPhanTinhLuongCu = coll.Get("maBoPhanTinhLuongCu");

                tblPhieuDeNghi.noiDung = coll.Get("noiDung");
                tblPhieuDeNghi.loaiDieuChuyen = "PhieuBoNhiem";
                lqPhieuDN.tbl_NS_PhieuDieuChuyens.InsertOnSubmit(tblPhieuDeNghi);
                lqPhieuDN.SubmitChanges();

                return RedirectToAction("Edit", new { id = tblPhieuDeNghi.soPhieu });
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }
        }

        [HttpPost]
        public ActionResult Delete(string id)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXoa);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            try
            {
                var phieu = lqPhieuDN.tbl_NS_PhieuDieuChuyens.Where(d => d.soPhieu == id).FirstOrDefault();
                lqPhieuDN.tbl_NS_PhieuDieuChuyens.DeleteOnSubmit(phieu);
                lqPhieuDN.SubmitChanges();

                LinqHeThongDataContext lqHT = new LinqHeThongDataContext();
                var nguoiDuyet = lqHT.tbl_HT_DMNguoiDuyets.Where(d => d.maPhieu == id).ToList();
                if (nguoiDuyet != null && nguoiDuyet.Count > 0)
                {
                    lqHT.tbl_HT_DMNguoiDuyets.DeleteAllOnSubmit(nguoiDuyet);
                    lqHT.SubmitChanges();
                }

                return RedirectToAction("Index");
            }
            catch
            {
                return View("error");
            }
        }
        public ActionResult Edit(string id)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            var dataPhieuCongTac = lqPhieuDN.tbl_NS_PhieuDieuChuyens.Where(d => d.soPhieu == id).FirstOrDefault();
            PhieuDeNghiModel = new PhieuDieuChuyen();
            PhieuDeNghiModel.maPhieu = dataPhieuCongTac.soPhieu;
            PhieuDeNghiModel.ngayLap = dataPhieuCongTac.ngayLap;
            PhieuDeNghiModel.maNhanVienLapPhieu = dataPhieuCongTac.nguoiLap;
            PhieuDeNghiModel.NhanVien = new BatDongSan.Models.NhanSu.NhanVienModel
            {
                maNhanVien = dataPhieuCongTac.maNhanVien,
                hoVaTen = HoVaTen(dataPhieuCongTac.maNhanVien)
            };
            PhieuDeNghiModel.NhanVienLapPhieu = new BatDongSan.Models.NhanSu.NhanVienModel
            {
                maNhanVien = dataPhieuCongTac.nguoiLap,
                hoVaTen = HoVaTen(dataPhieuCongTac.nguoiLap)
            };
            PhieuDeNghiModel.ngayChuyen = dataPhieuCongTac.ngayChuyen;


            PhieuDeNghiModel.maPhongBan = dataPhieuCongTac.maPhongBanMoi;
            PhieuDeNghiModel.tenPhongBan = linqDM.tbl_DM_PhongBans.Where(d => d.maPhongBan == dataPhieuCongTac.maPhongBanMoi).Select(d => d.tenPhongBan).FirstOrDefault();
            PhieuDeNghiModel.maChucDanhMoi = dataPhieuCongTac.maChucDanhMoi;
            PhieuDeNghiModel.maBoPhanTinhLuongMoi = dataPhieuCongTac.boPhanTinhLuongMoi;
            PhieuDeNghiModel.maPhanCaMoi = short.Parse(!string.IsNullOrEmpty(dataPhieuCongTac.maPhanCaMoi) ? dataPhieuCongTac.maPhanCaMoi : "0");
            PhieuDeNghiModel.tenPhanCaMoi = linqDM.tbl_NS_PhanCas.Where(d => d.maPhanCa == short.Parse(!string.IsNullOrEmpty(dataPhieuCongTac.maPhanCaMoi) ? dataPhieuCongTac.maPhanCaMoi : "0")).Select(d => d.tenPhanCa).FirstOrDefault();
            PhieuDeNghiModel.tenPhanCaCu = linqDM.tbl_NS_PhanCas.Where(d => d.maPhanCa == short.Parse(!string.IsNullOrEmpty(dataPhieuCongTac.maPhanCaCu) ? dataPhieuCongTac.maPhanCaCu : "0")).Select(d => d.tenPhanCa).FirstOrDefault();
            PhieuDeNghiModel.maPhanCaCu = short.Parse(!string.IsNullOrEmpty(dataPhieuCongTac.maPhanCaCu) ? dataPhieuCongTac.maPhanCaCu : "0");
            PhieuDeNghiModel.maPhongBanCu = dataPhieuCongTac.maPhongBanCu;
            PhieuDeNghiModel.tenPhongBanCu = linqDM.tbl_DM_PhongBans.Where(d => d.maPhongBan == dataPhieuCongTac.maPhongBanCu).Select(d => d.tenPhongBan).FirstOrDefault();

            PhieuDeNghiModel.maChucDanhCu = dataPhieuCongTac.maChucDanhCu;
            PhieuDeNghiModel.tenChucDanhCu = linqNS.Sys_ChucDanhs.Where(d => d.MaChucDanh == dataPhieuCongTac.maChucDanhCu).Select(d => d.TenChucDanh).FirstOrDefault();

            PhieuDeNghiModel.maBoPhanTinhLuongCu = dataPhieuCongTac.boPhanTinhLuongCu;
            PhieuDeNghiModel.tenBoPhanTinhLuongCu = linqDM.tbl_NS_KhoiTinhLuongs.Where(d => d.maKhoiTinhLuong == Convert.ToInt32(dataPhieuCongTac.boPhanTinhLuongCu)).Select(d => d.tenKhoiTinhLuong).FirstOrDefault();
            PhieuDeNghiModel.loaiDieuChuyen = "PhieuBoNhiem";



            PhieuDeNghiModel.maQuiTrinhDuyet = dataPhieuCongTac.maQuiTrinhDuyet ?? 0;

            PhieuDeNghiModel.noiDung = dataPhieuCongTac.noiDung;

            DMNguoiDuyetController nd = new DMNguoiDuyetController();
            PhieuDeNghiModel.Duyet = nd.GetDetailByMaPhieuTheoQuiTrinhDong(PhieuDeNghiModel.maPhieu, PhieuDeNghiModel.maQuiTrinhDuyet);
            LinqHeThongDataContext ht = new LinqHeThongDataContext();
            string hoTen = (ht.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == GetUser().manv).Select(d => d.ho).FirstOrDefault() ?? string.Empty) + " " + (ht.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == GetUser().manv).Select(d => d.ten).FirstOrDefault() ?? string.Empty);
            ViewBag.HoTen = hoTen;
            int trangThaiHT = (int?)ht.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == id).Select(d => d.trangThai).FirstOrDefault() ?? 0;
            ViewBag.TenTrangThaiDuyet = ht.tbl_HT_QuiTrinhDuyet_BuocDuyets.Where(d => d.id == trangThaiHT).Select(d => d.maBuocDuyet).FirstOrDefault() ?? string.Empty;
            var maBuocDuyet = ht.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == PhieuDeNghiModel.maPhieu).FirstOrDefault();
            ViewBag.TenBuocDuyet = ht.tbl_HT_QuiTrinhDuyet_BuocDuyets.Where(d => d.id == (maBuocDuyet == null ? 0 : maBuocDuyet.trangThai)).Select(d => d.maBuocDuyet).FirstOrDefault() ?? string.Empty;
            //Tất cả Nhân viên phòng nhân sự duyệt
            NhanVienQLNSDuyet();

            buitlChucDanh(dataPhieuCongTac.maChucDanhMoi);
            buitlPhanCa(dataPhieuCongTac.maPhanCaMoi);
            buitlBoPhanTinhLuong(dataPhieuCongTac.boPhanTinhLuongMoi);
            ViewBag.URL = Request.Url.AbsoluteUri.ToString();

            return View(PhieuDeNghiModel);
        }

        [HttpPost]
        [ValidateInput(false)]
        public ActionResult Edit(string id, FormCollection coll)
        {
            try
            {
                tblPhieuDeNghi = lqPhieuDN.tbl_NS_PhieuDieuChuyens.Where(d => d.soPhieu == id).FirstOrDefault();

                tblPhieuDeNghi.ngayChuyen = String.IsNullOrEmpty(coll.Get("ngayChuyen")) ? (DateTime?)null : DateTime.ParseExact(coll.Get("ngayChuyen"), "dd/MM/yyyy", CultureInfo.InvariantCulture);

                tblPhieuDeNghi.maPhongBanMoi = tblPhieuDeNghi.maPhongBanCu;
                tblPhieuDeNghi.maChucDanhMoi = coll.Get("maChucDanh");
                tblPhieuDeNghi.maPhanCaMoi = !String.IsNullOrEmpty(coll.Get("maPhanCaMoi")) ? coll.Get("maPhanCaMoi") : string.Empty;
               
                
                tblPhieuDeNghi.boPhanTinhLuongMoi = coll.Get("maBoPhanTinhLuong");

                tblPhieuDeNghi.noiDung = coll.Get("noiDung");
                tblPhieuDeNghi.loaiDieuChuyen = "PhieuBoNhiem";
                lqPhieuDN.SubmitChanges();

                return RedirectToAction("Edit", new { id = tblPhieuDeNghi.soPhieu });
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }
        }

        public ActionResult Details(string id)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            try
            {
                ThongTinPhieuDieuChuyen(id);
                return View(PhieuDeNghiModel);
            }
            catch (Exception ex)
            {
                ViewData["Message"] = ex.Message;
                return View("Error");
            }
        }


        public void ThongTinPhieuDieuChuyen(string id)
        {
            var dataPhieuCongTac = lqPhieuDN.tbl_NS_PhieuDieuChuyens.Where(d => d.soPhieu == id).FirstOrDefault();
            PhieuDeNghiModel = new PhieuDieuChuyen();
            PhieuDeNghiModel.maPhieu = dataPhieuCongTac.soPhieu;
            PhieuDeNghiModel.ngayLap = dataPhieuCongTac.ngayLap;
            PhieuDeNghiModel.maNhanVienLapPhieu = dataPhieuCongTac.nguoiLap;
            PhieuDeNghiModel.NhanVien = new BatDongSan.Models.NhanSu.NhanVienModel
            {
                maNhanVien = dataPhieuCongTac.maNhanVien,
                hoVaTen = HoVaTen(dataPhieuCongTac.maNhanVien),
                gioiTinh = GioiTinh(dataPhieuCongTac.maNhanVien),
            };
            PhieuDeNghiModel.NhanVienLapPhieu = new BatDongSan.Models.NhanSu.NhanVienModel
            {
                maNhanVien = dataPhieuCongTac.nguoiLap,
                hoVaTen = HoVaTen(dataPhieuCongTac.nguoiLap)
            };
            PhieuDeNghiModel.ngayChuyen = dataPhieuCongTac.ngayChuyen;


            PhieuDeNghiModel.maPhongBan = dataPhieuCongTac.maPhongBanMoi;
            PhieuDeNghiModel.tenPhongBan = linqDM.tbl_DM_PhongBans.Where(d => d.maPhongBan == dataPhieuCongTac.maPhongBanMoi).Select(d => d.tenPhongBan).FirstOrDefault();

            PhieuDeNghiModel.maChucDanhMoi = dataPhieuCongTac.maChucDanhMoi;
            PhieuDeNghiModel.tenChucDanhMoi = linqNS.Sys_ChucDanhs.Where(d => d.MaChucDanh == dataPhieuCongTac.maChucDanhMoi).Select(d => d.TenChucDanh).FirstOrDefault();

            PhieuDeNghiModel.maBoPhanTinhLuongMoi = dataPhieuCongTac.boPhanTinhLuongMoi;
            PhieuDeNghiModel.tenBoPhanTinhLuongMoi = linqDM.tbl_NS_KhoiTinhLuongs.Where(d => d.maKhoiTinhLuong == Convert.ToInt32(dataPhieuCongTac.boPhanTinhLuongMoi)).Select(d => d.tenKhoiTinhLuong).FirstOrDefault();

            PhieuDeNghiModel.maPhongBanCu = dataPhieuCongTac.maPhongBanCu;
            PhieuDeNghiModel.tenPhongBanCu = linqDM.tbl_DM_PhongBans.Where(d => d.maPhongBan == dataPhieuCongTac.maPhongBanCu).Select(d => d.tenPhongBan).FirstOrDefault();

            PhieuDeNghiModel.maChucDanhCu = dataPhieuCongTac.maChucDanhCu;
            PhieuDeNghiModel.tenChucDanhCu = linqNS.Sys_ChucDanhs.Where(d => d.MaChucDanh == dataPhieuCongTac.maChucDanhCu).Select(d => d.TenChucDanh).FirstOrDefault();

            PhieuDeNghiModel.maBoPhanTinhLuongCu = dataPhieuCongTac.boPhanTinhLuongCu;
            PhieuDeNghiModel.tenBoPhanTinhLuongCu = linqDM.tbl_NS_KhoiTinhLuongs.Where(d => d.maKhoiTinhLuong == Convert.ToInt32(dataPhieuCongTac.boPhanTinhLuongCu)).Select(d => d.tenKhoiTinhLuong).FirstOrDefault();
            PhieuDeNghiModel.maQuiTrinhDuyet = dataPhieuCongTac.maQuiTrinhDuyet ?? 0;
            PhieuDeNghiModel.maPhanCaMoi = short.Parse(!string.IsNullOrEmpty(dataPhieuCongTac.maPhanCaMoi) ? dataPhieuCongTac.maPhanCaMoi : "0");
            PhieuDeNghiModel.tenPhanCaMoi = linqDM.tbl_NS_PhanCas.Where(d => d.maPhanCa == short.Parse(!string.IsNullOrEmpty(dataPhieuCongTac.maPhanCaMoi) ? dataPhieuCongTac.maPhanCaMoi : "0")).Select(d => d.tenPhanCa).FirstOrDefault();
            PhieuDeNghiModel.tenPhanCaCu = linqDM.tbl_NS_PhanCas.Where(d => d.maPhanCa == short.Parse(!string.IsNullOrEmpty(dataPhieuCongTac.maPhanCaCu) ? dataPhieuCongTac.maPhanCaCu : "0")).Select(d => d.tenPhanCa).FirstOrDefault();
            PhieuDeNghiModel.maPhanCaCu = short.Parse(!string.IsNullOrEmpty(dataPhieuCongTac.maPhanCaCu) ? dataPhieuCongTac.maPhanCaCu : "0");

            PhieuDeNghiModel.noiDung = dataPhieuCongTac.noiDung;

            DMNguoiDuyetController nd = new DMNguoiDuyetController();
            PhieuDeNghiModel.Duyet = nd.GetDetailByMaPhieuTheoQuiTrinhDong(PhieuDeNghiModel.maPhieu, PhieuDeNghiModel.maQuiTrinhDuyet);
            LinqHeThongDataContext ht = new LinqHeThongDataContext();
            string hoTen = (ht.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == GetUser().manv).Select(d => d.ho).FirstOrDefault() ?? string.Empty) + " " + (ht.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == GetUser().manv).Select(d => d.ten).FirstOrDefault() ?? string.Empty);
            ViewBag.HoTen = hoTen;
            int trangThaiHT = (int?)ht.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == id).Select(d => d.trangThai).FirstOrDefault() ?? 0;
            ViewBag.TenTrangThaiDuyet = ht.tbl_HT_QuiTrinhDuyet_BuocDuyets.Where(d => d.id == trangThaiHT).Select(d => d.maBuocDuyet).FirstOrDefault() ?? string.Empty;
            ViewBag.URL = Request.Url.AbsoluteUri.ToString();
            var maBuocDuyet = ht.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == PhieuDeNghiModel.maPhieu).FirstOrDefault();
            ViewBag.TenBuocDuyet = ht.tbl_HT_QuiTrinhDuyet_BuocDuyets.Where(d => d.id == (maBuocDuyet == null ? 0 : maBuocDuyet.trangThai)).Select(d => d.maBuocDuyet).FirstOrDefault() ?? string.Empty;
            //Tất cả Nhân viên phòng nhân sự duyệt
            NhanVienQLNSDuyet();
        }

        public ActionResult ChonNhanVien()
        {
            StringBuilder buildTree = new StringBuilder();
            var phongBans = linqDM.tbl_DM_PhongBans.ToList();
            buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
            ViewBag.NVPB = buildTree.ToString();
            return PartialView("_ChonNhanVien");
        }
        public ActionResult LoadNhanVien(int? page, string searchString, string maPhongBan)
        {
            IList<sp_PB_DanhSachNhanVienResult> phongBan1s;
            phongBan1s = linqDM.sp_PB_DanhSachNhanVien(searchString, maPhongBan).ToList();
            ViewBag.isGet = "True";
            int currentPageIndex = page.HasValue ? page.Value : 1;
            ViewBag.Count = phongBan1s.Count();
            ViewBag.Search = searchString;
            ViewBag.MaPhongBan = maPhongBan;
            return PartialView("_LoadNhanVien", phongBan1s.ToPagedList(currentPageIndex, 10));
        }
        public void buitlChucDanh(string select)
        {
            IList<BatDongSan.Models.NhanSu.Sys_ChucDanh> chucDanh = linqNS.Sys_ChucDanhs.OrderByDescending(d => d.SoCapBac).ToList();

            chucDanh.Insert(0, new BatDongSan.Models.NhanSu.Sys_ChucDanh() { MaChucDanh = "", TenChucDanh = "[Chọn]" });
            ViewBag.ChucDanh = new SelectList(chucDanh, "MaChucDanh", "TenChucDanh", select);
        }
        public void buitlPhanCa(string select)
        {
            IList<BatDongSan.Models.DanhMuc.tbl_NS_PhanCa> phanCa = linqDM.tbl_NS_PhanCas.ToList();

            phanCa.Insert(0, new BatDongSan.Models.DanhMuc.tbl_NS_PhanCa() { maPhanCa = 0, tenPhanCa = "[Chọn]" });
            ViewBag.PhanCa = new SelectList(phanCa, "maPhanCa", "tenPhanCa", select);
        }
        public void buitlBoPhanTinhLuong(string select)
        {
            var boPhanTinhLuongs = linqDM.GetTable<BatDongSan.Models.DanhMuc.tbl_NS_KhoiTinhLuong>();
            ViewBag.BoPhanTinhLuong = new SelectList(boPhanTinhLuongs, "maKhoiTinhLuong", "tenKhoiTinhLuong", select);


        }
        public void buitlCoSo(string select)
        {
            IList<Sys_ChiNhanhVanPhong> tinhThanh = linqNS.Sys_ChiNhanhVanPhongs.ToList();

            tinhThanh.Insert(0, new Sys_ChiNhanhVanPhong() { maChiNhanh = "", tenChiNhanh = "[Chọn]" });
            ViewBag.ChiNhanhVanPhong = new SelectList(tinhThanh, "maChiNhanh", "tenChiNhanh", select);
        }
        public string IdGenerator()
        {

            StringBuilder sb = new StringBuilder();
            var date = DateTime.Now;
            string lastID = lqPhieuDN.tbl_NS_PhieuDieuChuyens.OrderByDescending(d => d.ngayLap).Select(d => d.soPhieu).FirstOrDefault();
            string nam = date.Year.ToString();
            nam = nam.Remove(0, 2);
            string thang = string.Empty;
            if (date.Month < 10)
            {
                thang = "0" + date.Month;
            }
            else
            {
                thang = date.Month.ToString();
            }
            if (String.IsNullOrEmpty(lastID))
            {
                return "PBN-" + nam + thang + "001";
            }
            else
            {
                int? idSuffix = Convert.ToInt32(lastID.Substring(lastID.Length - 3)) + 1;
                if (idSuffix <= 0 || idSuffix == null)
                    return "PBN-" + nam + thang + "001";
                else
                {
                    sb.Append(idSuffix.ToString());
                    while (sb.Length < 3)
                    {
                        sb.Insert(0, "0");
                    }
                    return "PBN-" + nam + thang + sb.ToString();
                }
            }
        }
        public string GetMax()
        {
            return lqPhieuDN.tbl_NS_PhieuDieuChuyens.OrderByDescending(d=>d.ngayLap).Select(d => d.soPhieu).FirstOrDefault() ?? string.Empty;
        }
        public string HoVaTen(string MaNV)
        {

            return linqNS.tbl_NS_NhanViens.Where(d => d.maNhanVien == MaNV).Select(d => d.ho + " " + d.ten).FirstOrDefault();
        }

        public bool GioiTinh(string MaNV)
        {

            return (linqNS.tbl_NS_NhanViens.Where(d => d.maNhanVien == MaNV).Select(d => d.gioiTinh).FirstOrDefault() ?? true);
        }

        public ActionResult ChonPhongBan()
        {
            StringBuilder buildTree = new StringBuilder();
            IList<tbl_DM_PhongBan> phongBans = linqDM.tbl_DM_PhongBans.ToList();
            buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
            ViewBag.PhongBan = buildTree.ToString();
            return PartialView("_ChonPhongBan");
        }

        public ActionResult GetInforNhanVien(string maNhanVien)
        {
            LinqDanhMucDataContext linqDM = new LinqDanhMucDataContext();
            var NhanVien = linqNS.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == maNhanVien).FirstOrDefault();
            var nhanVienMa = linqNS.tbl_NS_NhanViens.Where(d => d.maNhanVien == maNhanVien).FirstOrDefault();
            var boPhanTinhluong = linqDM.tbl_NS_KhoiTinhLuongs.Where(d => d.maKhoiTinhLuong == nhanVienMa.idKhoiTinhLuong).FirstOrDefault();


            var maPhanCas = linqDM.tbl_NS_PhanCaNhanViens.Where(d => d.maNhanVien == maNhanVien).FirstOrDefault();
            var tenPhanCa = string.Empty;
            if (maPhanCas != null)
            {
                var dataPhanCa = linqDM.tbl_NS_PhanCas.Where(d => d.maPhanCa == maPhanCas.maPhanCa).FirstOrDefault();
                if (dataPhanCa != null) tenPhanCa = dataPhanCa.tenPhanCa;
            }
            return Json(new
            {
                maCoSo = NhanVien.maChiNhanh,
                tenCoSo = NhanVien.tenChiNhanh,
                maPhongBan = NhanVien.maPhongBan,
                tenPhongBan = NhanVien.tenPhongBan,
                maChucVu = NhanVien.maChucDanh,
                tenChucVu = NhanVien.TenChucDanh,
                maBoPhanTinhLuong = boPhanTinhluong != null ? boPhanTinhluong.maKhoiTinhLuong : 0,
                tenBoPhanTinhLuong = boPhanTinhluong != null ? boPhanTinhluong.tenKhoiTinhLuong : string.Empty,
                maPhanCaCu = maPhanCas != null ? maPhanCas.maPhanCa : 0,
                tenPhanCaCu = maPhanCas != null ? tenPhanCa : string.Empty,
            });
        }

        #region Duyệt qui trình động
        [AcceptVerbs(HttpVerbs.Get)]
        public ActionResult ViewsApproval(string id)
        {
            try
            {
                return RedirectToAction("Details", new { id = id });// detail de duyet
            }
            catch
            {
                return View();
            }
        }


        /// <summary>
        /// Duyệt theo qui trình động phiếu nghỉ phép
        /// </summary>
        /// <param name="maPhieu"></param>
        /// <param name="maQuiTrinh"></param>
        /// <param name="sendMail"></param>
        /// <param name="sendSMS"></param>
        /// <returns></returns>
        public ActionResult DuyetTheoQuiTrinhDong(string maPhieu, string maCongViec, bool sendMail, bool sendSMS, string maNhanVien, int soNgayNghiPhep, string lyDo, int idQuiTrinh)
        {
            try
            {
                LinqHeThongDataContext ht = new LinqHeThongDataContext();
                string hoTen = (ht.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == GetUser().manv).Select(d => d.ho).FirstOrDefault() ?? string.Empty) + " " + (ht.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == GetUser().manv).Select(d => d.ten).FirstOrDefault() ?? string.Empty);
                DMNguoiDuyetController _nguoiDuyet = new DMNguoiDuyetController();
                var kq = _nguoiDuyet.DuyeTheoQuiTrinhDong(maPhieu, maCongViec, sendMail, sendSMS, GetUser().manv, hoTen, soNgayNghiPhep, lyDo, idQuiTrinh, string.Empty);
                return kq;
            }
            catch (Exception e)
            {
                return Json("Lỗi: " + e.Message);
            }
        }
        #endregion

        #region In Phiếu điều chuyển bằng template
        public ActionResult InPhieuDieuChuyen(string maPhieu)
        {
            #region Role user
            permission = GetPermission("PhieuDieuChuyen", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            ViewBag.NoiDung = NoiDungPhieuDieuChuyen(maPhieu);
            return PartialView("InPhieuDieuChuyenTemplate");
        }


        public ActionResult In(string maPhieu)
        {
            byte[] fileContent = null;
            var utf8 = new UTF8Encoding();
            var resultFill = NoiDungPhieuDieuChuyen(maPhieu);
            fileContent = utf8.GetBytes(resultFill);
            return File(fileContent, "application/msword", "PhieuBoNhiem" + ".doc");
        }

        public string NoiDungPhieuDieuChuyen(string maPhieu)
        {
            string noiDung = string.Empty;
            ThongTinPhieuDieuChuyen(maPhieu);
            var ds = lqHT.Sys_PrintTemplates.Where(d => d.maMauIn == "DCNS").FirstOrDefault();
            if (ds != null)
            {
                string gioiTinh = PhieuDeNghiModel.NhanVien.gioiTinh == true ? "Ông" : "Bà";
                string ongHoacBa = PhieuDeNghiModel.NhanVien.hoVaTen;
                noiDung = ds.html.Replace("{$ngay}", DateTime.Now.Day.ToString())
                    .Replace("{$thang}", DateTime.Now.Month.ToString())
                    .Replace("{$nam}", DateTime.Now.Year.ToString())
                    .Replace("{$ongHoacBa}", gioiTinh)
                    .Replace("{$hoTenNhanVien}", ongHoacBa)
                    .Replace("{$chucDanh}", PhieuDeNghiModel.tenChucDanhCu)
                    .Replace("{$phongBan}", PhieuDeNghiModel.tenPhongBanCu)
                    .Replace("{$phongBanMoi}", PhieuDeNghiModel.tenPhongBan)
                    .Replace("{$ngayKy}", PhieuDeNghiModel.ngayChuyen.Value.Day.ToString())
                    .Replace("{$thangKy}", PhieuDeNghiModel.ngayChuyen.Value.Month.ToString())
                    .Replace("{$namKy}", PhieuDeNghiModel.ngayChuyen.Value.Year.ToString())
                    ;
            }
            return noiDung;
        }
        #endregion

    }
}
