using BatDongSan.Helper.Common;
using BatDongSan.Helper.Utils;
using BatDongSan.Models.DanhMuc;
using BatDongSan.Models.NhanSu;
using BatDongSan.Models.PhieuDeNghi;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;


namespace BatDongSan.Controllers.PhieuDeNghi
{
    public class PhieuDanhGiaNhanVienPhongBanController : ApplicationController
    {
        private StringBuilder buildTree = null;
        private IList<tbl_DM_PhongBan> phongBans;
        LinqPhieuDeNghiDataContext linqPDN = new LinqPhieuDeNghiDataContext();
        LinqDanhMucDataContext linqDM = new LinqDanhMucDataContext();
        private LinqNhanSuDataContext linqNS = new LinqNhanSuDataContext();
        public const string taskIDSystem = "DanhGiaLVTN";
        public bool? permission;
        private PhieuDanhGiaNVPBModel modelPDG;
        private tbl_NS_PhieuDanhGiaNhanVien_LamViecTaiNha phieuDG;
        //
        // GET: /PhieuDanhGiaNhanVienPhongBan/

        #region Danh sách phiếu đánh giá nhân viên làm việc tại nhà
        public ActionResult Index()
        {
            try
            {
                #region Role user
                permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXem);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion
                return View();
            }
            catch (Exception ex)
            {
                return View("error");
            }

        }

        public ActionResult ViewIndex(string qSearch, string tuNgay, string denNgay, string trangThai, int _page = 0)
        {


            try
            {
                #region Role user
                permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXem);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion

                var userName = GetUser().manv;
                DateTime? fromDate = null;
                DateTime? toDate = null;
                if (!String.IsNullOrEmpty(tuNgay))
                {
                    fromDate = String.IsNullOrEmpty(tuNgay) ? (DateTime?)null : DateTime.ParseExact(tuNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                }
                if (!String.IsNullOrEmpty(denNgay))
                {
                    toDate = String.IsNullOrEmpty(denNgay) ? (DateTime?)null : DateTime.ParseExact(denNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                }

                int page = _page == 0 ? 1 : _page;
                int pIndex = page;

                int total;
                var datas = linqPDN.sp_NS_PhieuDanhGia_LamViecTaiNha_Index(GetUser().manv, fromDate, toDate, qSearch, page, 50).ToList();
                total = datas.Count == 0 ? 0 : (datas[0].TongSoDong ?? 0);

                PagingLoaderController(string.Empty, total, page, string.Empty);

                ViewData["lsDanhSach"] = datas;

                return PartialView("ViewIndex");
            }
            catch (Exception ex)
            {

                return PartialView("error");
            }


        }
        #endregion

        #region Thêm, xóa, sủa phiếu đánh giá nhân viên phòng ban
        public ActionResult Create()
        {
            try
            {
                #region Role user
                permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenThem);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion
                ThongTinPhieuDanhGia(string.Empty);
                return View(modelPDG);
            }
            catch (Exception ex)
            {
                return View("error");
            }

        }

        [HttpPost]
        [ValidateInput(false)]
        public ActionResult Create(FormCollection coll)
        {
            try
            {
                #region Role user
                permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenThem);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion
                BindDataTosave(coll, string.Empty);
                linqPDN.tbl_NS_PhieuDanhGiaNhanVien_LamViecTaiNhas.InsertOnSubmit(phieuDG);
                linqPDN.SubmitChanges();
                return RedirectToAction("Edit", new { id = phieuDG.MaPhieu });
            }
            catch (Exception ex)
            {
                return View("error");
            }

        }

        public ActionResult Edit(string id)
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
                ThongTinPhieuDanhGia(id);
                return View(modelPDG);
            }
            catch (Exception ex)
            {
                return View("error");
            }
        }

        [HttpPost]
        [ValidateInput(false)]
        public ActionResult Edit(FormCollection coll)
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
                BindDataTosave(coll, coll.Get("MaPhieu"));
                linqPDN.SubmitChanges();
                return RedirectToAction("Edit", new { id = phieuDG.MaPhieu });
            }
            catch (Exception ex)
            {
                return View("error");
            }
        }

        [HttpPost]
        public ActionResult Delete(string id)
        {

            try
            {
                #region Role user
                permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXoa);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion

                var delPhieu = linqPDN.tbl_NS_PhieuDanhGiaNhanVien_LamViecTaiNhas.Where(d => d.MaPhieu == id);
                linqPDN.tbl_NS_PhieuDanhGiaNhanVien_LamViecTaiNhas.DeleteAllOnSubmit(delPhieu);

                var delChiTiet = linqPDN.tbl_NS_PhieuDanhGiaNhanVien_LamViecTaiNha_ChiTiets.Where(d => d.MaPhieu == id);
                linqPDN.tbl_NS_PhieuDanhGiaNhanVien_LamViecTaiNha_ChiTiets.DeleteAllOnSubmit(delChiTiet);

                linqPDN.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                //write log
                Log4Net.WriteLog(log4net.Core.Level.Error, ex.Message);

                return View("error");
            }

        }

        public void BindDataTosave(FormCollection col, string maPhieu)
        {
            if (string.IsNullOrEmpty(maPhieu))
            {
                phieuDG = new tbl_NS_PhieuDanhGiaNhanVien_LamViecTaiNha();
                phieuDG.MaPhieu = GenerateUtil.CheckLetter("PDGNV", GetMax());
                phieuDG.NgayLap = DateTime.Now;
                phieuDG.MaNVLPhieu = GetUser().manv;
            }
            else
            {
                phieuDG = linqPDN.tbl_NS_PhieuDanhGiaNhanVien_LamViecTaiNhas.Where(d => d.MaPhieu == maPhieu).FirstOrDefault();
            }
            phieuDG.TuNgay = DateTime.ParseExact(col.Get("TuNgay"), "dd/MM/yyyy", null);
            phieuDG.DenNgay = DateTime.ParseExact(col.Get("DenNgay"), "dd/MM/yyyy", null);
            phieuDG.NoiDung = col.Get("NoiDung");

            //Insert chi tiết công việc
            var delCTCV = linqPDN.tbl_NS_PhieuDanhGiaNhanVien_LamViecTaiNha_ChiTiets.Where(d => d.MaPhieu == maPhieu);
            linqPDN.tbl_NS_PhieuDanhGiaNhanVien_LamViecTaiNha_ChiTiets.DeleteAllOnSubmit(delCTCV);
            string[] maNhanVien = col.GetValues("maNhanVien");
            if (maNhanVien != null && maNhanVien.Count() > 0)
            {
                List<tbl_NS_PhieuDanhGiaNhanVien_LamViecTaiNha_ChiTiet> listPhieu = new List<tbl_NS_PhieuDanhGiaNhanVien_LamViecTaiNha_ChiTiet>();
                tbl_NS_PhieuDanhGiaNhanVien_LamViecTaiNha_ChiTiet ct;
                for (int i = 0; i < maNhanVien.Count(); i++)
                {
                    ct = new tbl_NS_PhieuDanhGiaNhanVien_LamViecTaiNha_ChiTiet();
                    ct.MaPhieu = phieuDG.MaPhieu;
                    ct.MaNhanVien = maNhanVien[i];
                    ct.TyLeHoanThanhCongViec = string.IsNullOrEmpty(col.GetValues("TyLeHoanThanh")[i]) ? 0 : Convert.ToDouble(col.GetValues("TyLeHoanThanh")[i]);
                    ct.GhiChu = col.GetValues("GhiChu")[i];
                    ct.keHoachLamViec = col.GetValues("KeHoachLamViecTaiNha")[i];
                    listPhieu.Add(ct);
                }
                linqPDN.tbl_NS_PhieuDanhGiaNhanVien_LamViecTaiNha_ChiTiets.InsertAllOnSubmit(listPhieu);
            }
        }
        #endregion

        #region Thông tin phiếu đánh giá nhân viên phòng ban
        public void ThongTinPhieuDanhGia(string maPhieu)
        {
            if (string.IsNullOrEmpty(maPhieu))
            {
                modelPDG = new PhieuDanhGiaNVPBModel();
                modelPDG.MaPhieu = GenerateUtil.CheckLetter("PDGNV", GetMax());
                modelPDG.NgayLap = DateTime.Now;
                modelPDG.MaNhanVienLP = GetUser().manv;
                modelPDG.HoTenNVLP = HoVaTen(modelPDG.MaNhanVienLP);
                string parrentId = linqNS.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == GetUser().manv).Select(d => d.maPhongBan).FirstOrDefault() ?? string.Empty;
                var listNVPB = linqDM.sp_PB_DanhSachNhanVien(string.Empty, parrentId).ToList();
                ViewBag.ListNVPB = listNVPB;
            }
            else
            {
                modelPDG = (from lv in linqPDN.tbl_NS_PhieuDanhGiaNhanVien_LamViecTaiNhas
                            join nv in linqPDN.GetTable<tbl_NS_NhanVien>() on lv.MaNVLPhieu equals nv.maNhanVien into dtDong
                            from nv2 in dtDong.DefaultIfEmpty()
                            where lv.MaPhieu == maPhieu
                            select new PhieuDanhGiaNVPBModel
                            {
                                MaPhieu = lv.MaPhieu,
                                NgayLap = lv.NgayLap ?? DateTime.Now,
                                MaNhanVienLP = lv.MaNVLPhieu,
                                NoiDung = lv.NoiDung,
                                HoTenNVLP = nv2.ho + " " + nv2.ten,
                                TuNgay = lv.TuNgay,
                                DenNgay = lv.DenNgay,
                                XacNhan = lv.XacNhan ?? 0,
                            }).FirstOrDefault();
                string quanLyCC = AdminNhanSu(GetUser().manv);
                if (modelPDG.MaNhanVienLP == GetUser().manv || quanLyCC == "true")
                {
                    ViewBag.ListChiTiet = (from lv in linqPDN.tbl_NS_PhieuDanhGiaNhanVien_LamViecTaiNha_ChiTiets
                                           join nv in linqPDN.GetTable<tbl_NS_NhanVien>() on lv.MaNhanVien equals nv.maNhanVien into dtDong
                                           from nv2 in dtDong.DefaultIfEmpty()

                                           join pb in linqPDN.GetTable<vw_NS_DanhSachNhanVienTheoPhongBan>() on lv.MaNhanVien equals pb.maNhanVien into dtDong2
                                           from pb2 in dtDong2.DefaultIfEmpty()

                                           where lv.MaPhieu == maPhieu
                                           select new PhieuDanhGiaChiTiet
                                           {
                                               MaPhieu = lv.MaPhieu,
                                               HoTenNV = nv2.ho + " " + nv2.ten,
                                               GhiChu = lv.GhiChu,
                                               TyLeHoanThanh = lv.TyLeHoanThanhCongViec ?? 0,
                                               TenPhongBan = pb2.tenPhongBan,
                                               MaNhanVien = lv.MaNhanVien,
                                               KeHoachLamViecTaiNha = lv.keHoachLamViec ?? string.Empty,
                                           }).ToList();
                }
                else
                {
                    ViewBag.ListChiTiet = (from lv in linqPDN.tbl_NS_PhieuDanhGiaNhanVien_LamViecTaiNha_ChiTiets
                                           join nv in linqPDN.GetTable<tbl_NS_NhanVien>() on lv.MaNhanVien equals nv.maNhanVien into dtDong
                                           from nv2 in dtDong.DefaultIfEmpty()

                                           join pb in linqPDN.GetTable<vw_NS_DanhSachNhanVienTheoPhongBan>() on lv.MaNhanVien equals pb.maNhanVien into dtDong2
                                           from pb2 in dtDong2.DefaultIfEmpty()

                                           where lv.MaPhieu == maPhieu && lv.MaNhanVien == GetUser().manv
                                           select new PhieuDanhGiaChiTiet
                                           {
                                               MaPhieu = lv.MaPhieu,
                                               HoTenNV = nv2.ho + " " + nv2.ten,
                                               GhiChu = lv.GhiChu,
                                               TyLeHoanThanh = lv.TyLeHoanThanhCongViec ?? 0,
                                               TenPhongBan = pb2.tenPhongBan,
                                               MaNhanVien = lv.MaNhanVien,
                                               KeHoachLamViecTaiNha = lv.keHoachLamViec ?? string.Empty,
                                           }).ToList();
                }


            }
        }

        public string GetMax()
        {
            return linqPDN.tbl_NS_PhieuDanhGiaNhanVien_LamViecTaiNhas.OrderByDescending(d => d.NgayLap).Select(d => d.MaPhieu).FirstOrDefault();
        }
        #endregion

        #region Load danh sách nhân viên phòng ban

        public ActionResult DanhSachNhanVien()
        {
            try
            {


                buildTree = new StringBuilder();
                phongBans = linqDM.tbl_DM_PhongBans.OrderBy(d => d.tenPhongBan).ToList();
                buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
                ViewBag.departments = buildTree.ToString();
                ViewBag.shiftType = linqDM.tbl_NS_PhanCas.Where(t => t.tenPhanCa != "").ToList();
                ViewBag.page = 0;
                ViewBag.total = 0;
                return View(phongBans);
            }
            catch (Exception ex)
            {
                //write log
                Log4Net.WriteLog(log4net.Core.Level.Error, ex.Message);

                return View("error");
            }

        }

        public ActionResult LoadNhanVien(string qSearch, int _page, string parrentId)
        {
            try
            {
                #region Role user
                permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXemChiTiet);
                if (!permission.HasValue)
                    return PartialView("LogIn");
                if (!permission.Value)
                    return PartialView("AccessDenied");
                #endregion
                parrentId = linqNS.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == GetUser().manv).Select(d => d.maPhongBan).FirstOrDefault() ?? string.Empty;
                //string parentID = linqDM.tbl_DM_PhongBans.Where(d => d.maPhongBan == parrentId).Select(d => d.maPhongBan).FirstOrDefault() ?? string.Empty;
                //if (String.IsNullOrEmpty(parentID))
                //{
                //    parrentId = string.Empty;
                //}
                int page = _page == 0 ? 1 : _page;
                int pIndex = page;
                int total = linqDM.sp_PB_DanhSachNhanVien(qSearch, null).Count();
                PagingLoaderController("/PhieuDanhGiaNhanVienPhongBan/LoadNhanVien/", total, page, "?qsearch=" + qSearch + "&parrentId=" + parrentId);
                ViewBag.nhanVien = linqDM.sp_PB_DanhSachNhanVien(qSearch, null).Skip(start).Take(offset).ToList();
                ViewBag.parrentId = parrentId;
                ViewBag.qSearch = qSearch ?? string.Empty;
                ViewBag.currentMaNV = GetUser().manv;

                return PartialView("LoadNhanVien");
            }
            catch (Exception ex)
            {
                //write log
                Log4Net.WriteLog(log4net.Core.Level.Error, ex.Message);

                return PartialView("error");
            }

        }
        #endregion

        #region Xác nhận gửi đến hành chính nhân sự
        public JsonResult XacNhanGuiDenHSNS(string id)
        {
            try
            {
                var update = linqPDN.tbl_NS_PhieuDanhGiaNhanVien_LamViecTaiNhas.Where(d => d.MaPhieu == id).FirstOrDefault();
                update.XacNhan = 1;
                linqPDN.SubmitChanges();
                return Json(string.Empty);
            }
            catch
            {
                return Json("Error");
            }
        }

        public JsonResult XacNhanKeHoachDenNVPB(string id)
        {
            try
            {
                var update = linqPDN.tbl_NS_PhieuDanhGiaNhanVien_LamViecTaiNhas.Where(d => d.MaPhieu == id).FirstOrDefault();
                update.XacNhan = 2;
                linqPDN.SubmitChanges();
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
