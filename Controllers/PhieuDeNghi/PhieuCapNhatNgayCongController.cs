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
    public class PhieuCapNhatNgayCongController : ApplicationController
    {
        private StringBuilder buildTree = null;
        private IList<tbl_DM_PhongBan> phongBans;
        LinqPhieuDeNghiDataContext lqPhieuDN = new LinqPhieuDeNghiDataContext();
        LinqNhanSuDataContext linqNS = new LinqNhanSuDataContext();
        LinqDanhMucDataContext linqDM = new LinqDanhMucDataContext();
        PhieuCapNhatNgayCong phieuCapNhatModel;
        tbl_NS_PhieuCapNhatNgayCong ngayCong;
        public const string taskIDSystem = "PhieuCapNhatNgayCong";
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
            var tblPhieuDeNghis = lqPhieuDN.sp_NS_PhieuCapNhatNgayCong_Index(trangThai, fromDate, toDate, searchString, userName).ToList();
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
                toDate = String.IsNullOrEmpty(denNgay) ? (DateTime?)null : DateTime.ParseExact(denNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);

            ViewBag.isGet = "True";
            var tblPhieuDeNghis = lqPhieuDN.sp_NS_PhieuCapNhatNgayCong_Index(trangThai, fromDate, toDate, qSearch, userName).ToList();
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
            try
            {
                phieuCapNhatModel = new PhieuCapNhatNgayCong();
                thang(DateTime.Now.Month);
                nam(DateTime.Now.Year);
                phieuCapNhatModel.maPhieuCapNhatNgayCong = GenerateUtil.CheckLetter("PCNNC", GetMax());
                phieuCapNhatModel.ngayLap = DateTime.Now;
                phieuCapNhatModel.nguoiLap = new NhanVienModel(GetUser().manv, HoVaTen(GetUser().manv));
                return View(phieuCapNhatModel);
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }
        }

        [HttpPost]
        [ValidateInput(false)]
        public ActionResult Create(FormCollection coll)
        {
            ViewBag.phieuTangCa = 1;
            try
            {
                BindDataToSave(coll, true);
                lqPhieuDN.tbl_NS_PhieuCapNhatNgayCongs.InsertOnSubmit(ngayCong);
                lqPhieuDN.SubmitChanges();
                return RedirectToAction("Edit", new { id = ngayCong.maPhieuCapNhatNgayCong });
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
                var phieu = lqPhieuDN.tbl_NS_PhieuCapNhatNgayCongs.Where(d => d.maPhieuCapNhatNgayCong == id);
                lqPhieuDN.tbl_NS_PhieuCapNhatNgayCongs.DeleteAllOnSubmit(phieu);

                var delChiTietNV = lqPhieuDN.tbl_NS_PhieuCapNhatNgayCong_NhanViens.Where(d => d.maPhieuCapNhatNgayCong == id);
                lqPhieuDN.tbl_NS_PhieuCapNhatNgayCong_NhanViens.DeleteAllOnSubmit(delChiTietNV);


                LinqHeThongDataContext lqHT = new LinqHeThongDataContext();
                var nguoiDuyet = lqHT.tbl_HT_DMNguoiDuyets.Where(d => d.maPhieu == id);
                lqHT.tbl_HT_DMNguoiDuyets.DeleteAllOnSubmit(nguoiDuyet);

                lqPhieuDN.SubmitChanges();
                lqHT.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
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
                GetThongTinPhieuCapNhat(id);
                nam(phieuCapNhatModel.nam);
                thang(phieuCapNhatModel.thang);
                return View(phieuCapNhatModel);
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }
        }

        [HttpPost]
        [ValidateInput(false)]
        public ActionResult Edit(FormCollection coll)
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
                var chiTiet = lqPhieuDN.tbl_NS_PhieuCapNhatNgayCong_NhanViens.Where(d => d.maPhieuCapNhatNgayCong == coll.Get("maPhieuCapNhatNgayCong"));
                lqPhieuDN.tbl_NS_PhieuCapNhatNgayCong_NhanViens.DeleteAllOnSubmit(chiTiet);
                BindDataToSave(coll, false);
                lqPhieuDN.SubmitChanges();
                return RedirectToAction("Edit", new { id = ngayCong.maPhieuCapNhatNgayCong });
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
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            try
            {
                GetThongTinPhieuCapNhat(id);
                return View(phieuCapNhatModel);
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }
        }


        /// <summary>
        /// Hàm get max số phiếu tăng ca
        /// </summary>
        /// <returns></returns>
        public string GetMax()
        {
            return lqPhieuDN.tbl_NS_PhieuCapNhatNgayCongs.OrderByDescending(d=>d.ngayLap).Select(d => d.maPhieuCapNhatNgayCong).FirstOrDefault()??String.Empty;
        }


        /// <summary>
        /// Danh sách nhân viên
        /// </summary>
        /// <param name="page"></param>
        /// <param name="nhomUser"></param>
        /// <param name="searchString"></param>
        /// <returns></returns>
        public ActionResult GetMoreUsers(int? page, string maPB, string searchString)
        {
            string nhanVienNhanSu = AdminNhanSu(GetUser().manv);
            if (nhanVienNhanSu == "true")
            {
                maPB = string.Empty;
            }
            int currentPageIndex = page.HasValue ? page.Value : 1;
            int? tongSoDong = 0;
            ViewBag.Search = searchString ?? string.Empty;
            var users = lqPhieuDN.sp_PhieuDeNghi_DanhSachNhanVienTheoPhongBan_Index(maPB, searchString, currentPageIndex, 20).ToList();
            try
            {
                ViewBag.Count = users[0].tongSoDong;
                tongSoDong = users[0].tongSoDong;
            }
            catch
            {
                ViewBag.Count = 0;
            }
            TempData["Params"] = maPB + "," + searchString;
            return PartialView("PartialUsers", users.ToPagedList(currentPageIndex, 20, true, tongSoDong));
        }

        #region Create, Edit, Details


        /// <summary>
        /// Thêm và cập nhật thông tin phiếu cập nhật ngày công
        /// </summary>
        /// <param name="col"></param>
        /// <param name="isCreate"></param>        
        public void BindDataToSave(FormCollection col, bool isCreate)
        {
            if (isCreate == true)
            {
                ngayCong = new tbl_NS_PhieuCapNhatNgayCong();
                ngayCong.maPhieuCapNhatNgayCong = GenerateUtil.CheckLetter("PCNNC", GetMax());
                ngayCong.ngayLap = DateTime.Now;
                ngayCong.nguoiLap = GetUser().manv;
            }
            else
            {
                ngayCong = lqPhieuDN.tbl_NS_PhieuCapNhatNgayCongs.Where(d => d.maPhieuCapNhatNgayCong == col.Get("maPhieuCapNhatNgayCong")).FirstOrDefault();
            }
            ngayCong.thang = Convert.ToInt32(col.Get("thang"));
            ngayCong.nam = Convert.ToInt32(col.Get("nam"));
            ngayCong.noiDung = col.Get("ghiChu");
            //Insert chi tiết nhân viên tăng ca thuộc phòng ban
            string[] maNhanVien = col.GetValues("maNhanVien");
            List<tbl_NS_PhieuCapNhatNgayCong_NhanVien> chiTiet = new List<tbl_NS_PhieuCapNhatNgayCong_NhanVien>();
            tbl_NS_PhieuCapNhatNgayCong_NhanVien ct;
            if (maNhanVien != null && maNhanVien.Count() > 0)
            {
                for (int i = 0; i < maNhanVien.Count(); i++)
                {
                    ct = new tbl_NS_PhieuCapNhatNgayCong_NhanVien();
                    ct.maPhieuCapNhatNgayCong = ngayCong.maPhieuCapNhatNgayCong;
                    ct.maNhanVien = col.GetValues("maNhanVien")[i];
                    ct.soNgayCong = string.IsNullOrEmpty(col.GetValues("soNgayCong")[i]) ? 0 : Convert.ToDouble(col.GetValues("soNgayCong")[i]);
                    ct.soNgayCongChoViec = string.IsNullOrEmpty(col.GetValues("soNgayCongChoViec")[i]) ? 0 : Convert.ToDouble(col.GetValues("soNgayCongChoViec")[i]);
                    ct.soNgayNghiPhep = string.IsNullOrEmpty(col.GetValues("soNgayNghiPhep")[i]) ? 0 : Convert.ToDouble(col.GetValues("soNgayNghiPhep")[i]);
                    ct.soNgayLe = string.IsNullOrEmpty(col.GetValues("soNgayLe")[i]) ? 0 : Convert.ToDouble(col.GetValues("soNgayLe")[i]);
                    ct.soNgayTangCaThuong = string.IsNullOrEmpty(col.GetValues("soNgayTangCaThuong")[i]) ? 0 : Convert.ToDouble(col.GetValues("soNgayTangCaThuong")[i]);
                    ct.soNgayTangChuNhat = string.IsNullOrEmpty(col.GetValues("soNgayTangChuNhat")[i]) ? 0 : Convert.ToDouble(col.GetValues("soNgayTangChuNhat")[i]);
                    ct.soNgayTangCaLe = string.IsNullOrEmpty(col.GetValues("soNgayTangCaLe")[i]) ? 0 : Convert.ToDouble(col.GetValues("soNgayTangCaLe")[i]);

                    chiTiet.Add(ct);
                }
            }
            if (chiTiet != null && chiTiet.Count > 0)
            {
                lqPhieuDN.tbl_NS_PhieuCapNhatNgayCong_NhanViens.InsertAllOnSubmit(chiTiet);
            }
        }

        /// <summary>
        /// Thông tin phiếu cập nhật ngày công
        /// </summary>
        /// <param name="id"></param>
        public void GetThongTinPhieuCapNhat(string id)
        {
            phieuCapNhatModel = new PhieuCapNhatNgayCong();
            DMNguoiDuyetController nguoiDuyet = new DMNguoiDuyetController();
            var ds = lqPhieuDN.tbl_NS_PhieuCapNhatNgayCongs.Where(d => d.maPhieuCapNhatNgayCong == id).FirstOrDefault();
            if (ds != null)
            {
                phieuCapNhatModel.maPhieuCapNhatNgayCong = ds.maPhieuCapNhatNgayCong;
                phieuCapNhatModel.ngayLap = ds.ngayLap.Value;
                phieuCapNhatModel.nguoiLap = new NhanVienModel(ds.nguoiLap, HoVaTen(ds.nguoiLap));
                phieuCapNhatModel.ghiChu = ds.noiDung;
                phieuCapNhatModel.maQuiTrinhDuyet = ds.maQuiTrinhDuyet ?? 0;
                phieuCapNhatModel.thang = ds.thang ?? DateTime.Now.Month;
                phieuCapNhatModel.nam = ds.nam ?? DateTime.Now.Year;
                phieuCapNhatModel.Duyet = nguoiDuyet.GetDetailByMaPhieuTheoQuiTrinhDong(ds.maPhieuCapNhatNgayCong, phieuCapNhatModel.maQuiTrinhDuyet);
                //Danh sách chi tiêt nhân viên cập nhật ngày công
                ViewBag.ChiTietNhanVien = lqPhieuDN.tbl_NS_PhieuCapNhatNgayCong_NhanViens.Where(d => d.maPhieuCapNhatNgayCong == id).Select(g => new PhieuCapNhatNgayCongNhanVien
                {
                    nhanVien = new NhanVienModel
                    {
                        maNhanVien = g.maNhanVien,
                        hoVaTen = HoVaTen(g.maNhanVien),
                        tenChucDanh = lqPhieuDN.GetTable<vw_NS_DanhSachNhanVienTheoPhongBan>().Where(d => d.maNhanVien == g.maNhanVien).Select(d => d.TenChucDanh).FirstOrDefault() ?? string.Empty,
                        tenPhongBan = lqPhieuDN.GetTable<vw_NS_DanhSachNhanVienTheoPhongBan>().Where(d => d.maNhanVien == g.maNhanVien).Select(d => d.tenPhongBan).FirstOrDefault() ?? string.Empty,
                    },
                    soNgayCong = g.soNgayCong ?? 0,
                    soNgayCongChoViec = g.soNgayCongChoViec ?? 0,
                    soNgayNghiPhep = g.soNgayNghiPhep ?? 0,
                    soNgayLe = g.soNgayLe ?? 0,
                    soNgayTangCaThuong = g.soNgayTangCaThuong ?? 0,
                    soNgayTangChuNhat = g.soNgayTangChuNhat ?? 0,
                    soNgayTangCaLe = g.soNgayTangCaLe ?? 0,
                }).ToList();

                //Duyệt
                DMNguoiDuyetController nd = new DMNguoiDuyetController();
                phieuCapNhatModel.Duyet = nd.GetDetailByMaPhieuTheoQuiTrinhDong(phieuCapNhatModel.maPhieuCapNhatNgayCong, phieuCapNhatModel.maQuiTrinhDuyet);
                LinqHeThongDataContext ht = new LinqHeThongDataContext();
                ViewBag.HoTen = HoVaTen(GetUser().manv);
                int trangThaiHT = (int?)ht.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == ds.maPhieuCapNhatNgayCong).Select(d => d.trangThai).FirstOrDefault() ?? 0;
                ViewBag.TenTrangThaiDuyet = ht.tbl_HT_QuiTrinhDuyet_BuocDuyets.Where(d => d.id == trangThaiHT).Select(d => d.maBuocDuyet).FirstOrDefault() ?? string.Empty;
                ViewBag.URL = Request.Url.AbsoluteUri.ToString();
                var maBuocDuyet = ht.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == ds.maPhieuCapNhatNgayCong).FirstOrDefault();
                ViewBag.TenBuocDuyet = ht.tbl_HT_QuiTrinhDuyet_BuocDuyets.Where(d => d.id == (maBuocDuyet == null ? 0 : maBuocDuyet.trangThai)).Select(d => d.maBuocDuyet).FirstOrDefault() ?? string.Empty;
            }
        }

        #endregion
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

        #region Cây sơ đồ tổ chức nhân viên và phòng ban

        //Load danh sach nhan vien khi click button add them:
        public ActionResult DanhSachNhanVien()
        {
            try
            {
                buildTree = new StringBuilder();
                phongBans = linqDM.tbl_DM_PhongBans.ToList();
                buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
                ViewBag.departments = buildTree.ToString();
                ViewBag.shiftType = linqDM.tbl_NS_PhanCas.Where(t => t.tenPhanCa != "").ToList();
                ViewBag.page = 0;
                ViewBag.total = 0;
                return View(phongBans);
            }
            catch (Exception ex)
            {

                ViewBag.Message = ex.Message;
                return View("Error");
            }
        }
        public ActionResult LoadNhanVien(string qSearch, int _page, string parrentId)
        {
            try
            {

                string parentID = linqDM.tbl_DM_PhongBans.Where(d => d.maPhongBan == parrentId).Select(d => d.maPhongBan).FirstOrDefault() ?? string.Empty;
                if (String.IsNullOrEmpty(parentID))
                {
                    parrentId = string.Empty;
                }
                int page = _page == 0 ? 1 : _page;
                int pIndex = page;
                int total = linqDM.sp_PB_DanhSachNhanVien(qSearch, parrentId).Count();
                PagingLoaderController("/PhieuTangCa/Index/", total, page, "?qsearch=" + qSearch + "&parrentId=" + parrentId);
                ViewBag.nhanVien = linqDM.sp_PB_DanhSachNhanVien(qSearch, parrentId).Skip(start).Take(offset).ToList();
                ViewBag.parrentId = parrentId;
                ViewBag.qSearch = qSearch ?? string.Empty;
                ViewBag.currentMaNV = GetUser().manv;
                return PartialView("LoadNhanVien");
            }
            catch (Exception e)
            {
                ViewData["Message"] = e.Message;
                return View("Error");
            }
        }


        #endregion

        private void thang(int value)
        {
            Dictionary<int, string> dics = new Dictionary<int, string>();
            for (int i = 1; i < 13; i++)
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

    }
}
