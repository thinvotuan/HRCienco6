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
    public class PhieuTangCaController : ApplicationController
    {
        private StringBuilder buildTree = null;
        private IList<tbl_DM_PhongBan> phongBans;
        LinqPhieuDeNghiDataContext lqPhieuDN = new LinqPhieuDeNghiDataContext();
        LinqNhanSuDataContext linqNS = new LinqNhanSuDataContext();
        LinqDanhMucDataContext linqDM = new LinqDanhMucDataContext();
        //tbl_NS_PhieuTangCa tblPhieuDeNghi;        
        //IList<tbl_NS_PhieuTangCa> phieuTangCas;
        tbl_NS_PhieuTangCa phieuTC;
        PhieuDeNghiCongTac PhieuDeNghiModel;
        PhieuDeNghiTangCa phieuTangCa;
        public const string taskIDSystem = "PhieuTangCa";
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
            try
            {
                Administrator(GetUser().manv);
                BindDataTrangThai(taskIDSystem);
                var userName = GetUser().manv;
                DateTime? fromDate = null;
                DateTime? toDate = null;
                if (!String.IsNullOrEmpty(tuNgay))
                    fromDate = DateTime.ParseExact(tuNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                if (!String.IsNullOrEmpty(denNgay))
                    toDate = DateTime.ParseExact(denNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);

                ViewBag.isGet = "True";
                var tblPhieuDeNghis = lqPhieuDN.sp_NS_PhieuDangKyTangCa_Index(trangThai, fromDate, toDate, searchString, userName).ToList();
                int currentPageIndex = page.HasValue ? page.Value : 1;
                ViewBag.Count = tblPhieuDeNghis.Count();
                ViewBag.Search = searchString;
                ViewBag.tuNgay = tuNgay;
                ViewBag.denNgay = tuNgay;
                ViewBag.trangThai = trangThai;
                return View(tblPhieuDeNghis.ToPagedList(currentPageIndex, 50));
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }

        }
        public ActionResult ViewIndex(int? page, string qSearch, string tuNgay, string denNgay, string trangThai)
        {
            try
            {
                Administrator(GetUser().manv);
                var userName = GetUser().manv;
                DateTime? fromDate = null;
                DateTime? toDate = null;
                if (!String.IsNullOrEmpty(tuNgay))
                    fromDate = String.IsNullOrEmpty(tuNgay) ? (DateTime?)null : DateTime.ParseExact(tuNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                if (!String.IsNullOrEmpty(denNgay))
                    toDate = String.IsNullOrEmpty(denNgay) ? (DateTime?)null : DateTime.ParseExact(denNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);

                ViewBag.isGet = "True";
                var tblPhieuDeNghis = lqPhieuDN.sp_NS_PhieuDangKyTangCa_Index(trangThai, fromDate, toDate, qSearch, userName).ToList();
                int currentPageIndex = page.HasValue ? page.Value : 1;
                ViewBag.Count = tblPhieuDeNghis.Count();
                ViewBag.Search = qSearch;
                ViewBag.tuNgay = tuNgay;
                ViewBag.denNgay = denNgay;
                ViewBag.trangThai = trangThai;
                return PartialView("ViewIndex", tblPhieuDeNghis.ToPagedList(currentPageIndex, 50));
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }

        }

        public ActionResult Create()
        {
            ViewBag.phieuTangCa = 1;
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            try
            {
                Administrator(GetUser().manv);

                phieuTangCa = new PhieuDeNghiTangCa();
                phieuTangCa.maPhieu = GenerateUtil.CheckLetter("DKTC", GetMax());
                phieuTangCa.ngayLap = DateTime.Now;
                phieuTangCa.NhanVien = new BatDongSan.Models.NhanSu.NhanVienModel
                {
                    maNhanVien = GetUser().manv,
                    hoVaTen = HoVaTen(GetUser().manv)
                };

                phieuTangCa.maPhongBan = linqNS.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == GetUser().manv).Select(d => d.maPhongBan).FirstOrDefault() ?? string.Empty;
                phieuTangCa.tenPhongBan = linqNS.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == GetUser().manv).Select(d => d.tenPhongBan).FirstOrDefault() ?? string.Empty;
                phieuTangCa.ngayTangCa = DateTime.Now;
                phieuTangCa.soNgayCongTac = 1;
                phieuTangCa.tangCaTinhChamCong = true;

                BindDataLoaiTangCa(0);

                return View(phieuTangCa);
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
                phieuTC = new tbl_NS_PhieuTangCa();
                BindDataToSave(coll, true);
                lqPhieuDN.tbl_NS_PhieuTangCas.InsertOnSubmit(phieuTC);
                lqPhieuDN.SubmitChanges();
                return RedirectToAction("Edit", new { id = phieuTC.soPhieu });
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
                var phieu = lqPhieuDN.tbl_NS_PhieuTangCas.Where(d => d.soPhieu == id);
                lqPhieuDN.tbl_NS_PhieuTangCas.DeleteAllOnSubmit(phieu);

                var delChiTietNV = lqPhieuDN.tbl_NS_PhieuTangCa_DSNhanViens.Where(d => d.soPhieu == id);
                lqPhieuDN.tbl_NS_PhieuTangCa_DSNhanViens.DeleteAllOnSubmit(delChiTietNV);


                LinqHeThongDataContext lqHT = new LinqHeThongDataContext();
                var nguoiDuyet = lqHT.tbl_HT_DMNguoiDuyets.Where(d => d.maPhieu == id);
                lqHT.tbl_HT_DMNguoiDuyets.DeleteAllOnSubmit(nguoiDuyet);

                LuuLichSuCapNhatPhieu(id, taskIDSystem, 1);
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
            ViewBag.phieuTangCa = 1;
            try
            {
                #region Role user
                permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenSua);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion

                GetThongTinPhieuTangCa(id);
                Administrator(GetUser().manv);

                return View(phieuTangCa);
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
            ViewBag.phieuTangCa = 1;
            try
            {
                phieuTC = new tbl_NS_PhieuTangCa();
                BindDataToSave(coll, false);
                LuuLichSuCapNhatPhieu(phieuTC.soPhieu, taskIDSystem, 0);
                lqPhieuDN.SubmitChanges();
                return RedirectToAction("Edit", new { id = phieuTC.soPhieu });
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
                GetThongTinPhieuTangCa(id);
                return View(phieuTangCa);
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }
        }


        //public ActionResult checkNgayBatDau(string ngayBatDau, string maNhanVien)
        //{
        //    DateTime fromDate = DateTime.ParseExact(ngayBatDau, "dd/MM/yyyy", null);
        //    DateTime ngayLap = DateTime.Now;
        //    var ls = lqPhieuDN.tbl_DN_PhieuCongTacs.Where(t => t.maNhanVien.Equals(maNhanVien) && (t.ngayBatDau == fromDate ||
        //        (t.ngayBatDau < fromDate && t.ngayKetThuc != null && fromDate <= t.ngayKetThuc))).ToList();

        //    if ((ngayLap - fromDate).Days > 2)
        //        return Json(1);
        //    else if (ls.Count() > 0)
        //        return Json(2);
        //    return Json(0);
        //}

        //public JsonResult Generate(string dateOne, string dateTwo)
        //{
        //    try
        //    {
        //        DateTime? startDate = DateTime.ParseExact(dateOne, "dd/MM/yyyy", null);
        //        DateTime? endDate = DateTime.ParseExact(dateTwo, "dd/MM/yyyy", null);

        //        TimeSpan timeSpan = endDate.Value.Date - startDate.Value.Date;

        //        return Json(new { soNgay = timeSpan.Days + 1 });
        //    }
        //    catch (Exception ex)
        //    {
        //        return Json(ex.Message);
        //    }
        //}

        //public ActionResult Details(string id)
        //{
        //    #region Role user
        //    permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenSua);
        //    if (!permission.HasValue)
        //        return View("LogIn");
        //    if (!permission.Value)
        //        return View("AccessDenied");
        //    #endregion

        //    var dataPhieuCongTac = lqPhieuDN.tbl_DN_PhieuCongTacs.Where(d => d.maPhieu == id).FirstOrDefault();
        //    PhieuDeNghiModel = new PhieuDeNghiCongTac();
        //    PhieuDeNghiModel.maPhieu = dataPhieuCongTac.maPhieu;
        //    PhieuDeNghiModel.ngayLap = dataPhieuCongTac.ngayLap;
        //    PhieuDeNghiModel.NhanVien = new BatDongSan.Models.NhanSu.NhanVienModel
        //    {
        //        maNhanVien = dataPhieuCongTac.maNhanVien,
        //        hoVaTen = HoVaTen(dataPhieuCongTac.maNhanVien)
        //    };
        //    PhieuDeNghiModel.ngayBatDau = dataPhieuCongTac.ngayBatDau;
        //    PhieuDeNghiModel.ngayKetThuc = dataPhieuCongTac.ngayKetThuc;
        //    PhieuDeNghiModel.soNgayCongTac = dataPhieuCongTac.soNgayCongTac;
        //    PhieuDeNghiModel.phuCap = dataPhieuCongTac.phuCap;
        //    PhieuDeNghiModel.loaiCongTac = dataPhieuCongTac.loaiCongTac;
        //    PhieuDeNghiModel.gioBatDau = dataPhieuCongTac.gioBatDau;
        //    PhieuDeNghiModel.gioKetThuc = dataPhieuCongTac.gioKetThuc;

        //    PhieuDeNghiModel.maTinh = dataPhieuCongTac.noiCongTac;
        //    PhieuDeNghiModel.tenTinh = linqNS.Sys_TinhThanhs.Where(d=>d.maTinhThanh==dataPhieuCongTac.noiCongTac).Select(d=>d.tenTinhThanh).FirstOrDefault();
        //    PhieuDeNghiModel.maQuocGia = dataPhieuCongTac.quocGiaCongTac;
        //    PhieuDeNghiModel.tenQuocGia = linqNS.Sys_QuocTiches.Where(d => d.maQuocTich == dataPhieuCongTac.quocGiaCongTac).Select(d => d.tenQuocTich).FirstOrDefault();

        //    PhieuDeNghiModel.phongBanCongTac = dataPhieuCongTac.phongBanCongTac;
        //    PhieuDeNghiModel.tenPhongBan = linqDM.tbl_DM_PhongBans.Where(d => d.maPhongBan == dataPhieuCongTac.phongBanCongTac).Select(d => d.tenPhongBan).FirstOrDefault();
        //    PhieuDeNghiModel.phuongTien = dataPhieuCongTac.phuongTien;
        //    PhieuDeNghiModel.lyDo = dataPhieuCongTac.lyDo;
        //    PhieuDeNghiModel.ghiChu = dataPhieuCongTac.ghiChu;
        //    buitlTinhThanh(dataPhieuCongTac.noiCongTac);
        //    buitlQuocGia(dataPhieuCongTac.quocGiaCongTac);

        //    DMNguoiDuyetController nd = new DMNguoiDuyetController();
        //    PhieuDeNghiModel.Duyet = nd.GetDetailByMaPhieuTheoQuiTrinhDong(PhieuDeNghiModel.maPhieu, 2);
        //    LinqHeThongDataContext ht = new LinqHeThongDataContext();
        //    string hoTen = (ht.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == GetUser().manv).Select(d => d.ho).FirstOrDefault() ?? string.Empty) + " " + (ht.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == GetUser().manv).Select(d => d.ten).FirstOrDefault() ?? string.Empty);
        //    ViewBag.HoTen = hoTen;
        //    int trangThaiHT = (int?)ht.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == id).Select(d => d.trangThai).FirstOrDefault() ?? 0;
        //    ViewBag.TenTrangThaiDuyet = ht.tbl_HT_QuiTrinhDuyet_BuocDuyets.Where(d => d.id == trangThaiHT).Select(d => d.maBuocDuyet).FirstOrDefault() ?? string.Empty;
        //    return View(PhieuDeNghiModel);
        //}

        //public ActionResult CreateTrucTiep()
        //{
        //    #region Role user
        //    permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenThemTrucTiep);
        //    if (!permission.HasValue)
        //        return View("LogIn");
        //    if (!permission.Value)
        //        return View("AccessDenied");
        //    #endregion

        //    PhieuDeNghiModel = new PhieuDeNghiCongTac();
        //    PhieuDeNghiModel.maPhieu = GenerateUtil.CheckLetter("DKCT", GetMax());
        //    PhieuDeNghiModel.ngayLap = DateTime.Now;
        //    PhieuDeNghiModel.NhanVien = new BatDongSan.Models.NhanSu.NhanVienModel
        //    {
        //        maNhanVien = GetUser().manv,
        //        hoVaTen = HoVaTen(GetUser().manv)
        //    };
        //    PhieuDeNghiModel.ngayBatDau = DateTime.Now;
        //    PhieuDeNghiModel.ngayKetThuc = DateTime.Now;
        //    PhieuDeNghiModel.soNgayCongTac = 1;
        //    PhieuDeNghiModel.phuCap = 0;

        //    buitlTinhThanh(string.Empty);
        //    buitlQuocGia(string.Empty);

        //    return View(PhieuDeNghiModel);
        //}


        //[HttpPost]
        //[ValidateInput(false)]
        //public ActionResult CreateTrucTiep(FormCollection coll)
        //{
        //    try
        //    {
        //        tblPhieuDeNghi = new tbl_DN_PhieuCongTac();
        //        tblPhieuDeNghi.maPhieu = GenerateUtil.CheckLetter("DKCT", GetMax());
        //        tblPhieuDeNghi.ngayLap = DateTime.Now;
        //        tblPhieuDeNghi.maNhanVien = coll.Get("NhanVien.maNhanVien");

        //        tblPhieuDeNghi.ngayBatDau = String.IsNullOrEmpty(coll.Get("ngayBatDau")) ? (DateTime?)null : DateTime.ParseExact(coll.Get("ngayBatDau"), "dd/MM/yyyy", CultureInfo.InvariantCulture);

        //        tblPhieuDeNghi.ngayKetThuc = String.IsNullOrEmpty(coll.Get("ngayKetThuc")) ? (DateTime?)null : DateTime.ParseExact(coll.Get("ngayKetThuc"), "dd/MM/yyyy", CultureInfo.InvariantCulture);


        //        tblPhieuDeNghi.soNgayCongTac = Convert.ToInt32(coll.Get("soNgayCongTac"));
        //        tblPhieuDeNghi.phuCap = Convert.ToDecimal(coll.Get("phuCap"));

        //        string date = String.Format("{0:dd/MM/yyyy}", DateTime.Now);
        //        tblPhieuDeNghi.gioBatDau = Convert.ToDateTime(date + " " + coll.Get("gioBatDau"));
        //        tblPhieuDeNghi.gioKetThuc = Convert.ToDateTime(date + " " + coll.Get("gioKetThuc"));
        //        //tblPhieuDeNghi.loaiCongTac = Convert.ToBoolean(coll.Get("loaiCongTac"));

        //        //if (!string.IsNullOrEmpty(coll.Get("maTinh")))
        //        //{
        //        //    tblPhieuDeNghi.noiCongTac = coll.Get("maTinh");
        //        //}
        //        //if (!string.IsNullOrEmpty(coll.Get("maQuocGia")))
        //        //{
        //        //    tblPhieuDeNghi.quocGiaCongTac = coll.Get("maQuocGia");
        //        //}
        //        tblPhieuDeNghi.lyDo = coll.Get("lyDo");
        //        tblPhieuDeNghi.phuongTien = coll.Get("phuongTien");
        //        tblPhieuDeNghi.trangThai = 1;
        //        tblPhieuDeNghi.ghiChu = coll.Get("ghiChu");

        //        tblPhieuDeNghi.phongBanCongTac = coll.Get("phongBanCongTac");

        //        lqPhieuDN.tbl_DN_PhieuCongTacs.InsertOnSubmit(tblPhieuDeNghi);
        //        lqPhieuDN.SubmitChanges();



        //        //DMNguoiDuyet record = new DMNguoiDuyet();
        //        //record.maPhieu = tblPhieuDeNghi.maPhieu;
        //        //record.ngayDuyet = DateTime.Now.Date;
        //        //record.maCongViec = "RECEIVEREGWORK";
        //        //record.trangThai = 4;
        //        //record.nguoiDuyet = new Models.NhanSu.NhanVienModel{maNhanVien=GetUser().manv,hoVaTen = hovatn}
        //        //new SqlDMNguoiDuyetRepository().Save(record);

        //        return RedirectToAction("Index");
        //    }
        //    catch (Exception ex)
        //    {

        //        ViewData["Message"] = ex.Message;
        //        return View("error");
        //    }
        //}     

        /// <summary>
        /// Hàm get max số phiếu tăng ca
        /// </summary>
        /// <returns></returns>
        public string GetMax()
        {
            return lqPhieuDN.tbl_NS_PhieuTangCas.OrderByDescending(d=>d.ngayLap).Select(d => d.soPhieu).FirstOrDefault();
        }

        /// <summary>
        /// Lấy họ và tên nhân viên
        /// </summary>
        /// <param name="MaNV"></param>
        /// <returns></returns>
        public string HoVaTen(string MaNV)
        {

            return linqNS.tbl_NS_NhanViens.Where(d => d.maNhanVien == MaNV).Select(d => d.ho + " " + d.ten).FirstOrDefault();
        }


        /// <summary>
        /// Binddata loại tăng ca
        /// </summary>
        /// <param name="loaiTangCa"></param>
        public void BindDataLoaiTangCa(int loaiTangCa)
        {
            IList<tbl_DM_LoaiTangCa> loaiTangCas = linqDM.tbl_DM_LoaiTangCas.ToList();
            loaiTangCas.Insert(0, new tbl_DM_LoaiTangCa() { id = 0, loaiTangCa = "[Chọn]" });
            ViewBag.LoaiTangCa = new SelectList(loaiTangCas, "id", "loaiTangCa", loaiTangCa);
        }

        /// <summary>
        /// Lấy hệ số tăng ca của loại tăng ca
        /// </summary>
        /// <param name="maLoaiTangCa"></param>
        /// <returns></returns>
        public JsonResult HeSoTangCa(int maLoaiTangCa)
        {
            double heSoTangCa = linqDM.tbl_DM_LoaiTangCas.Where(d => d.id == maLoaiTangCa).Select(d => d.heSoTangCa).FirstOrDefault() ?? 0;
            return Json(heSoTangCa);
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


        /// <summary>
        /// Thêm và cập nhật thông tin phiếu tăng ca
        /// </summary>
        /// <param name="col"></param>
        /// <param name="isCreate"></param>
        public void BindDataToSave(FormCollection col, bool isCreate)
        {
            if (isCreate == true)
            {
                phieuTC.soPhieu = GenerateUtil.CheckLetter("DKTC", GetMax());
                phieuTC.nguoiLap = GetUser().manv;
                phieuTC.ngayLap = DateTime.Now;
            }
            else
            {
                phieuTC = lqPhieuDN.tbl_NS_PhieuTangCas.Where(d => d.soPhieu == col.Get("maPhieu")).FirstOrDefault();
                //Delete danh sách nhân viên
                var delNV = lqPhieuDN.tbl_NS_PhieuTangCa_DSNhanViens.Where(d => d.soPhieu == col.Get("maPhieu"));
                lqPhieuDN.tbl_NS_PhieuTangCa_DSNhanViens.DeleteAllOnSubmit(delNV);
            }
            phieuTC.ngayTangCa = DateTime.ParseExact(col.Get("ngayTangCa"), "dd/MM/yyyy", null);
            try
            {
                string date = String.Format("{0:MM/dd/yyyy}", phieuTC.ngayTangCa);
                phieuTC.gioBatDau = Convert.ToDateTime(date + " " + col.Get("gioBatDau"));
                phieuTC.gioKetThuc = Convert.ToDateTime(date + " " + col.Get("gioKetThuc"));
            }
            catch
            {
                string date = String.Format("{0:dd/MM/yyyy}", phieuTC.ngayTangCa);
                phieuTC.gioBatDau = Convert.ToDateTime(date + " " + col.Get("gioBatDau"));
                phieuTC.gioKetThuc = Convert.ToDateTime(date + " " + col.Get("gioKetThuc"));
            }

            phieuTC.loaiTangCa = Convert.ToInt32(col.Get("maLoaiTangCa"));
            try
            {
                string date = String.Format("{0:MM/dd/yyyy}", phieuTC.ngayTangCa);
                phieuTC.soGioTangCa = Convert.ToDateTime(date + " " + col.Get("soGioTangCa"));
                phieuTC.batDauNghiGiuaCa = col.Get("thoiGianNghiGiuaCa") == string.Empty ? (DateTime?)null : Convert.ToDateTime(date + " " + col.Get("thoiGianNghiGiuaCa"));
                phieuTC.ketThucNghiGiuaCa = col.Get("thoiGianNghiKetThucGiuaCa") == string.Empty ? (DateTime?)null : Convert.ToDateTime(date + " " + col.Get("thoiGianNghiKetThucGiuaCa"));
            }
            catch
            {
                string date = String.Format("{0:dd/MM/yyyy}", phieuTC.ngayTangCa);
                phieuTC.soGioTangCa = Convert.ToDateTime(date + " " + col.Get("soGioTangCa"));
                phieuTC.batDauNghiGiuaCa = col.Get("thoiGianNghiGiuaCa") == string.Empty ? (DateTime?)null : Convert.ToDateTime(date + " " + col.Get("thoiGianNghiGiuaCa"));
                phieuTC.ketThucNghiGiuaCa = col.Get("thoiGianNghiKetThucGiuaCa") == string.Empty ? (DateTime?)null : Convert.ToDateTime(date + " " + col.Get("thoiGianNghiKetThucGiuaCa"));
            }
            phieuTC.thoiGianNghiGiuaCa = Convert.ToDouble(col.Get("soGioNghiGiuaCa"));
            phieuTC.noiDungTangCa = col.Get("ghiChu");
            phieuTC.hinhThucTangCa = col.Get("hinhThucTangCa") == "tcnb" ? "tcnb" : "tctl";
            phieuTC.tangCaTinhChamCong = Convert.ToBoolean(col.Get("tangCaTinhChamCong") == "True" ? "True" : "False");

            //Insert chi tiết nhân viên tăng ca thuộc phòng ban
            string[] maNhanVien = col.GetValues("maNhanVien");
            List<tbl_NS_PhieuTangCa_DSNhanVien> chiTiet = new List<tbl_NS_PhieuTangCa_DSNhanVien>();
            tbl_NS_PhieuTangCa_DSNhanVien ct;
            if (maNhanVien != null && maNhanVien.Count() > 0)
            {
                for (int i = 0; i < maNhanVien.Count(); i++)
                {
                    ct = new tbl_NS_PhieuTangCa_DSNhanVien();
                    ct.soPhieu = phieuTC.soPhieu;
                    ct.maNhanVien = col.GetValues("maNhanVien")[i];
                    chiTiet.Add(ct);
                }
            }
            else { 
            // Nếu không có nhân viên nào thì tăng ca cho chính nó.
                ct = new tbl_NS_PhieuTangCa_DSNhanVien();
                ct.soPhieu = phieuTC.soPhieu;
                ct.maNhanVien = GetUser().manv;
                chiTiet.Add(ct);
            }
            if (chiTiet != null && chiTiet.Count > 0)
            {
                lqPhieuDN.tbl_NS_PhieuTangCa_DSNhanViens.InsertAllOnSubmit(chiTiet);
            }
        }

        public void GetThongTinPhieuTangCa(string id)
        {
            phieuTangCa = new PhieuDeNghiTangCa();
            var ds = lqPhieuDN.tbl_NS_PhieuTangCas.Where(d => d.soPhieu == id).FirstOrDefault();
            phieuTangCa.maPhieu = ds.soPhieu;
            phieuTangCa.ngayLap = ds.ngayLap;
            phieuTangCa.NhanVien = new BatDongSan.Models.NhanSu.NhanVienModel
            {
                maNhanVien = ds.nguoiLap,
                hoVaTen = HoVaTen(ds.nguoiLap)
            };

            phieuTangCa.maPhongBan = linqNS.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == ds.nguoiLap).Select(d => d.maPhongBan).FirstOrDefault() ?? string.Empty;
            phieuTangCa.tenPhongBan = linqNS.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == ds.nguoiLap).Select(d => d.tenPhongBan).FirstOrDefault() ?? string.Empty;
            phieuTangCa.ngayTangCa = ds.ngayTangCa;
            phieuTangCa.gioBatDau = ds.gioBatDau;
            phieuTangCa.gioKetThuc = ds.gioKetThuc;
            phieuTangCa.soGioTangCa = ds.soGioTangCa;
            BindDataLoaiTangCa(ds.loaiTangCa ?? 0);
            var tangCa = linqDM.tbl_DM_LoaiTangCas.Where(d => d.id == ds.loaiTangCa).FirstOrDefault();
            phieuTangCa.heSoTangCa = (double?)tangCa.heSoTangCa ?? 0;
            phieuTangCa.tenLoaiTangCa = tangCa.loaiTangCa;
            phieuTangCa.thoiGianNghiGiuaCa = ds.batDauNghiGiuaCa;
            phieuTangCa.thoiGianNghiKetThucGiuaCa = ds.ketThucNghiGiuaCa;
            phieuTangCa.soGioNghiGiuaCa = (int)ds.thoiGianNghiGiuaCa;
            phieuTangCa.ghiChu = ds.noiDungTangCa;
            phieuTangCa.tangCaTinhChamCong = ds.tangCaTinhChamCong ?? false;

            phieuTangCa.maQuiTrinhDuyet = ds.maQuiTrinhDuyet ?? 0;
            // Bi Check if maNhanVienDuyet thuoc Tbl_DM_PhongBan lay ma quy trinh duyet khac
            var CheckQT = lqPhieuDN.GetTable<tbl_DM_PhongBan>().Where(d => d.maNhanVienDuyet == ds.nguoiLap).FirstOrDefault();
            if (CheckQT != null)
            {
                phieuTangCa.maQuiTrinhDuyet = 31;
            }
            phieuTangCa.hinhThucTangCa = ds.hinhThucTangCa;
            //Danh sách chi tiêt nhân viên tham dự tăng ca
            ViewBag.ChiTietNhanVien = lqPhieuDN.tbl_NS_PhieuTangCa_DSNhanViens.Where(d => d.soPhieu == id).Select(g => new NhanVienModel
                {
                    maNhanVien = g.maNhanVien,                
                    hoVaTen = HoVaTen(g.maNhanVien),
                    tenChucDanh = lqPhieuDN.GetTable<vw_NS_DanhSachNhanVienTheoPhongBan>().Where(d => d.maNhanVien == g.maNhanVien).Select(d => d.TenChucDanh).FirstOrDefault() ?? string.Empty,
                    tenPhongBan = lqPhieuDN.GetTable<vw_NS_DanhSachNhanVienTheoPhongBan>().Where(d => d.maNhanVien == g.maNhanVien).Select(d => d.tenPhongBan).FirstOrDefault() ?? string.Empty,
                    email = lqPhieuDN.GetTable<vw_NS_DanhSachNhanVienTheoPhongBan>().Where(d => d.maNhanVien == g.maNhanVien).Select(d => d.email).FirstOrDefault() ?? string.Empty,
                    //linqNS.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == g.maNhanVien).Select(d => d.tenPhongBan).FirstOrDefault() ?? string.Empty,
                }).ToList();

            //Duyệt
            DMNguoiDuyetController nd = new DMNguoiDuyetController();
            phieuTangCa.Duyet = nd.GetDetailByMaPhieuTheoQuiTrinhDong(phieuTangCa.maPhieu, phieuTangCa.maQuiTrinhDuyet);
            LinqHeThongDataContext ht = new LinqHeThongDataContext();
            string hoTen = (ht.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == GetUser().manv).Select(d => d.ho).FirstOrDefault() ?? string.Empty) + " " + (ht.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == GetUser().manv).Select(d => d.ten).FirstOrDefault() ?? string.Empty);
            ViewBag.HoTen = hoTen;
            int trangThaiHT = (int?)ht.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == id).Select(d => d.trangThai).FirstOrDefault() ?? 0;
            ViewBag.TenTrangThaiDuyet = ht.tbl_HT_QuiTrinhDuyet_BuocDuyets.Where(d => d.id == trangThaiHT).Select(d => d.maBuocDuyet).FirstOrDefault() ?? string.Empty;
            ViewBag.URL = Request.Url.AbsoluteUri.ToString();
            //Tất cả Nhân viên phòng nhân sự duyệt
            NhanVienQLNSDuyet();
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

    }
}
