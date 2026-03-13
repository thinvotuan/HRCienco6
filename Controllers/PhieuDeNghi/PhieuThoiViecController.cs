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
    public class PhieuThoiViecController : ApplicationController
    {
        LinqPhieuDeNghiDataContext lqPhieuDN = new LinqPhieuDeNghiDataContext();
        LinqNhanSuDataContext linqNS = new LinqNhanSuDataContext();
        LinqDanhMucDataContext linqDM = new LinqDanhMucDataContext();
        tbl_NS_PhieuThoiViec tblPhieuDeNghi;
        IList<tbl_NS_PhieuThoiViec> tblPhieuDeNghis;
        PhieuThoiViec PhieuDeNghiModel;
        public const string taskIDSystem = "PhieuThoiViec";//REGWORKVOTE
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
            var tblPhieuDeNghis = lqPhieuDN.sp_NS_PhieuDeNghiThoiViec_Index(trangThai, fromDate, toDate, searchString, userName).ToList();
            int currentPageIndex = page.HasValue ? page.Value : 1;
            ViewBag.Count = tblPhieuDeNghis.Count();
            ViewBag.Search = searchString;
            ViewBag.tuNgay = tuNgay;
            ViewBag.denNgay = tuNgay;
            ViewBag.trangThai = trangThai;
            ViewBag.maNhanVien = GetUser().manv;
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
            var tblPhieuDeNghis = lqPhieuDN.sp_NS_PhieuDeNghiThoiViec_Index(trangThai, fromDate, toDate, qSearch, userName).ToList();
            int currentPageIndex = page.HasValue ? page.Value : 1;
            ViewBag.Count = tblPhieuDeNghis.Count();
            ViewBag.Search = qSearch;
            ViewBag.tuNgay = tuNgay;
            ViewBag.denNgay = denNgay;
            ViewBag.trangThai = trangThai;
            ViewBag.maNhanVien = GetUser().manv;
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

            PhieuDeNghiModel = new PhieuThoiViec();
            PhieuDeNghiModel.maPhieu = GenerateUtil.CheckLetter("DNTV", GetMax());
            PhieuDeNghiModel.ngayLap = DateTime.Now;
            PhieuDeNghiModel.NhanVienLapPhieu = new BatDongSan.Models.NhanSu.NhanVienModel
            {
                maNhanVien = GetUser().manv,
                hoVaTen = HoVaTen(GetUser().manv)
            };
            PhieuDeNghiModel.NhanVien = new BatDongSan.Models.NhanSu.NhanVienModel();
            PhieuDeNghiModel.ngayLap = DateTime.Now;
            PhieuDeNghiModel.ngayThoiViec = DateTime.Now;
            buitlLyDoThoiViec(string.Empty);
            return View(PhieuDeNghiModel);
        }

        [HttpPost]
        [ValidateInput(false)]
        public ActionResult Create(FormCollection coll)
        {
            try
            {
                tblPhieuDeNghi = new tbl_NS_PhieuThoiViec();
                tblPhieuDeNghi.soPhieu = GenerateUtil.CheckLetter("DNTV", GetMax());
                tblPhieuDeNghi.maNhanVien = coll.Get("NhanVien.maNhanVien");

                tblPhieuDeNghi.ngayLap = DateTime.Now;
                tblPhieuDeNghi.ngayThoiViec = String.IsNullOrEmpty(coll.Get("ngayThoiViec")) ? (DateTime?)null : DateTime.ParseExact(coll.Get("ngayThoiViec"), "dd/MM/yyyy", CultureInfo.InvariantCulture);

                tblPhieuDeNghi.maLyDoThoiViec = Convert.ToInt32(coll.Get("lyDoThoiViec"));
                tblPhieuDeNghi.ghiChu = coll.Get("ghiChu");
                tblPhieuDeNghi.nguoiLap = GetUser().manv;

                lqPhieuDN.tbl_NS_PhieuThoiViecs.InsertOnSubmit(tblPhieuDeNghi);
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
                var phieu = lqPhieuDN.tbl_NS_PhieuThoiViecs.Where(d => d.soPhieu == id).FirstOrDefault();
                lqPhieuDN.tbl_NS_PhieuThoiViecs.DeleteOnSubmit(phieu);
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

            var dataPhieuCongTac = lqPhieuDN.tbl_NS_PhieuThoiViecs.Where(d => d.soPhieu == id).FirstOrDefault();
            PhieuDeNghiModel = new PhieuThoiViec();
            PhieuDeNghiModel.maPhieu = dataPhieuCongTac.soPhieu;
            PhieuDeNghiModel.ngayLap = dataPhieuCongTac.ngayLap;

            PhieuDeNghiModel.NhanVienLapPhieu = new BatDongSan.Models.NhanSu.NhanVienModel
            {
                maNhanVien = dataPhieuCongTac.nguoiLap,
                hoVaTen = HoVaTen(dataPhieuCongTac.nguoiLap)
            };

            PhieuDeNghiModel.NhanVien = new BatDongSan.Models.NhanSu.NhanVienModel
            {
                maNhanVien = dataPhieuCongTac.maNhanVien,
                hoVaTen = HoVaTen(dataPhieuCongTac.maNhanVien)
            };
            PhieuDeNghiModel.ngayThoiViec = dataPhieuCongTac.ngayThoiViec;
            PhieuDeNghiModel.ngayLap = dataPhieuCongTac.ngayLap;
            PhieuDeNghiModel.ghiChu = dataPhieuCongTac.ghiChu;
            PhieuDeNghiModel.maQuiTrinhDuyet = dataPhieuCongTac.maQuiTrinhDuyet ?? 0;
            buitlLyDoThoiViec(dataPhieuCongTac.maLyDoThoiViec.ToString());

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
            return View(PhieuDeNghiModel);
        }

        [HttpPost]
        [ValidateInput(false)]
        public ActionResult Edit(string id, FormCollection coll)
        {
            try
            {
                tblPhieuDeNghi = lqPhieuDN.tbl_NS_PhieuThoiViecs.Where(d => d.soPhieu == id).FirstOrDefault();

                tblPhieuDeNghi.ngayThoiViec = String.IsNullOrEmpty(coll.Get("ngayThoiViec")) ? (DateTime?)null : DateTime.ParseExact(coll.Get("ngayThoiViec"), "dd/MM/yyyy", CultureInfo.InvariantCulture);

                tblPhieuDeNghi.maLyDoThoiViec = Convert.ToInt32(coll.Get("lyDoThoiViec"));
                tblPhieuDeNghi.ghiChu = coll.Get("ghiChu");
                lqPhieuDN.SubmitChanges();

                return RedirectToAction("Edit", new { id = tblPhieuDeNghi.soPhieu });
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }
        }


        public ActionResult checkNgayBatDau(string ngayBatDau, string maNhanVien)
        {
            DateTime fromDate = DateTime.ParseExact(ngayBatDau, "dd/MM/yyyy", null);
            DateTime ngayLap = DateTime.Now;
            var ls = lqPhieuDN.tbl_NS_PhieuCongTacs.Where(t => t.maNhanVien.Equals(maNhanVien) && (t.ngayBatDau == fromDate ||
                (t.ngayBatDau < fromDate && t.ngayKetThuc != null && fromDate <= t.ngayKetThuc))).ToList();

            if ((ngayLap - fromDate).Days > 2)
                return Json(1);
            else if (ls.Count() > 0)
                return Json(2);
            return Json(0);
        }

        public JsonResult Generate(string dateOne, string dateTwo)
        {
            try
            {
                DateTime? startDate = DateTime.ParseExact(dateOne, "dd/MM/yyyy", null);
                DateTime? endDate = DateTime.ParseExact(dateTwo, "dd/MM/yyyy", null);

                TimeSpan timeSpan = endDate.Value.Date - startDate.Value.Date;

                return Json(new { soNgay = timeSpan.Days + 1 });
            }
            catch (Exception ex)
            {
                return Json(ex.Message);
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

            var dataPhieuCongTac = lqPhieuDN.tbl_NS_PhieuThoiViecs.Where(d => d.soPhieu == id).FirstOrDefault();
            PhieuDeNghiModel = new PhieuThoiViec();
            PhieuDeNghiModel.maPhieu = dataPhieuCongTac.soPhieu;
            PhieuDeNghiModel.ngayLap = dataPhieuCongTac.ngayLap;

            PhieuDeNghiModel.NhanVienLapPhieu = new BatDongSan.Models.NhanSu.NhanVienModel
            {
                maNhanVien = dataPhieuCongTac.nguoiLap,
                hoVaTen = HoVaTen(dataPhieuCongTac.nguoiLap)
            };

            PhieuDeNghiModel.NhanVien = new BatDongSan.Models.NhanSu.NhanVienModel
            {
                maNhanVien = dataPhieuCongTac.maNhanVien,
                hoVaTen = HoVaTen(dataPhieuCongTac.maNhanVien)
            };
            PhieuDeNghiModel.ngayThoiViec = dataPhieuCongTac.ngayThoiViec;
            PhieuDeNghiModel.ngayLap = dataPhieuCongTac.ngayLap;
            PhieuDeNghiModel.ghiChu = dataPhieuCongTac.ghiChu;
            PhieuDeNghiModel.maQuiTrinhDuyet = dataPhieuCongTac.maQuiTrinhDuyet ?? 0;
            buitlLyDoThoiViec(dataPhieuCongTac.maLyDoThoiViec.ToString());
            ViewBag.TenLyDo = linqDM.tbl_DM_LyDoThoiViecs.Where(d => d.maLyDoThoiViec == dataPhieuCongTac.maLyDoThoiViec).Select(d => d.tenLyDoThoiViec).FirstOrDefault();

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
            return View(PhieuDeNghiModel);
        }

        public ActionResult CreateTrucTiep()
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenThemTrucTiep);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            PhieuDeNghiModel = new PhieuThoiViec();
            PhieuDeNghiModel.maPhieu = GenerateUtil.CheckLetter("DNTV", GetMax());
            PhieuDeNghiModel.ngayLap = DateTime.Now;
            PhieuDeNghiModel.NhanVien = new BatDongSan.Models.NhanSu.NhanVienModel
            {
                maNhanVien = GetUser().manv,
                hoVaTen = HoVaTen(GetUser().manv)
            };
            PhieuDeNghiModel.ngayLap = DateTime.Now;
            PhieuDeNghiModel.ngayThoiViec = DateTime.Now;

            buitlLyDoThoiViec(string.Empty);

            return View(PhieuDeNghiModel);
        }


        [HttpPost]
        [ValidateInput(false)]
        public ActionResult CreateTrucTiep(FormCollection coll)
        {
            try
            {
                tblPhieuDeNghi = new tbl_NS_PhieuThoiViec();
                tblPhieuDeNghi.soPhieu = GenerateUtil.CheckLetter("DNTV", GetMax());
                tblPhieuDeNghi.ngayLap = DateTime.Now;
                tblPhieuDeNghi.maNhanVien = coll.Get("NhanVien.maNhanVien");
                tblPhieuDeNghi.ngayLap = DateTime.Today;
                tblPhieuDeNghi.ngayThoiViec = String.IsNullOrEmpty(coll.Get("ngayThoiViec")) ? (DateTime?)null : DateTime.ParseExact(coll.Get("ngayThoiViec"), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                tblPhieuDeNghi.ghiChu = coll.Get("ghiChu");
                tblPhieuDeNghi.maLyDoThoiViec = Convert.ToInt32(coll.Get("lyDoThoiViec"));
                lqPhieuDN.tbl_NS_PhieuThoiViecs.InsertOnSubmit(tblPhieuDeNghi);
                lqPhieuDN.SubmitChanges();



                //DMNguoiDuyet record = new DMNguoiDuyet();
                //record.maPhieu = tblPhieuDeNghi.maPhieu;
                //record.ngayDuyet = DateTime.Now.Date;
                //record.maCongViec = "RECEIVEREGWORK";
                //record.trangThai = 4;
                //record.nguoiDuyet = new Models.NhanSu.NhanVienModel{maNhanVien=GetUser().manv,hoVaTen = hovatn}
                //new SqlDMNguoiDuyetRepository().Save(record);

                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }
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

        public void buitlLyDoThoiViec(string select)
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();
            var lyDoThoiViec = linqDM.tbl_DM_LyDoThoiViecs.ToList();

            dict.Add("", "[Chọn lý do thôi việc]");
            foreach (var item in lyDoThoiViec)
            {
                dict.Add(item.maLyDoThoiViec.ToString(), item.tenLyDoThoiViec);
            }
            ViewBag.LyDoThoiViec = new SelectList(dict, "Key", "Value", select);
            ViewBag.MaLyDo = select;
        }

        public string GetMax()
        {
            return lqPhieuDN.tbl_NS_PhieuThoiViecs.OrderByDescending(d=>d.ngayLap).Select(d => d.soPhieu).FirstOrDefault() ?? string.Empty;
        }
        public string HoVaTen(string MaNV)
        {

            return linqNS.tbl_NS_NhanViens.Where(d => d.maNhanVien == MaNV).Select(d => d.ho + " " + d.ten).FirstOrDefault();
        }
        public ActionResult ChonPhongBan()
        {
            StringBuilder buildTree = new StringBuilder();
            IList<tbl_DM_PhongBan> phongBans = linqDM.tbl_DM_PhongBans.ToList();
            buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
            ViewBag.PhongBan = buildTree.ToString();
            return PartialView("_ChonPhongBan");
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

    }
}
