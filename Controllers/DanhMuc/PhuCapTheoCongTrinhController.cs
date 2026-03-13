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
using BatDongSan.Models.NhanSu;
using BatDongSan.Utils.Paging;
namespace BatDongSan.Controllers.DanhMuc
{
    public class PhuCapTheoCongTrinhController : ApplicationController
    {
        private LinqDanhMucDataContext context = new LinqDanhMucDataContext();
        private IList<tbl_NS_PhuCapTheoCongTrinh> _DanhSachPhuCaps;
        private IList<sp_NS_PhuCapTheoCongTrinh_IndexResult> danhSachCNs;
        private tbl_NS_PhuCapTheoCongTrinh _DanhSachPhuCap;
        private StringBuilder buildTree = null;
        private IList<tbl_DM_PhongBan> phongBans;
        private readonly string MCV = "PhuCapTheoCongTrinh";
        private bool? permission;
        //
        // GET: /DanhSachCongNhan/


        public ActionResult Index(int? page, int? pageSize, string searchString, string maPhongBan, string trangThai)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            try
            {
                BindDataTrangThai(MCV);
                buildTree = new StringBuilder();
                phongBans = context.tbl_DM_PhongBans.ToList();
                buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
                ViewBag.PhongBans = buildTree.ToString();
                int currentPageIndex = page.HasValue ? page.Value : 1;
                pageSize = pageSize ?? 20;

                danhSachCNs = context.sp_NS_PhuCapTheoCongTrinh_Index(null, searchString, maPhongBan, currentPageIndex, 20, trangThai).OrderByDescending(d => d.id).ToList();

                int? tongSoDong = 0;
                try
                {
                    ViewBag.Count = danhSachCNs[0].tongSoDong;
                    tongSoDong = danhSachCNs[0].tongSoDong;
                }
                catch
                {
                    ViewBag.Count = 0;
                }
                if (Request.IsAjaxRequest())
                {
                    return PartialView("PartialIndex", danhSachCNs.ToPagedList(currentPageIndex, 20, true, tongSoDong));
                }

                ViewBag.searchString = searchString;
                ViewBag.maPhongBan = maPhongBan;
                return View(danhSachCNs.ToPagedList(currentPageIndex, 20, true, tongSoDong));
            }
            catch
            {
                return View("Error");
            }
        }

        public void GetDataFromView(FormCollection collection)
        {
            _DanhSachPhuCap.maPhuCap = collection["maPhuCap"];
            _DanhSachPhuCap.tenPhuCap = collection["tenPhuCap"];
            _DanhSachPhuCap.maPhongBan = collection["maPhongBan"];
            _DanhSachPhuCap.ngayApDung = String.IsNullOrEmpty(collection["ngayApDung"]) ? (DateTime?)null : DateTime.ParseExact(collection["ngayApDung"], "dd/MM/yyyy", CultureInfo.InvariantCulture);

            _DanhSachPhuCap.soTien = decimal.Parse(collection["soTien"]);



        }
        public string IdGenerator()
        {

            StringBuilder sb = new StringBuilder();
            var date = DateTime.Now;
            string lastID = context.tbl_NS_PhuCapTheoCongTrinhs.OrderByDescending(d => d.ngayLap).Select(d => d.maPhuCap).FirstOrDefault();
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
                return "PC-" + nam + thang + "001";
            }
            else
            {
                int? idSuffix = Convert.ToInt32(lastID.Substring(lastID.Length - 3)) + 1;
                if (idSuffix <= 0 || idSuffix == null)
                    return "PC-" + nam + thang + "001";
                else
                {
                    sb.Append(idSuffix.ToString());
                    while (sb.Length < 3)
                    {
                        sb.Insert(0, "0");
                    }
                    return "PC-" + nam + thang + sb.ToString();
                }
            }
        }
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
                _DanhSachPhuCap = new tbl_NS_PhuCapTheoCongTrinh();
                ViewBag.maPhuCap = IdGenerator();
                buildTree = new StringBuilder();
                phongBans = context.tbl_DM_PhongBans.ToList();
                buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
                ViewBag.PhongBans = buildTree.ToString();
                return PartialView("Create", _DanhSachPhuCap);
            }
            catch
            {
                return View("Error");
            }
        }

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
                // check ton tai cong nhan
                var checkList = context.tbl_NS_PhuCapTheoCongTrinhs.Where(d => d.maPhuCap == collection["maPhuCap"]).ToList();
                if (checkList.Count() > 0)
                {

                    return View("error");
                }
                _DanhSachPhuCap = new tbl_NS_PhuCapTheoCongTrinh();
                GetDataFromView(collection);
                _DanhSachPhuCap.ngayLap = DateTime.Now;
                _DanhSachPhuCap.nguoiLap = GetUser().manv;
                _DanhSachPhuCap.maPhuCap = collection["maPhuCap"];



                context.tbl_NS_PhuCapTheoCongTrinhs.InsertOnSubmit(_DanhSachPhuCap);
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View("Error");
            }
        }

        public ActionResult Edit(int id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            try
            {
                _DanhSachPhuCap = context.tbl_NS_PhuCapTheoCongTrinhs.Where(s => s.id == id).FirstOrDefault();
                return PartialView("Edit", _DanhSachPhuCap);
            }
            catch
            {
                return View("Error");
            }
        }
        public ActionResult Details(int id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            try
            {
                _DanhSachPhuCap = context.tbl_NS_PhuCapTheoCongTrinhs.Where(s => s.id == id).FirstOrDefault();
                DMNguoiDuyetController nd = new DMNguoiDuyetController();
                BatDongSan.Models.HeThong.LinqHeThongDataContext ht = new BatDongSan.Models.HeThong.LinqHeThongDataContext();
                string hoTen = (ht.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == GetUser().manv).Select(d => d.ho).FirstOrDefault() ?? string.Empty) + " " + (ht.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == GetUser().manv).Select(d => d.ten).FirstOrDefault() ?? string.Empty);
                ViewBag.HoTen = hoTen;
                int trangThaiHT = (int?)ht.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == _DanhSachPhuCap.maPhuCap).Select(d => d.trangThai).FirstOrDefault() ?? 0;
                ViewBag.TenTrangThaiDuyet = ht.tbl_HT_QuiTrinhDuyet_BuocDuyets.Where(d => d.id == trangThaiHT).Select(d => d.maBuocDuyet).FirstOrDefault() ?? string.Empty;
                ViewBag.URL = Request.Url.AbsoluteUri.ToString();
                var maBuocDuyet = ht.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == _DanhSachPhuCap.maPhuCap).FirstOrDefault();
                ViewBag.TenBuocDuyet = ht.tbl_HT_QuiTrinhDuyet_BuocDuyets.Where(d => d.id == (maBuocDuyet == null ? 0 : maBuocDuyet.trangThai)).Select(d => d.maBuocDuyet).FirstOrDefault() ?? string.Empty;
                ViewBag.Duyet = nd.GetDetailByMaPhieuTheoQuiTrinhDong(_DanhSachPhuCap.maPhuCap, _DanhSachPhuCap.maQuiTrinhDuyet ?? 0);
                return View("Details", _DanhSachPhuCap);
            }
            catch
            {
                return View("Error");
            }
        }

        //
        // POST: /ChucDanh/Edit/5

        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            try
            {
                _DanhSachPhuCap = context.tbl_NS_PhuCapTheoCongTrinhs.Where(s => s.id == id).FirstOrDefault();
                _DanhSachPhuCap.ngayLap = _DanhSachPhuCap.ngayLap;
                _DanhSachPhuCap.nguoiLap = _DanhSachPhuCap.nguoiLap;

                //Lưu dữ liệu nhân viên
                GetDataFromView(collection);
                _DanhSachPhuCap.maPhuCap = _DanhSachPhuCap.maPhuCap;

                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View("Error");
            }
        }


        //
        // POST: /ChucDanh/Delete/5

        [HttpPost]
        public ActionResult Delete(int id)
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
                _DanhSachPhuCap = context.tbl_NS_PhuCapTheoCongTrinhs.Where(s => s.id == id).FirstOrDefault();
                context.tbl_NS_PhuCapTheoCongTrinhs.DeleteOnSubmit(_DanhSachPhuCap);
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View("Error");
            }
        }

        [AcceptVerbs(HttpVerbs.Get)]
        public ActionResult ViewsApproval(string id)
        {
            try
            {
                int ID = context.tbl_NS_PhuCapTheoCongTrinhs.Where(d => d.maPhuCap == id).Select(d => d.id).FirstOrDefault();
                return RedirectToAction("Details", new { id = ID });// detail de duyet
            }
            catch
            {
                return View("Error");
            }
        }

    }
}
