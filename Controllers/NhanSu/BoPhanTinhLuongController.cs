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

namespace BatDongSan.Controllers.NhanSu
{
    public class BoPhanTinhLuongController : ApplicationController
    {
        //
        // GET: /BoPhanTinhLuong/
        private LinqDanhMucDataContext linqDM = new LinqDanhMucDataContext();
        LinqNhanSuDataContext linqNS = new LinqNhanSuDataContext();
        private List<tbl_NS_BoPhanTinhLuong> boPhanTinhLuongs;
        private tbl_NS_BoPhanTinhLuong boPhanTinhLuong;
        List<DanhSachNhanVien> ListNhanVienChon;
        private List<BoPhanTinhLuong> lsModel;
        private bool? permission;
        public const string taskIDSystem = "BoPhanTinhLuong";

        public ActionResult Index(int? page, int? pageSize, string qSearch)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            int currentPageIndex = page.HasValue ? page.Value : 1;
            pageSize = pageSize ?? 30;
            IList<sp_BoPhanTinhLuong_IndexResult> list;
            using (linqNS = new LinqNhanSuDataContext())
            {
                list = linqNS.sp_BoPhanTinhLuong_Index(qSearch).ToList();
                ViewBag.Count = list.Count();
                ViewBag.searchString = qSearch;
                TempData["Params"] = qSearch;
            }
            if (Request.IsAjaxRequest())
            {
                return PartialView("ViewIndex", list.ToPagedList(currentPageIndex, 30));
            }
            return View(list.ToPagedList(currentPageIndex, 30));
        }
        public void BuilBoPhan(string selected)
        {
            var ls = linqNS.tbl_NS_BoPhanTinhLuongs.ToList();
            Dictionary<string, string> dict = new Dictionary<string, string>();
            dict.Add("","[Chọn bộ phận]");
            foreach (var item in ls)
            {
                if (!dict.Keys.Contains(item.maBoPhan))
                {
                    dict.Add(item.maBoPhan, item.tenBoPhan);
                }
            }
            ViewBag.lsBoPhan = new SelectList(dict,"Key","Value",selected);
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

            boPhanTinhLuong = new tbl_NS_BoPhanTinhLuong();
            return View();

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
                boPhanTinhLuongs = new List<tbl_NS_BoPhanTinhLuong>();
                string[] maNhanVien = col.GetValues("maNhanVien");
                for (int i = 0; i < maNhanVien.Count(); i++)
                {
                    boPhanTinhLuong = new tbl_NS_BoPhanTinhLuong();
                    boPhanTinhLuong.maBoPhan = col.Get("maBoPhan");
                    boPhanTinhLuong.tenBoPhan = col.Get("tenBoPhan");
                    boPhanTinhLuong.maNhanVien = col.GetValues("maNhanVien")[i];
                    boPhanTinhLuongs.Add(boPhanTinhLuong);

                }
                linqNS.tbl_NS_BoPhanTinhLuongs.InsertAllOnSubmit(boPhanTinhLuongs);
                linqNS.SubmitChanges();
                
                return RedirectToAction("Edit", new { id = col.Get("maBoPhan") });
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
            boPhanTinhLuongs = linqNS.tbl_NS_BoPhanTinhLuongs.Where(d=>d.maBoPhan==id).ToList();
            lsModel = (from p in boPhanTinhLuongs
                       select new BoPhanTinhLuong
                       {
                           MaBoPhan = p.maBoPhan,
                           TenBoPhan = p.tenBoPhan,
                           MaNhanVien = p.maNhanVien,
                           TenNhanVien = linqNS.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d=>d.maNhanVien==p.maNhanVien).Select(d=>d.hoTen).FirstOrDefault()??string.Empty,
                           TenChucDanh = linqNS.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == p.maNhanVien).Select(d => d.TenChucDanh).FirstOrDefault() ?? string.Empty,
                           TenPhongBan = linqNS.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == p.maNhanVien).Select(d => d.tenPhongBan).FirstOrDefault() ?? string.Empty,
                       }).ToList();

            ViewBag.ListBP = lsModel;
            return View();
        }
        public ActionResult Details(string id)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXemChiTiet);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            boPhanTinhLuongs = linqNS.tbl_NS_BoPhanTinhLuongs.Where(d => d.maBoPhan == id).ToList();
            lsModel = (from p in boPhanTinhLuongs
                       select new BoPhanTinhLuong
                       {
                           MaBoPhan = p.maBoPhan,
                           TenBoPhan = p.tenBoPhan,
                           MaNhanVien = p.maNhanVien,
                           TenNhanVien = linqNS.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == p.maNhanVien).Select(d => d.hoTen).FirstOrDefault() ?? string.Empty,
                           TenChucDanh = linqNS.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == p.maNhanVien).Select(d => d.TenChucDanh).FirstOrDefault() ?? string.Empty,
                           TenPhongBan = linqNS.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == p.maNhanVien).Select(d => d.tenPhongBan).FirstOrDefault() ?? string.Empty,
                       }).ToList();

            ViewBag.ListBP = lsModel;
            return View();
        }
        public ActionResult Delete(string maBP)
        {
            var hasValue = "";
            try
            {
                var ls = linqNS.tbl_NS_BoPhanTinhLuongs.Where(d => d.maBoPhan == maBP).ToList();
                if (ls != null && ls.Count > 0)
                {
                    linqNS.tbl_NS_BoPhanTinhLuongs.DeleteAllOnSubmit(ls);
                    linqNS.SubmitChanges();
                }
            }
            catch
            {
                hasValue = "error";
            }
            return Json(hasValue);
        }
        [HttpPost]
        public ActionResult Edit(FormCollection col)
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
                var ls = linqNS.tbl_NS_BoPhanTinhLuongs.Where(d => d.maBoPhan == col.Get("maBoPhan")).ToList();
                if (ls != null && ls.Count > 0)
                {
                    linqNS.tbl_NS_BoPhanTinhLuongs.DeleteAllOnSubmit(ls);
                }
                boPhanTinhLuongs = new List<tbl_NS_BoPhanTinhLuong>();
                string[] maNhanVien = col.GetValues("maNhanVien");
                for (int i = 0; i < maNhanVien.Count(); i++)
                {
                    boPhanTinhLuong = new tbl_NS_BoPhanTinhLuong();
                    boPhanTinhLuong.maBoPhan = col.Get("maBoPhan");
                    boPhanTinhLuong.tenBoPhan = col.Get("tenBoPhan");
                    boPhanTinhLuong.maNhanVien = col.GetValues("maNhanVien")[i];
                    boPhanTinhLuongs.Add(boPhanTinhLuong);

                }
                linqNS.tbl_NS_BoPhanTinhLuongs.InsertAllOnSubmit(boPhanTinhLuongs);
                linqNS.SubmitChanges();

                return RedirectToAction("Edit", new { id = col.Get("maBoPhan") });
            }
            catch
            {
                return View("error");
            }

        }
        public ActionResult CheckBoPhanTinhLuong(string maBP)
        {
            var hasvalue = string.Empty;
            var boPhan = linqNS.tbl_NS_BoPhanTinhLuongs.Where(d => d.maBoPhan == maBP).FirstOrDefault();
            if (boPhan != null)
            {
                hasvalue = "1";
            }
            return Json(hasvalue);
        }
        public ActionResult ChonNhanVien()
        {
            var buildTree = new StringBuilder();
            var phongBans = linqDM.tbl_DM_PhongBans.ToList();
            buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
            ViewBag.NVPB = buildTree.ToString();
            return PartialView("_ChonNhanVien");
        }
        public ActionResult LoadNhanVien(int? page, string searchString, string maPhongBan)
        {
            var NhanVienDaTonTaiBoPhan = linqNS.tbl_NS_BoPhanTinhLuongs.Select(v => v.maNhanVien.Trim()).ToList();

            IList<sp_PB_DanhSachNhanVienResult> phongBan1s;
            phongBan1s = linqDM.sp_PB_DanhSachNhanVien(searchString, maPhongBan).Where(d => !NhanVienDaTonTaiBoPhan.Contains(d.maNhanVien.Trim())).ToList();
            ViewBag.isGet = "True";
            int currentPageIndex = page.HasValue ? page.Value : 1;
            ViewBag.Count = phongBan1s.Count();
            ViewBag.Search = searchString;
            ViewBag.MaPhongBan = maPhongBan;
            return PartialView("_LoadChonNhanVien", phongBan1s.ToPagedList(currentPageIndex, 10));
        }
        public ActionResult DanhSachNhanVienChon(FormCollection coll, string[] MaNV)
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

                        var ListNhanVien = linqNS.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == maNhanVienDaAdd[i]).FirstOrDefault();
                        DanhSachNhanVien DanhSach = new DanhSachNhanVien();
                        DanhSach.maPhongBan = ListNhanVien.maPhongBan;
                        DanhSach.tenPhongBan = ListNhanVien.tenPhongBan;

                        DanhSach.maChucVu = ListNhanVien.maChucDanh;
                        DanhSach.tenChucVu = ListNhanVien.TenChucDanh;

                        DanhSach.tenChiNhanh = ListNhanVien.tenChiNhanh;

                        DanhSach.maNhanVien = ListNhanVien.maNhanVien;
                        DanhSach.tenNhanVien = ListNhanVien.hoTen;
                        ListNhanVienChon.Add(DanhSach);
                    }
                }

                if (splitStr != null && splitStr.Length > 0)
                {
                    for (int j = 0; j < splitStr.Length; j++)
                    {
                        if (!ListNhanVienChon.Select(d => d.maNhanVien).Contains(splitStr[j]))
                        {
                            var ListNhanVien = linqNS.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == splitStr[j]).FirstOrDefault();
                            DanhSachNhanVien DanhSach = new DanhSachNhanVien();
                            DanhSach.maPhongBan = ListNhanVien.maPhongBan;
                            DanhSach.tenPhongBan = ListNhanVien.tenPhongBan;

                            DanhSach.tenChiNhanh = ListNhanVien.tenChiNhanh;

                            DanhSach.maChucVu = ListNhanVien.maChucDanh;
                            DanhSach.tenChucVu = ListNhanVien.TenChucDanh;

                            DanhSach.maNhanVien = ListNhanVien.maNhanVien;
                            DanhSach.tenNhanVien = ListNhanVien.hoTen;
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
    }
}
