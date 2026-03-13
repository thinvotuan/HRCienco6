using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Models.NhanSu;
using BatDongSan.Models.DanhMuc;
using BatDongSan.Helper.Common;

namespace BatDongSan.Controllers.NhanSu
{
    public class BangPhanTachMucLuongController : ApplicationController
    {
        private LinqNhanSuDataContext context;
        private LinqDanhMucDataContext contextDanhMuc;
        private IList<tbl_NS_BangPhanTachMucLuong> bangPhanTachs;
        private IList<PhuCapModel> phuCaps;
        private BangPhanTachLuongModel model;
        private tbl_NS_BangPhanTachMucLuong bangPhanTach;
        private readonly string MCV = "BangPhanTachMucLuong";
        private bool? permission;

        public ActionResult Index()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            context = new LinqNhanSuDataContext();
            bangPhanTachs = context.tbl_NS_BangPhanTachMucLuongs.ToList();
            return View(bangPhanTachs);
        }

        //
        // GET: /BangPhanTachMucLuong/Details/5

        public ActionResult Details(int id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXemChiTiet);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            return View();
        }

        //
        // GET: /BangPhanTachMucLuong/Create

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            return PartialView("Create");
        }

        //
        // POST: /BangPhanTachMucLuong/Create

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
                context = new LinqNhanSuDataContext();
                bangPhanTach = new tbl_NS_BangPhanTachMucLuong();
                bangPhanTach.ghiChu = collection["ghiChu"];
                bangPhanTach.luongCoBan = Convert.ToDecimal(collection["luongCoBan"]);
                bangPhanTach.mucLuongTu = Convert.ToDecimal(collection["mucLuongTu"]);
                bangPhanTach.mucLuongDen = Convert.ToDecimal(collection["mucLuongDen"]);
                bangPhanTach.maMucLuong = collection["maMucLuong"];
                bangPhanTach.tenMucLuong = collection["tenMucLuong"];
                context.tbl_NS_BangPhanTachMucLuongs.InsertOnSubmit(bangPhanTach);
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /BangPhanTachMucLuong/Edit/5

        public ActionResult Edit(string id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            context = new LinqNhanSuDataContext();
            bangPhanTach = context.tbl_NS_BangPhanTachMucLuongs.Where(s => s.maMucLuong == id).FirstOrDefault();
            return PartialView("Edit", bangPhanTach);
        }

        //
        // POST: /BangPhanTachMucLuong/Edit/5

        [HttpPost]
        public ActionResult Edit(string id, FormCollection collection)
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
                context = new LinqNhanSuDataContext();
                bangPhanTach = context.tbl_NS_BangPhanTachMucLuongs.Where(s => s.maMucLuong == id).FirstOrDefault();
                bangPhanTach.ghiChu = collection["ghiChu"];
                bangPhanTach.luongCoBan = Convert.ToDecimal(collection["luongCoBan"]);
                bangPhanTach.mucLuongTu = Convert.ToDecimal(collection["mucLuongTu"]);
                bangPhanTach.mucLuongDen = Convert.ToDecimal(collection["mucLuongDen"]);
                bangPhanTach.tenMucLuong = collection["tenMucLuong"];
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        public ActionResult PhanBoChiTiet(string id) {
            using (context = new LinqNhanSuDataContext())
            {
                model = new BangPhanTachLuongModel();
                bangPhanTach = context.tbl_NS_BangPhanTachMucLuongs.Where(s => s.maMucLuong == id).FirstOrDefault();
                model.tenMucLuong = bangPhanTach.tenMucLuong;
                model.maMucLuong = bangPhanTach.maMucLuong;
                model.luongCoBan = bangPhanTach.luongCoBan;
                model.mucLuongDen = bangPhanTach.mucLuongDen;
                model.mucLuongTu = bangPhanTach.mucLuongTu;
                model.phuCapKhongTheoTiLe = bangPhanTach.mucLuongTu - bangPhanTach.luongCoBan;
                model.phuCapTheoTiLe = bangPhanTach.mucLuongDen - bangPhanTach.luongCoBan - model.phuCapKhongTheoTiLe;
                model.phuCapKhongTheoTiLe = model.phuCapKhongTheoTiLe < 0 ? 0 : model.phuCapKhongTheoTiLe;

                model.chiTiets = (from ct in context.tbl_NS_BangPhanTachMucLuongChiTiets
                                  join b in context.tbl_NS_BangPhanTachMucLuongs on ct.maMucLuong equals b.maMucLuong
                                  join pc in context.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_PhuCap>() on ct.maPhuCap equals pc.maPhuCap
                                  join l in context.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_LoaiPhuCap>() on pc.loaiPhuCap equals l.id
                                  where ct.maMucLuong == id
                                  select new BangPhanTachLuongChiTietModel
                                  {
                                      ghiChu = ct.ghiChu,
                                      id = ct.id,
                                      loaiTyLe = ct.loaiTyLe,
                                      idLoaiPhuCap = pc.loaiPhuCap,
                                      tenLoaiPhuCap = l.tenLoaiPhuCap,
                                      maMucLuong = ct.maMucLuong,
                                      maPhuCap = ct.maPhuCap,
                                      salaryTemplate = ct.salaryTemplate,
                                      tenPhuCap = pc.tenPhuCap,
                                      tenMucLuong = b.tenMucLuong,
                                      tyLe = ct.tyLe
                                  }).ToList();
            }
            return PartialView("PartialPhuCapChiTiet", model);
        }        

        //
        // POST: /BangPhanTachMucLuong/Delete/5

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
                context = new LinqNhanSuDataContext();
                bangPhanTach = context.tbl_NS_BangPhanTachMucLuongs.Where(s => s.maMucLuong == id).FirstOrDefault();
                context.tbl_NS_BangPhanTachMucLuongs.DeleteOnSubmit(bangPhanTach);
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        public ActionResult GetPhuCapList() {
            using (contextDanhMuc = new LinqDanhMucDataContext())
            {
                phuCaps = (from pc in contextDanhMuc.tbl_DM_PhuCaps
                           join l in contextDanhMuc.tbl_DM_LoaiPhuCaps on pc.loaiPhuCap equals l.id
                           join pb in contextDanhMuc.tbl_DM_PhongBans on pc.maPhongBan equals pb.maPhongBan
                           select new PhuCapModel
                           {
                               coTinhThueTNCN = pc.coTinhThueTNCN,
                               ghiChu = pc.ghiChu,
                               loaiPhuCap = l.id,
                               maPhongBan = pb.maPhongBan,
                               maPhuCap = pc.maPhuCap,
                               ngayCapNhat = pc.ngayCapNhat,
                               ngayLap = pc.ngayLap,
                               nguoiLap = pc.nguoiLap,
                               soTien = pc.soTien,
                               tenLoaiPhuCap = l.tenLoaiPhuCap,
                               tenPhongBan = pb.tenPhongBan,
                               tenPhuCap = pc.tenPhuCap
                           }).ToList();
            }
            return PartialView("PartialPhuCapDanhSach", phuCaps);
        }


        public ActionResult SavePhuCap(string id, FormCollection collection)
        {
            try
            {
                context = new LinqNhanSuDataContext();
                var list = context.tbl_NS_BangPhanTachMucLuongChiTiets.Where(s => s.maMucLuong == id);
                context.tbl_NS_BangPhanTachMucLuongChiTiets.DeleteAllOnSubmit(list);

                string[] maPhuCaps = collection.GetValues("maPhuCap");
                IList<tbl_NS_BangPhanTachMucLuongChiTiet> bangPhanTachCTs = new List<tbl_NS_BangPhanTachMucLuongChiTiet>();
                if (maPhuCaps != null)
                {
                    for (int i = 0; i < maPhuCaps.Length; i++)
                    {
                        tbl_NS_BangPhanTachMucLuongChiTiet chiTiet = new tbl_NS_BangPhanTachMucLuongChiTiet();
                        chiTiet.ghiChu = collection.GetValues("ghiChu")[i];
                        chiTiet.loaiTyLe = collection.GetValues("loaiTyLe")[i];
                        chiTiet.maMucLuong = id;
                        chiTiet.maPhuCap = collection.GetValues("maPhuCap")[i];
                        chiTiet.salaryTemplate = collection.GetValues("salaryTemplate")[i];
                        chiTiet.tyLe = Convert.ToDecimal(collection.GetValues("tyLe")[i]);                        
                        bangPhanTachCTs.Add(chiTiet);
                    }                    
                    context.tbl_NS_BangPhanTachMucLuongChiTiets.InsertAllOnSubmit(bangPhanTachCTs);
                    context.SubmitChanges();
                }                
                return Json(String.Empty);
            }
            catch
            {
                return View();
            }
        }
    }
}
