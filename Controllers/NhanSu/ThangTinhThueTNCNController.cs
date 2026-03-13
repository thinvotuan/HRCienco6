using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Models.NhanSu;

namespace BatDongSan.Controllers.NhanSu
{
    public class ThangTinhThueTNCNController : ApplicationController
    {
        private LinqNhanSuDataContext context;
        private IList<tbl_NS_ThanhTinhThueTNCN> thangTinhThues;
        private IList<tbl_NS_ThangTinhThueTNCNChiTiet> chiTiets;
        private ThangTinhThueTNCNModel model;
        private tbl_NS_ThanhTinhThueTNCN thangTinhThue;
        private readonly string MCV = "ThangTinhThueTNCN";
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
            thangTinhThues = context.tbl_NS_ThanhTinhThueTNCNs.ToList();
            return View(thangTinhThues);
        }

        //
        // GET: /ThangTinhThueTNCN/Details/5

        public ActionResult Details(int id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXemChiTiet);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            context = new LinqNhanSuDataContext();
            thangTinhThue = context.tbl_NS_ThanhTinhThueTNCNs.Where(s => s.id == id).FirstOrDefault();
            model = new ThangTinhThueTNCNModel();
            model.chiTiets = thangTinhThue.tbl_NS_ThangTinhThueTNCNChiTiets.ToList();
            model.ghiChu = thangTinhThue.ghiChu;
            model.giamTruCaNhan = thangTinhThue.giamTruCaNhan;
            model.giamTruGiaCanh = thangTinhThue.giamTruGiaCanh;
            model.id = thangTinhThue.id;
            model.moTa = thangTinhThue.moTa;
            model.ngayApDung = thangTinhThue.ngayApDung;
            model.ngayCapNhat = thangTinhThue.ngayCapNhat;
            model.nguoiLap = thangTinhThue.nguoiLap;
            model.tinhTrang = thangTinhThue.tinhTrang;
            return View(model);
        }

        //
        // GET: /ThangTinhThueTNCN/Create

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            return View();
        }

        //
        // POST: /ThangTinhThueTNCN/Create

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
                thangTinhThue = new tbl_NS_ThanhTinhThueTNCN();
                thangTinhThue.ghiChu = collection["ghiChu"];
                thangTinhThue.giamTruCaNhan = Convert.ToDecimal(collection["giamTruCaNhan"]);
                thangTinhThue.giamTruGiaCanh = Convert.ToDecimal(collection["giamTruGiaCanh"]);
                thangTinhThue.moTa = collection["moTa"];
                thangTinhThue.ngayApDung = DateTime.ParseExact(collection["ngayApDung"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                thangTinhThue.ngayCapNhat = DateTime.Now;
                thangTinhThue.nguoiLap = GetUser().manv;
                context.tbl_NS_ThanhTinhThueTNCNs.InsertOnSubmit(thangTinhThue);
                context.SubmitChanges();
                InsertChiTiet(collection);
                return RedirectToAction("Index");
            }
            catch
            {
                return View(thangTinhThue);
            }
        }

        //
        // GET: /ThangTinhThueTNCN/Edit/5

        public ActionResult Edit(int id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            context = new LinqNhanSuDataContext();
            thangTinhThue = context.tbl_NS_ThanhTinhThueTNCNs.Where(s => s.id == id).FirstOrDefault();
            model = new ThangTinhThueTNCNModel();
            model.chiTiets = thangTinhThue.tbl_NS_ThangTinhThueTNCNChiTiets.ToList();
            model.ghiChu = thangTinhThue.ghiChu;
            model.giamTruCaNhan = thangTinhThue.giamTruCaNhan;
            model.giamTruGiaCanh = thangTinhThue.giamTruGiaCanh;
            model.id = thangTinhThue.id;
            model.moTa = thangTinhThue.moTa;
            model.ngayApDung = thangTinhThue.ngayApDung;
            model.ngayCapNhat = thangTinhThue.ngayCapNhat;
            model.nguoiLap = thangTinhThue.nguoiLap;
            model.tinhTrang = thangTinhThue.tinhTrang;
            return View(model);
        }

        //
        // POST: /ThangTinhThueTNCN/Edit/5

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
                context = new LinqNhanSuDataContext();
                thangTinhThue = context.tbl_NS_ThanhTinhThueTNCNs.Where(s => s.id == id).FirstOrDefault();
                thangTinhThue.ghiChu = collection["ghiChu"];
                thangTinhThue.giamTruCaNhan = Convert.ToDecimal(collection["giamTruCaNhan"]);
                thangTinhThue.giamTruGiaCanh = Convert.ToDecimal(collection["giamTruGiaCanh"]);
                thangTinhThue.moTa = collection["moTa"];
                thangTinhThue.ngayApDung = DateTime.ParseExact(collection["ngayApDung"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                thangTinhThue.ngayCapNhat = DateTime.Now;                                
                context.SubmitChanges();
                InsertChiTiet(collection);
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // POST: /ThangTinhThueTNCN/Delete/5

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
                context = new LinqNhanSuDataContext();
                thangTinhThue = context.tbl_NS_ThanhTinhThueTNCNs.Where(s => s.id == id).FirstOrDefault();
                context.tbl_NS_ThanhTinhThueTNCNs.DeleteOnSubmit(thangTinhThue);
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        public bool InsertChiTiet(FormCollection colllection)
        {
            try
            {
                chiTiets = new List<tbl_NS_ThangTinhThueTNCNChiTiet>();
                string[] noiDungs = colllection.GetValues("noiDung");
                if (noiDungs != null)
                {
                    var list = context.tbl_NS_ThangTinhThueTNCNChiTiets.Where(s => s.idThangTinhThue == thangTinhThue.id);
                    context.tbl_NS_ThangTinhThueTNCNChiTiets.DeleteAllOnSubmit(list);
                    for (int i = 1; i < noiDungs.Length; i++)
                    {
                        tbl_NS_ThangTinhThueTNCNChiTiet chiTiet = new tbl_NS_ThangTinhThueTNCNChiTiet();
                        chiTiet.idThangTinhThue = thangTinhThue.id;
                        chiTiet.mucChiuThueDen = Convert.ToDecimal(colllection.GetValues("mucChiuThueDen")[i]);
                        chiTiet.mucChiuThueTu = Convert.ToDecimal(colllection.GetValues("mucChiuThueTu")[i]);
                        chiTiet.ngayCapNhat = DateTime.Now;
                        chiTiet.nguoiLap = GetUser().manv;
                        chiTiet.noiDung = colllection.GetValues("noiDung")[i];
                        chiTiet.thue = Convert.ToDecimal(colllection.GetValues("thue")[i]);
                        chiTiets.Add(chiTiet);
                    }
                    context.tbl_NS_ThangTinhThueTNCNChiTiets.InsertAllOnSubmit(chiTiets);
                    context.SubmitChanges();
                }
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
