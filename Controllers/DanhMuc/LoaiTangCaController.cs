using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Models.DanhMuc;
using BatDongSan.Models.NhanSu;

namespace BatDongSan.Controllers.DanhMuc
{
    public class LoaiTangCaController : ApplicationController
    {
        private LinqDanhMucDataContext context;
        private IList<LoaiTangCaModel> loaiTangCas;
        private tbl_DM_LoaiTangCa loaiTangCa;
        private readonly string MCV = "LoaiTangCa";
        private bool? permission;
        private string tenNguoiLap(string maNhanVien)
        {
            try
            {
                LinqNhanSuDataContext lq = new LinqNhanSuDataContext();
                string abc = lq.tbl_NS_NhanViens.Where(x => x.maNhanVien == maNhanVien).Select(x => x.ho + " " + x.ten).FirstOrDefault();
                return abc != null ? abc : maNhanVien;
            }
            catch
            {
                return maNhanVien;
            }
        }
        public ActionResult Index()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            using (context = new LinqDanhMucDataContext())
            {
                loaiTangCas = context.tbl_DM_LoaiTangCas.Select(t => new LoaiTangCaModel
                {
                        heSoTangCa = t.heSoTangCa ,
                        ngayLap = t.ngayLap ,
                        nguoiLap = tenNguoiLap(t.nguoiLap),
                        id =t.id ,
                        loaiTangCa = t.loaiTangCa 
                }).ToList();
            }
            return View(loaiTangCas);
        }
       
        //
        // GET: /NgayLe/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /NgayLe/Create

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            context = new LinqDanhMucDataContext();
            loaiTangCa = new tbl_DM_LoaiTangCa();
            return PartialView("Create", loaiTangCa);
        }

        //
        // POST: /NgayLe/Create

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
                using (context = new LinqDanhMucDataContext())
                {
                    loaiTangCa = new tbl_DM_LoaiTangCa();
                    loaiTangCa.loaiTangCa = collection["loaiTangCa"];
                    loaiTangCa.nguoiLap = GetUser().manv;
                    loaiTangCa.ngayLap = DateTime.Now;
                    loaiTangCa.heSoTangCa = Convert.ToDouble(collection["heSoTangCa"]);
                    context.tbl_DM_LoaiTangCas.InsertOnSubmit(loaiTangCa);
                    context.SubmitChanges();
                }
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /NgayLe/Edit/5

        public ActionResult Edit(int id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            context = new LinqDanhMucDataContext();
            loaiTangCa = context.tbl_DM_LoaiTangCas.Where(s => s.id == id).FirstOrDefault();

            return PartialView("Edit", loaiTangCa);            
        }

        //
        // POST: /NgayLe/Edit/5

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
                using (context = new LinqDanhMucDataContext())
                {
                    loaiTangCa = context.tbl_DM_LoaiTangCas.Where(s => s.id == id).FirstOrDefault();
                    loaiTangCa.loaiTangCa = collection["loaiTangCa"];
                    loaiTangCa.heSoTangCa =Convert.ToDouble ( collection["heSoTangCa"]);
                    loaiTangCa.ngayLap = DateTime.Now;
                    context.SubmitChanges();
                }

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // POST: /NgayLe/Delete/5

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
                using (context = new LinqDanhMucDataContext())
                {
                    loaiTangCa = context.tbl_DM_LoaiTangCas.Where(s => s.id == id).FirstOrDefault();
                    context.tbl_DM_LoaiTangCas.DeleteOnSubmit(loaiTangCa);
                    context.SubmitChanges();
                }

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
    }
}
