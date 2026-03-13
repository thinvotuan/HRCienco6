using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Models.DanhMuc;

namespace BatDongSan.Controllers.DanhMuc
{
    public class PhuCapController : ApplicationController
    {
        private LinqDanhMucDataContext context;
        private IList<PhuCapModel> phuCaps;
        private tbl_DM_PhuCap phuCap;
        private readonly string MCV = "PhuCap";
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

            using (context = new LinqDanhMucDataContext())
            {
                phuCaps = (from pc in context.tbl_DM_PhuCaps
                           join l in context.tbl_DM_LoaiPhuCaps on pc.loaiPhuCap equals l.id
                           join pb in context.tbl_DM_PhongBans on pc.maPhongBan equals pb.maPhongBan
                           select new PhuCapModel
                           {
                               coTinhThueTNCN = pc.coTinhThueTNCN,
                               coTinhBaoHiem = pc.coTinhBaoHiem,
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
            return View(phuCaps);
        }

        //
        // GET: /PhuCap/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /PhuCap/Create

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

            var phongBans = context.tbl_DM_PhongBans;
            ViewBag.PhongBans = new SelectList(phongBans, "maPhongBan", "tenPhongBan");

            var loaiPhuCaps = context.tbl_DM_LoaiPhuCaps;
            ViewBag.LoaiPhuCaps = new SelectList(loaiPhuCaps, "id", "tenLoaiPhuCap");

            phuCap = new tbl_DM_PhuCap();
            return PartialView("Create", phuCap);
        }

        //
        // POST: /PhuCap/Create

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
                    phuCap = new tbl_DM_PhuCap();
                    phuCap.coTinhThueTNCN = collection["coTinhThueTNCN"].Contains("true");
                    phuCap.coTinhBaoHiem = collection["coTinhBaoHiem"].Contains("true");
                    phuCap.ghiChu = collection["ghiChu"];
                    phuCap.loaiPhuCap = Convert.ToInt32(collection["loaiPhuCap"]);
                    phuCap.maPhuCap = collection["maPhuCap"];
                    phuCap.maPhongBan = collection["maPhongBan"];
                    phuCap.ngayLap = DateTime.Now;
                    phuCap.nguoiLap = GetUser().manv;
                    phuCap.soTien = Convert.ToDecimal(collection["soTien"]);
                    phuCap.tenPhuCap = collection["tenPhuCap"];
                    var obj = context.tbl_DM_PhuCaps.Where(s => s.maPhuCap == phuCap.maPhuCap).FirstOrDefault();
                    if (obj != null)
                    {
                        TempData["TonTai"] = "Lưu không thành công - mã phụ cấp đã tồn tại";
                        return RedirectToAction("Index");
                    }
                    context.tbl_DM_PhuCaps.InsertOnSubmit(phuCap);
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
        // GET: /PhuCap/Edit/5

        public ActionResult Edit(string id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            context = new LinqDanhMucDataContext();

            var phongBans = context.tbl_DM_PhongBans;
            ViewBag.PhongBans = new SelectList(phongBans, "maPhongBan", "tenPhongBan");

            var loaiPhuCaps = context.tbl_DM_LoaiPhuCaps;
            ViewBag.LoaiPhuCaps = new SelectList(loaiPhuCaps, "id", "tenLoaiPhuCap");

            phuCap = context.tbl_DM_PhuCaps.Where(s => s.maPhuCap == id).FirstOrDefault();

            return PartialView("Edit", phuCap);
        }

        //
        // POST: /PhuCap/Edit/5

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
                using (context = new LinqDanhMucDataContext())
                {
                    phuCap = context.tbl_DM_PhuCaps.Where(s => s.maPhuCap == id).FirstOrDefault();
                    phuCap.coTinhThueTNCN = collection["coTinhThueTNCN"].Contains("true");
                    phuCap.coTinhBaoHiem = collection["coTinhBaoHiem"].Contains("true");
                    phuCap.ghiChu = collection["ghiChu"];
                    phuCap.loaiPhuCap = Convert.ToInt32(collection["loaiPhuCap"]);
                    //phuCap.maPhuCap = collection["maPhuCap"];
                    phuCap.maPhongBan = collection["maPhongBan"];
                    phuCap.ngayCapNhat = DateTime.Now;
                    phuCap.soTien = Convert.ToDecimal(collection["soTien"]);
                    phuCap.tenPhuCap = collection["tenPhuCap"];
                    var obj = context.tbl_DM_PhuCaps.Where(s => s.maPhuCap == phuCap.maPhuCap).FirstOrDefault();
                    if (obj != null && obj.maPhuCap != id)
                    {
                        TempData["TonTai"] = "Mã đã tồn tại";
                        return PartialView("Edit", phuCap);
                    }
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
        // POST: /PhuCap/Delete/5

        [HttpPost]
        public ActionResult Delete(string id, FormCollection collection)
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
                    phuCap = context.tbl_DM_PhuCaps.Where(s => s.maPhuCap == id).FirstOrDefault();
                    context.tbl_DM_PhuCaps.DeleteOnSubmit(phuCap);
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
