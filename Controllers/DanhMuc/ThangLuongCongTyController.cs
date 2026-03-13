using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Models.DanhMuc;

namespace BatDongSan.Controllers.DanhMuc
{
    public class ThangLuongCongTyController : ApplicationController
    {
        private LinqDanhMucDataContext context = new LinqDanhMucDataContext();
        private IList<tbl_DM_ThangLuongCongTy> thangLuongs;
        private ThangLuongCongTyModel model;
        private tbl_DM_ThangLuongCongTy thangLuong;
        private readonly string MCV = "ThangLuongCongTy";
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

            thangLuongs = context.tbl_DM_ThangLuongCongTies.OrderBy(s => s.bac).ToList();
            return View(thangLuongs);
        }

        //
        // GET: /ThangLuongCongTy/Details/5

        public ActionResult Details(int id)
        {

            return View();
        }

        //
        // GET: /ThangLuongCongTy/Create

        public ActionResult Create()
        {
            model = new ThangLuongCongTyModel();
            return PartialView("Create", model);
        }

        //
        // POST: /ThangLuongCongTy/Create

        [HttpPost]
        public ActionResult Create(ThangLuongCongTyModel viewModel)
        {
            try
            {
                thangLuong = new tbl_DM_ThangLuongCongTy();
                thangLuong.bac = Convert.ToInt32(viewModel.bac);
                thangLuong.capBacChucVu = Convert.ToInt32(viewModel.capBacChucVu);
                SetDataFromView(viewModel);

                var check = context.tbl_DM_ThangLuongCongTies.Where(s => s.capBacChucVu == thangLuong.capBacChucVu && s.bac == thangLuong.bac).FirstOrDefault();
                if (check != null)
                {
                    TempData["TonTai"] = "Số bậc lương và cấp bậc chức vụ tương ứng đã tồn tại. Vui lòng nhập số khác";
                    return RedirectToAction("Index");
                }
                context.tbl_DM_ThangLuongCongTies.InsertOnSubmit(thangLuong);
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /ThangLuongCongTy/Edit/5

        public ActionResult Edit(int bac, int bacChucVu)
        {
            model = (from t in context.tbl_DM_ThangLuongCongTies
                     where t.bac == bac && t.capBacChucVu == bacChucVu
                     select new ThangLuongCongTyModel
                     {
                         bac = t.bac,
                         capBacChucVu = t.capBacChucVu,
                         luongCanBanChinhThuc = t.luongCanBanChinhThuc,
                         luongCanBanKhoiDiem = t.luongCanBanKhoiDiem,
                         luongCanBanThucThu = t.luongCanBanThucThu,
                         luongThanhTichChinhThuc = t.luongThanhTichChinhThuc,
                         luongThanhTichKhoiDiem = t.luongThanhTichKhoiDiem,
                         luongThanhTichThucThu = t.luongThanhTichThucThu,
                         tongLuongChinhThuc = t.tongLuongChinhThuc,
                         tongLuongKhoiDiem = t.tongLuongKhoiDiem,
                         tongLuongThucThu = t.tongLuongThucThu,
                     }).FirstOrDefault();
            return PartialView("Edit", model);
        }

        //
        // POST: /ThangLuongCongTy/Edit/5

        [HttpPost]
        public ActionResult Edit(ThangLuongCongTyModel viewModel)
        {
            try
            {
                thangLuong = context.tbl_DM_ThangLuongCongTies.Where(s => s.capBacChucVu == viewModel.capBacChucVu && s.bac == viewModel.bac).FirstOrDefault();
                SetDataFromView(viewModel);
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // POST: /ThangLuongCongTy/Delete/5

        [HttpPost]
        public ActionResult Delete(int bac, int bacChucVu)
        {
            try
            {
                // TODO: Add delete logic here
                thangLuong = context.tbl_DM_ThangLuongCongTies.Where(s => s.capBacChucVu == bacChucVu && s.bac == bac).FirstOrDefault();
                context.tbl_DM_ThangLuongCongTies.DeleteOnSubmit(thangLuong);
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        public void SetDataFromView(ThangLuongCongTyModel viewModel)
        {
            thangLuong.luongCanBanChinhThuc = viewModel.luongCanBanChinhThuc;
            thangLuong.luongCanBanKhoiDiem = viewModel.luongCanBanKhoiDiem;
            thangLuong.luongCanBanThucThu = viewModel.luongCanBanThucThu;
            thangLuong.luongThanhTichChinhThuc = viewModel.luongThanhTichChinhThuc;
            thangLuong.luongThanhTichKhoiDiem = viewModel.luongThanhTichKhoiDiem;
            thangLuong.luongThanhTichThucThu = viewModel.luongThanhTichThucThu;
            thangLuong.tongLuongChinhThuc = viewModel.tongLuongChinhThuc;
            thangLuong.tongLuongKhoiDiem = viewModel.tongLuongKhoiDiem;
            thangLuong.tongLuongThucThu = viewModel.tongLuongThucThu;
        }
    }
}
