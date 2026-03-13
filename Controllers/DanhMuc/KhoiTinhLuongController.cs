using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Models.DanhMuc;
using BatDongSan.Controllers;
using BatDongSan.Helper.Common;
using BatDongSan.Models.HeThong;

namespace BatDongSan.Controllers.DanhMuc
{

    public class KhoiTinhLuongController : ApplicationController
    {
        private LinqDanhMucDataContext context = new LinqDanhMucDataContext();
        private IList<tbl_NS_KhoiTinhLuong> khoiTinhLuongs;
        private tbl_NS_KhoiTinhLuong khoiTinhLuong;
        private tbl_NS_PhanCa phanCa;
        private tbl_NS_PhanCaChiTiet phanCaChiTiet;
        private bool? permission;
        public const string taskIDSystem = "KhoiTinhLuong";
        //
        // GET: /KhoiTinhLuong/

        public ActionResult Index()
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            khoiTinhLuongs = context.tbl_NS_KhoiTinhLuongs.OrderByDescending(d => d.maKhoiTinhLuong).ToList();
            return View(khoiTinhLuongs);
        }
        public ActionResult ViewIndex()
        {
            khoiTinhLuongs = context.tbl_NS_KhoiTinhLuongs.OrderByDescending(d => d.maKhoiTinhLuong).ToList();
            return PartialView("_ViewIndex", khoiTinhLuongs);
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
            ViewBag.TitleHead = "Thêm";
            khoiTinhLuong = new tbl_NS_KhoiTinhLuong();

            return PartialView("_CreateEdit", khoiTinhLuong);
        }

        [HttpPost]
        public ActionResult Create(FormCollection coll)
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
                khoiTinhLuong = new tbl_NS_KhoiTinhLuong();
                khoiTinhLuong.tenKhoiTinhLuong = coll.Get("tenKhoiTinhLuong");
                khoiTinhLuong.ghiChu = coll.Get("ghiChu");
                khoiTinhLuong.userName = GetUser().userName;
                khoiTinhLuong.ngayCapNhat = DateTime.Now;
                context.tbl_NS_KhoiTinhLuongs.InsertOnSubmit(khoiTinhLuong);
                context.SubmitChanges();
                khoiTinhLuongs = context.tbl_NS_KhoiTinhLuongs.OrderByDescending(d => d.maKhoiTinhLuong).ToList();
                return PartialView("_ViewIndex", khoiTinhLuongs);
            }
            catch
            {
                return View();
            }
        }
        public ActionResult Edit(int id)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            ViewBag.TitleHead = "Cập nhật";
            khoiTinhLuong = context.tbl_NS_KhoiTinhLuongs.Where(d => d.maKhoiTinhLuong == id).FirstOrDefault();
            return PartialView("_CreateEdit", khoiTinhLuong);
        }
        [HttpPost]
        public ActionResult Edit(FormCollection coll, int id)
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
                khoiTinhLuong = context.tbl_NS_KhoiTinhLuongs.Where(d => d.maKhoiTinhLuong == id).FirstOrDefault();
                khoiTinhLuong.tenKhoiTinhLuong = coll.Get("tenKhoiTinhLuong");
                khoiTinhLuong.ghiChu = coll.Get("ghiChu");
                khoiTinhLuong.ngayCapNhat = DateTime.Now;
                context.SubmitChanges();
                khoiTinhLuongs = context.tbl_NS_KhoiTinhLuongs.OrderByDescending(d => d.maKhoiTinhLuong).ToList();
                return PartialView("_ViewIndex", khoiTinhLuongs);
            }
            catch
            {
                return View();
            }
        }
        [HttpPost]
        public ActionResult Delete(int id)
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
                khoiTinhLuong = context.tbl_NS_KhoiTinhLuongs.Where(d => d.maKhoiTinhLuong == id).FirstOrDefault();
                context.tbl_NS_KhoiTinhLuongs.DeleteOnSubmit(khoiTinhLuong);

                var ListPhanCa = context.tbl_NS_PhanCas.Where(d => d.maKhoiTinhLuong == id).ToList();
                if (ListPhanCa != null && ListPhanCa.Count > 0)
                {
                    foreach (var ct in ListPhanCa)
                    {
                        var chiTietPC = context.tbl_NS_PhanCaChiTiets.Where(d => d.maPhanCa == ct.maPhanCa).ToList();
                        if (chiTietPC != null && chiTietPC.Count > 0)
                        {
                            context.tbl_NS_PhanCaChiTiets.DeleteAllOnSubmit(chiTietPC);
                        }
                    }
                    context.tbl_NS_PhanCas.DeleteAllOnSubmit(ListPhanCa);
                }

                context.SubmitChanges();
                khoiTinhLuongs = context.tbl_NS_KhoiTinhLuongs.OrderByDescending(d => d.maKhoiTinhLuong).ToList();
                return PartialView("_ViewIndex", khoiTinhLuongs);
            }
            catch
            {
                return View();
            }
        }
        [HttpPost]
        public ActionResult DeletePhanCa(int id)
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
                var phanCaCT = context.tbl_NS_PhanCaChiTiets.Where(d => d.maPhanCa == id).ToList();
                if (phanCaCT != null && phanCaCT.Count > 0)
                {
                    context.tbl_NS_PhanCaChiTiets.DeleteAllOnSubmit(phanCaCT);
                }
                phanCa = context.tbl_NS_PhanCas.Where(d => d.maPhanCa == id).FirstOrDefault();
                context.tbl_NS_PhanCas.DeleteOnSubmit(phanCa);
                context.SubmitChanges();
                return Json(string.Empty);
            }
            catch
            {
                return View();
            }
        }

        public ActionResult PhanCa(int id)
        {
            khoiTinhLuong = context.tbl_NS_KhoiTinhLuongs.Where(d => d.maKhoiTinhLuong == id).FirstOrDefault();
            ViewBag.ListPhanCa = context.tbl_NS_PhanCas.Where(d => d.maKhoiTinhLuong == id).ToList();
            
            return View(khoiTinhLuong);
        }

        public ActionResult CreatePhanCa(int id)
        {
            ViewBag.khoiTinhLuong = context.tbl_NS_KhoiTinhLuongs.Where(d => d.maKhoiTinhLuong == id).FirstOrDefault();
            phanCa = new tbl_NS_PhanCa();
            return View(phanCa);
        }

        [HttpPost]
        public ActionResult CreatePhanCa(int id, FormCollection coll)
        {
            try
            {

                phanCa = new tbl_NS_PhanCa();
                phanCa.tenPhanCa = coll.Get("tenPhanCa");
                phanCa.maKhoiTinhLuong = id;
                phanCa.ghiChu = coll.Get("ghiChu");
                try
                {
                    phanCa.tongGioLamViec = Convert.ToDecimal(coll.Get("tongGioLamViec"));
                }
                catch
                {
                    phanCa.tongGioLamViec = 0;
                }
                phanCa.tinhNgayGioCong = coll.Get("isSalaryByMont").Contains("true");
                phanCa.thuHai = coll.Get("monday").Contains("true");
                phanCa.thuBa = coll.Get("tuesday").Contains("true");
                phanCa.thuTu = coll.Get("wednesday").Contains("true");
                phanCa.thuNam = coll.Get("thursday").Contains("true");
                phanCa.thuSau = coll.Get("friday").Contains("true");
                phanCa.thuBay = coll.Get("sathurday").Contains("true");
                phanCa.chuNhat = coll.Get("sunday").Contains("true");
                phanCa.userName = GetUser().userName;
                phanCa.ngayCapNhat = DateTime.Now;

                context.tbl_NS_PhanCas.InsertOnSubmit(phanCa);
                context.SubmitChanges();
                return RedirectToAction("EditPhanCa", new { id = phanCa.maPhanCa });
            }
            catch (Exception ex)
            {
                ViewData["Message"] = ex.Message;
                return View("Error");
            }
        }
        public ActionResult EditPhanCa(int id)
        {
            phanCa = context.tbl_NS_PhanCas.Where(d => d.maPhanCa == id).FirstOrDefault();
            ViewBag.khoiTinhLuong = context.tbl_NS_KhoiTinhLuongs.Where(d => d.maKhoiTinhLuong == phanCa.maKhoiTinhLuong).FirstOrDefault();
            return View(phanCa);

        }
        [HttpPost]
        public ActionResult EditPhanCa(int id, FormCollection coll)
        {
            try
            {

                phanCa = context.tbl_NS_PhanCas.Where(d => d.maPhanCa == id).FirstOrDefault();
                phanCa.tenPhanCa = coll.Get("tenPhanCa");
                phanCa.ghiChu = coll.Get("ghiChu");
                try
                {
                    phanCa.tongGioLamViec = Convert.ToDecimal(coll.Get("tongGioLamViec"));
                }
                catch
                {
                    phanCa.tongGioLamViec = 0;
                }
                phanCa.tinhNgayGioCong = coll.Get("isSalaryByMont").Contains("true");
                phanCa.thuHai = coll.Get("monday").Contains("true");
                phanCa.thuBa = coll.Get("tuesday").Contains("true");
                phanCa.thuTu = coll.Get("wednesday").Contains("true");
                phanCa.thuNam = coll.Get("thursday").Contains("true");
                phanCa.thuSau = coll.Get("friday").Contains("true");
                phanCa.thuBay = coll.Get("sathurday").Contains("true");
                phanCa.chuNhat = coll.Get("sunday").Contains("true");
                phanCa.userName = GetUser().userName;
                phanCa.ngayCapNhat = DateTime.Now;
                context.SubmitChanges();
                return RedirectToAction("EditPhanCa", new { id = phanCa.maPhanCa });
            }
            catch (Exception ex)
            {
                ViewData["Message"] = ex.Message;
                return View("Error");
            }
        }
        public ActionResult DetailPhanCa(int id,string weekdayId, string shiftDetailId)
        {
            phanCa = context.tbl_NS_PhanCas.Where(d => d.maPhanCa == id).FirstOrDefault();
            ViewBag.khoiTinhLuong = context.tbl_NS_KhoiTinhLuongs.Where(d => d.maKhoiTinhLuong == phanCa.maKhoiTinhLuong).FirstOrDefault();
            //ViewBag.ChiTietPhanCa = context.tbl_PhanCaChiTiets.Where(d => d.shiftId == id).FirstOrDefault();

            var ChiTietPhanCa = (from p in context.tbl_NS_PhanCaChiTiets.Where(d => d.maPhanCaChiTiet == id)
                                 select new DMPhanCaChiTiet
                                {
                                    soGio = p.soGio,
                                    veSomChoPhep = p.veSomChoPhep ?? 0,
                                    diTreChoPhep =p.diTreChoPhep??0,
                                    nghiGiuaCa = p.nghiGiuaCa ??false,
                                    ghiChu =p.ghiChu,
                                    gioKetThuc = p.gioKetThuc,
                                    heSo = p.heSo,
                                    gioBatDau = p.gioBatDau,
                                    thoiGianNghiGiuaCa = p.thoiGianNghiGiuaCa ?? 0,
                                    ketThucNghiGiuaCa = p.ketThucNghiGiuaCa,
                                    batDauNghiGiuaCa = p.batDauNghiGiuaCa,
                                    
                                }).FirstOrDefault();
            ViewBag.ChiTietPhanCa = ChiTietPhanCa;
            ViewData["weekdayId"] = String.IsNullOrEmpty(weekdayId) ? String.Empty : weekdayId;
            ViewData["shiftDetailId"] = String.IsNullOrEmpty(shiftDetailId) ? String.Empty : shiftDetailId;
            return View(phanCa);
        }
       
        public ActionResult GetShiftDetail(int dayOfWeek, int shiftId)
        {
            phanCaChiTiet = GetShiftDetailByIdAndDayOfWeek(dayOfWeek, shiftId);
            if (phanCaChiTiet != null)
            {
                tbl_NS_PhanCaChiTiet json = new tbl_NS_PhanCaChiTiet();
                json.maThu = phanCaChiTiet.maThu;
                json.maPhanCa = phanCaChiTiet.maPhanCa;
                json.maPhanCaChiTiet = phanCaChiTiet.maPhanCaChiTiet;
                json.soGio = phanCaChiTiet.soGio;
                json.hourStartShift = String.Format("{0:HH:mm}", phanCaChiTiet.gioBatDau);
                json.hourEndShift = String.Format("{0:HH:mm}", phanCaChiTiet.gioKetThuc);
                json.heSo = phanCaChiTiet.heSo;
                json.diTreChoPhep = phanCaChiTiet.diTreChoPhep;
                json.veSomChoPhep = phanCaChiTiet.veSomChoPhep;
                json.nghiGiuaCa = phanCaChiTiet.nghiGiuaCa; 
                json.hourStartBetweenShift = String.Format("{0:HH:mm}", phanCaChiTiet.batDauNghiGiuaCa);
                json.hourEndBetweenShift = String.Format("{0:HH:mm}", phanCaChiTiet.ketThucNghiGiuaCa);
                json.ketThucNghiGiuaCa = phanCaChiTiet.ketThucNghiGiuaCa;
                json.ghiChu = phanCaChiTiet.ghiChu;

                return Json(json);
            }
            else
                return Json(String.Empty);
        }
        [HttpPost]
        public ActionResult DetailPhanCa(int id, FormCollection colection)
        {
            phanCa = context.tbl_NS_PhanCas.Where(d => d.maPhanCa == id).FirstOrDefault();
            string date = String.Format("{0:dd/MM/yyyy}", DateTime.Now);
            if (String.IsNullOrEmpty(colection.Get("shiftDetailId")))
            {
                phanCaChiTiet = new tbl_NS_PhanCaChiTiet();
                phanCaChiTiet.maPhanCaChiTiet = (short)(GetMaxDetailPhanCa() + 1);
            }
            else
            {
                phanCaChiTiet = context.tbl_NS_PhanCaChiTiets.Where(d => d.maPhanCaChiTiet == Convert.ToInt16(colection.Get("shiftDetailId"))).FirstOrDefault();
            }
            phanCaChiTiet.maThu = Convert.ToByte(colection.Get("weekdayId"));
            phanCaChiTiet.maPhanCa = (short)id;
            phanCaChiTiet.soGio = Convert.ToDecimal(colection.Get("chiTiet.soGio"));
            if (!String.IsNullOrEmpty(colection.Get("chiTiet.diTreChoPhep")))
                phanCaChiTiet.diTreChoPhep = Convert.ToByte(colection.Get("chiTiet.diTreChoPhep"));
            if (!String.IsNullOrEmpty(colection.Get("chiTiet.veSomChoPhep")))
                phanCaChiTiet.veSomChoPhep = Convert.ToByte(colection.Get("chiTiet.veSomChoPhep"));
            phanCaChiTiet.nghiGiuaCa = Convert.ToBoolean(colection.Get("chiTiet.nghiGiuaCa"));
            if (phanCaChiTiet.nghiGiuaCa ?? false)
            {
                phanCaChiTiet.thoiGianNghiGiuaCa = Convert.ToByte(colection.Get("chiTiet.thoiGianNghiGiuaCa"));
                phanCaChiTiet.batDauNghiGiuaCa = Convert.ToDateTime(date + " " + colection.Get("chiTiet.batDauNghiGiuaCa"));
                phanCaChiTiet.ketThucNghiGiuaCa = Convert.ToDateTime(date + " " + colection.Get("chiTiet.ketThucNghiGiuaCa"));
            }
            else
            {
                phanCaChiTiet.thoiGianNghiGiuaCa = null;
                phanCaChiTiet.batDauNghiGiuaCa = null;
                phanCaChiTiet.ketThucNghiGiuaCa = null;
            }
            phanCaChiTiet.ghiChu = colection.Get("chiTiet.ghiChu");
            phanCaChiTiet.gioKetThuc = Convert.ToDateTime(date + " " + colection.Get("chiTiet.gioKetThuc"));
            phanCaChiTiet.heSo = Convert.ToDecimal(colection.Get("chiTiet.heSo"));
            phanCaChiTiet.gioBatDau = Convert.ToDateTime(date + " " + colection.Get("chiTiet.gioBatDau"));

            phanCaChiTiet.userName = GetUser().userName;
            phanCaChiTiet.ngayCapNhat = DateTime.Now;

            if (String.IsNullOrEmpty(colection.Get("shiftDetailId")))
            {
                context.tbl_NS_PhanCaChiTiets.InsertOnSubmit(phanCaChiTiet);
                context.SubmitChanges();
            }
                else
                    context.SubmitChanges();

            if (colection.Get("autoUpdate").Contains("true"))
            {
                AutoUpdateDayOfShift(phanCaChiTiet);
            }


            return RedirectToAction("DetailPhanCa", new { id = phanCa.maPhanCa, weekdayId = colection.Get("weekdayId"), shiftDetailId = phanCaChiTiet.maPhanCaChiTiet });
            
        }

        public void AutoUpdateDayOfShift(tbl_NS_PhanCaChiTiet shiftDetail)
        {
            try
            {
                tbl_NS_PhanCa shift = context.tbl_NS_PhanCas.Where(d => d.maPhanCa == shiftDetail.maPhanCa).FirstOrDefault();

                if (shift.thuHai ?? false)
                    AutoUpdateDayOfShiftExtend(2, shiftDetail);

                if (shift.thuBa ?? false)
                    AutoUpdateDayOfShiftExtend(3, shiftDetail);

                if (shift.thuTu ?? false)
                    AutoUpdateDayOfShiftExtend(4, shiftDetail);

                if (shift.thuNam ?? false)
                    AutoUpdateDayOfShiftExtend(5, shiftDetail);

                if (shift.thuSau ?? false)
                    AutoUpdateDayOfShiftExtend(6, shiftDetail);

                if (shift.thuBay ?? false)
                    AutoUpdateDayOfShiftExtend(7, shiftDetail);

                if (shift.chuNhat ?? false)
                    AutoUpdateDayOfShiftExtend(1, shiftDetail);
            }
            catch { }
        }
        private void AutoUpdateDayOfShiftExtend(int weekDayId, tbl_NS_PhanCaChiTiet shiftDetail)
        {
            var record = GetShiftDetailByIdAndDayOfWeek(weekDayId, shiftDetail.maPhanCa);
            if (record == null)
            {
                tbl_NS_PhanCaChiTiet shiftDetailNew = new tbl_NS_PhanCaChiTiet();
                shiftDetailNew.maThu = (byte)weekDayId;
                shiftDetailNew.maPhanCaChiTiet = (short)(GetMaxDetailPhanCa() + 1);
                shiftDetailNew.maPhanCa = shiftDetail.maPhanCa;
                shiftDetailNew.soGio = shiftDetail.soGio;
                shiftDetailNew.gioBatDau = shiftDetail.gioBatDau;
                shiftDetailNew.gioKetThuc = shiftDetail.gioKetThuc;
                shiftDetailNew.heSo = shiftDetail.heSo;
                shiftDetailNew.diTreChoPhep = shiftDetail.diTreChoPhep;
                shiftDetailNew.veSomChoPhep = shiftDetail.veSomChoPhep;
                shiftDetailNew.nghiGiuaCa = shiftDetail.nghiGiuaCa;
                shiftDetailNew.thoiGianNghiGiuaCa = shiftDetail.thoiGianNghiGiuaCa;
                shiftDetailNew.batDauNghiGiuaCa = shiftDetail.batDauNghiGiuaCa;
                shiftDetailNew.ketThucNghiGiuaCa = shiftDetail.ketThucNghiGiuaCa;
                shiftDetailNew.ghiChu = shiftDetail.ghiChu;
                shiftDetailNew.userName = shiftDetail.userName;
                shiftDetailNew.ngayCapNhat = shiftDetail.ngayCapNhat;
                context.tbl_NS_PhanCaChiTiets.InsertOnSubmit(shiftDetailNew);
               context.SubmitChanges();
            }
        }
        public tbl_NS_PhanCaChiTiet GetShiftDetailByIdAndDayOfWeek(int dayOfWeek, int shiftId)
        {
            var shiftDetail = context.tbl_NS_PhanCaChiTiets.Where(d => d.maPhanCa == shiftId && d.maThu == dayOfWeek).FirstOrDefault();

            return shiftDetail;
        }
        public short GetMaxDetailPhanCa()
        {
            try
            {
                return context.tbl_NS_PhanCaChiTiets.Max(d => d.maPhanCaChiTiet);
            }
            catch
            {
                return 0;
            }
        }
    }
}
