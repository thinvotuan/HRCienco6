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
    public class PhanCaNhanVienController : ApplicationController
    {
        //
        // GET: /PhanCaNhanVien/
        LinqDanhMucDataContext linqDM = new LinqDanhMucDataContext();
        LinqNhanSuDataContext linqNS = new LinqNhanSuDataContext();
        public bool? permission;
        public const string taskIDSystem = "PhanCaNhanVien";
        private IList<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan> phongBans;
        private StringBuilder buildTree;
        List<tbl_NS_PhanCa> listPhanCa = new List<tbl_NS_PhanCa>();
        tbl_NS_PhanCaChiTiet phanCaChiTiet;

        public ActionResult Index(int? page, string qSearch)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            buildTree = new StringBuilder();
            phongBans = linqNS.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan>().ToList();
            buildTree = TreePhongBanAjax.BuildTreeDepartment(phongBans);
            ViewBag.PhongBans = buildTree.ToString();

            var List = (from p in linqDM.sp_DS_NhanVienTheoPhanCa_Index(qSearch).ToList()
                        select new DMPhanca
                        {
                            maPhanca = p.maPhanca ?? 0,
                            tenPhanca = p.tenPhanca,
                            maKhoiTinhLuong = p.maKhoiTinhLuong ?? 0,
                            soLuongNhanVienTrongCa = p.soLuongNhanVienTrongCa ?? 0,
                            ghiChu = p.ghiChu,
                            ngayLap = p.ngayLap,
                        }).ToList();


            ViewBag.isGet = "True";
            int currentPageIndex = page.HasValue ? page.Value : 1;
            ViewBag.Count = List.Count();
            ViewBag.Search = qSearch;

            if (Request.IsAjaxRequest())
            {
                return PartialView("ViewIndex", List);
            }

            return View(List);
        }
        public ActionResult LoadPhanCaNhanVienChiTiet(string qSearch, string maPhongBan, int _page = 0)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            int page = _page == 0 ? 1 : _page;
            int pIndex = page;
            int total = linqNS.sp_NS_PhanCaNhanVienChiTiet(maPhongBan, qSearch).Count();
            PagingLoaderController("/PhanCaNhanVien/LoadPhanCaNhanVienChiTiet/", total, page, "?qsearch=" + qSearch);
            ViewData["lsDanhSach"] = linqNS.sp_NS_PhanCaNhanVienChiTiet(maPhongBan, qSearch).Skip(start).Take(offset).ToList();

            ViewData["qSearch"] = qSearch;
            return PartialView("_LoadPhanCaNhanVienChiTiet");
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

            tbl_NS_PhanCa phanCa = new tbl_NS_PhanCa();

            return View(phanCa);

        }
        [HttpPost]
        public ActionResult Create(FormCollection coll)
        {
            try
            {

                tbl_NS_PhanCa phanCa = new tbl_NS_PhanCa();
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
                try
                {
                    phanCa.isNghi2NgayThu7DauTienTrongThang = coll.Get("isNghi2NgayThu7DauTienTrongThang").Contains("true");
                }
                catch
                {
                    phanCa.isNghi2NgayThu7DauTienTrongThang = false;
                }
                linqDM.tbl_NS_PhanCas.InsertOnSubmit(phanCa);
                linqDM.SubmitChanges();
                BindDatatoSave(coll, phanCa.maPhanCa);
                linqDM.SubmitChanges();
                return RedirectToAction("Edit", new { id = phanCa.maPhanCa });
            }
            catch (Exception ex)
            {
                ViewData["Message"] = ex.Message;
                return View("Error");
            }
        }

        public void BindDatatoSave(FormCollection col, int maPhanCa)
        {
            List<tbl_NS_PhanCa_SoNgayTinhCongMin> soNgayTinhCong = new List<tbl_NS_PhanCa_SoNgayTinhCongMin>();
            tbl_NS_PhanCa_SoNgayTinhCongMin tc;
            string[] thang = col.GetValues("thang");
            if (thang != null && thang.Count() > 0)
            {
                for (int i = 0; i < thang.Count(); i++)
                {
                    tc = new tbl_NS_PhanCa_SoNgayTinhCongMin();
                    tc.maPhanCa = maPhanCa;
                    tc.thang = Convert.ToInt32(col.GetValues("thang")[i]);
                    tc.nam = Convert.ToInt32(col.GetValues("nam")[i]);
                    tc.soNgayTinhCongMin = Convert.ToDouble(col.GetValues("soNgayTinhCongMin")[i]);
                    soNgayTinhCong.Add(tc);
                }
            }
            var delPhanCaTC = linqDM.tbl_NS_PhanCa_SoNgayTinhCongMins.Where(d => d.maPhanCa == maPhanCa);
            linqDM.tbl_NS_PhanCa_SoNgayTinhCongMins.DeleteAllOnSubmit(delPhanCaTC);
            linqDM.tbl_NS_PhanCa_SoNgayTinhCongMins.InsertAllOnSubmit(soNgayTinhCong);
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
            var phanCa = linqDM.tbl_NS_PhanCas.Where(d => d.maPhanCa == id).FirstOrDefault();
            ViewBag.ListSNTCMin = linqDM.tbl_NS_PhanCa_SoNgayTinhCongMins.Where(d => d.maPhanCa == phanCa.maPhanCa).ToList();
            return View(phanCa);

        }
        [HttpPost]
        public ActionResult Edit(int id, FormCollection coll)
        {
            try
            {

                var phanCa = linqDM.tbl_NS_PhanCas.Where(d => d.maPhanCa == id).FirstOrDefault();
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
                if (phanCa.thuHai == false)
                {
                    FunDeletePhanCaChiTiet(id, 2);
                }
                phanCa.thuBa = coll.Get("tuesday").Contains("true");
                if (phanCa.thuBa == false)
                {
                    FunDeletePhanCaChiTiet(id, 3);
                }
                phanCa.thuTu = coll.Get("wednesday").Contains("true");
                if (phanCa.thuTu == false)
                {
                    FunDeletePhanCaChiTiet(id, 4);
                }
                phanCa.thuNam = coll.Get("thursday").Contains("true");
                if (phanCa.thuNam == false)
                {
                    FunDeletePhanCaChiTiet(id, 5);
                }
                phanCa.thuSau = coll.Get("friday").Contains("true");
                if (phanCa.thuSau == false)
                {
                    FunDeletePhanCaChiTiet(id, 6);
                }
                phanCa.thuBay = coll.Get("sathurday").Contains("true");
                if (phanCa.thuBay == false)
                {
                    FunDeletePhanCaChiTiet(id, 7);
                }
                phanCa.chuNhat = coll.Get("sunday").Contains("true");
                if (phanCa.chuNhat == false)
                {
                    FunDeletePhanCaChiTiet(id, 1);
                }
                phanCa.isNghi2NgayThu7DauTienTrongThang = coll.Get("isNghi2NgayThu7DauTienTrongThang").Contains("true");
                phanCa.userName = GetUser().userName;
                phanCa.ngayCapNhat = DateTime.Now;
                BindDatatoSave(coll, phanCa.maPhanCa);
                linqDM.SubmitChanges();
                return RedirectToAction("Edit", new { id = phanCa.maPhanCa });
            }
            catch (Exception ex)
            {
                ViewData["Message"] = ex.Message;
                return View("Error");
            }
        }
        public int FunDeletePhanCaChiTiet(int id, int thu)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return 0;
            if (!permission.Value)
                return 0;
            #endregion
            var checkMaPhanCaChiTiet = linqDM.tbl_NS_PhanCaChiTiets.Where(d => d.maPhanCa == id && d.maThu == thu).FirstOrDefault();
            if (checkMaPhanCaChiTiet != null)
            {
                linqDM.tbl_NS_PhanCaChiTiets.DeleteOnSubmit(checkMaPhanCaChiTiet);
                linqDM.SubmitChanges();
            }

            return 1;
        }

        public ActionResult Details(int id, string weekdayId, string shiftDetailId)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXemChiTiet);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            var phanCa = linqDM.tbl_NS_PhanCas.Where(d => d.maPhanCa == id).FirstOrDefault();
            var ChiTietPhanCa = (from p in linqDM.tbl_NS_PhanCaChiTiets.Where(d => d.maPhanCaChiTiet == id)
                                 select new DMPhanCaChiTiet
                                 {
                                     soGio = p.soGio,
                                     veSomChoPhep = p.veSomChoPhep ?? 0,
                                     diTreChoPhep = p.diTreChoPhep ?? 0,
                                     nghiGiuaCa = p.nghiGiuaCa ?? false,
                                     ghiChu = p.ghiChu,
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
            var phanCaChiTiet = GetShiftDetailByIdAndDayOfWeek(dayOfWeek, shiftId);
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
                json.thoiGianNghiGiuaCa = phanCaChiTiet.thoiGianNghiGiuaCa;
                json.hourStartBetweenShift = String.Format("{0:HH:mm}", phanCaChiTiet.batDauNghiGiuaCa);
                json.hourEndBetweenShift = String.Format("{0:HH:mm}", phanCaChiTiet.ketThucNghiGiuaCa);
                json.ghiChu = phanCaChiTiet.ghiChu;

                return Json(json);
            }
            else
                return Json(String.Empty);
        }
        [HttpPost]
        public ActionResult Details(int id, FormCollection colection)
        {
            var phanCa = linqDM.tbl_NS_PhanCas.Where(d => d.maPhanCa == id).FirstOrDefault();
            string date = String.Format("{0:dd/MM/yyyy}", DateTime.Now);
            if (String.IsNullOrEmpty(colection.Get("shiftDetailId")))
            {
                phanCaChiTiet = new tbl_NS_PhanCaChiTiet();
                phanCaChiTiet.maPhanCaChiTiet = (short)(GetMaxDetailPhanCa() + 1);
            }
            else
            {
                phanCaChiTiet = linqDM.tbl_NS_PhanCaChiTiets.Where(d => d.maPhanCaChiTiet == Convert.ToInt16(colection.Get("shiftDetailId"))).FirstOrDefault();
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
                phanCaChiTiet.batDauNghiGiuaCa = DateTime.ParseExact(date + " " + colection.Get("chiTiet.batDauNghiGiuaCa"), "dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
                //Convert.ToDateTime(date + " " + colection.Get("chiTiet.batDauNghiGiuaCa"));
                phanCaChiTiet.ketThucNghiGiuaCa = DateTime.ParseExact(date + " " + colection.Get("chiTiet.ketThucNghiGiuaCa"), "dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
                //Convert.ToDateTime(date + " " + colection.Get("chiTiet.ketThucNghiGiuaCa"));
            }
            else
            {
                phanCaChiTiet.thoiGianNghiGiuaCa = null;
                phanCaChiTiet.batDauNghiGiuaCa = null;
                phanCaChiTiet.ketThucNghiGiuaCa = null;
            }
            phanCaChiTiet.ghiChu = colection.Get("chiTiet.ghiChu");
            phanCaChiTiet.gioKetThuc = DateTime.ParseExact(date + " " + colection.Get("chiTiet.gioKetThuc"), "dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            //Convert.ToDateTime(date + " " + colection.Get("chiTiet.gioKetThuc"));
            phanCaChiTiet.heSo = Convert.ToDecimal(colection.Get("chiTiet.heSo"));
            phanCaChiTiet.gioBatDau = DateTime.ParseExact(date + " " + colection.Get("chiTiet.gioBatDau"), "dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            //Convert.ToDateTime(date + " " + colection.Get("chiTiet.gioBatDau"));

            phanCaChiTiet.userName = GetUser().userName;
            phanCaChiTiet.ngayCapNhat = DateTime.Now;

            if (String.IsNullOrEmpty(colection.Get("shiftDetailId")))
            {
                linqDM.tbl_NS_PhanCaChiTiets.InsertOnSubmit(phanCaChiTiet);
                linqDM.SubmitChanges();
            }
            else
                linqDM.SubmitChanges();

            if (colection.Get("autoUpdate").Contains("true"))
            {
                AutoUpdateDayOfShift(phanCaChiTiet);
            }


            return RedirectToAction("Details", new { id = phanCa.maPhanCa, weekdayId = colection.Get("weekdayId"), shiftDetailId = phanCaChiTiet.maPhanCaChiTiet });

        }

        public void AutoUpdateDayOfShift(tbl_NS_PhanCaChiTiet shiftDetail)
        {
            try
            {
                tbl_NS_PhanCa shift = linqDM.tbl_NS_PhanCas.Where(d => d.maPhanCa == shiftDetail.maPhanCa).FirstOrDefault();

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
                linqDM.tbl_NS_PhanCaChiTiets.InsertOnSubmit(shiftDetailNew);
                linqDM.SubmitChanges();
            }
        }
        public tbl_NS_PhanCaChiTiet GetShiftDetailByIdAndDayOfWeek(int dayOfWeek, int shiftId)
        {
            var shiftDetail = linqDM.tbl_NS_PhanCaChiTiets.Where(d => d.maPhanCa == shiftId && d.maThu == dayOfWeek).FirstOrDefault();

            return shiftDetail;
        }
        public short GetMaxDetailPhanCa()
        {
            try
            {
                return linqDM.tbl_NS_PhanCaChiTiets.Max(d => d.maPhanCaChiTiet);
            }
            catch
            {
                return 0;
            }
        }
        public ActionResult EditNhanVien(int id)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            var List = (from p in linqDM.tbl_NS_PhanCas
                        where p.maPhanCa == id
                        orderby p.maKhoiTinhLuong descending
                        select new DMPhanca
                        {
                            maPhanca = p.maPhanCa,
                            tenPhanca = p.tenPhanCa,
                            maKhoiTinhLuong = p.maKhoiTinhLuong,
                            tenKhoiTinhLuong = linqDM.tbl_NS_KhoiTinhLuongs.Where(d => d.maKhoiTinhLuong == p.maKhoiTinhLuong).Select(d => d.tenKhoiTinhLuong).FirstOrDefault(),
                            ghiChu = p.ghiChu,
                            ngayLap = p.ngayCapNhat,
                        }).FirstOrDefault();

            var listNhanVien = (from p in linqNS.sp_DS_NhanVienTheoPhanCa(id)
                                select new DMPhanCaNhanVien
                                {
                                    hinhThucTinhLuong = string.Empty,
                                    hoVaTen = HoVaTen(p.maNhanVien),
                                    maNhanVien = p.maNhanVien,
                                    ngayApDung = linqDM.tbl_NS_PhanCaNhanViens.Where(d => d.maNhanVien == p.maNhanVien && d.maPhanCa == id).OrderByDescending(d => d.ngayApDung).Select(d => d.ngayApDung).FirstOrDefault(),
                                    ngayVaoLam = p.ngayVaoLam,
                                    tenChucDanh = p.chucDanh,
                                    tenPhongBan = p.phongBan,
                                }).ToList();

            ViewBag.NguoiCapNhat = HoVaTen(GetUser().manv);
            List.ListNhanVien = listNhanVien;
            return View(List);
        }

        public ActionResult DanhSachNhanVien(int? page, int maPhanCa)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXemChiTiet);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            var tenCa = linqDM.tbl_NS_PhanCas.Where(d => d.maPhanCa == maPhanCa).Select(d => d.tenPhanCa).FirstOrDefault();
            ViewBag.TenCa = tenCa;
            var listNhanVien = (from p in linqNS.sp_DS_NhanVienTheoPhanCa(maPhanCa)
                                select new DMPhanCaNhanVien
                                {
                                    hinhThucTinhLuong = string.Empty,
                                    hoVaTen = HoVaTen(p.maNhanVien),
                                    maNhanVien = p.maNhanVien,
                                    ngayApDung = linqDM.tbl_NS_PhanCaNhanViens.Where(d => d.maNhanVien == p.maNhanVien && d.maPhanCa == maPhanCa).OrderByDescending(d => d.ngayApDung).Select(d => d.ngayApDung).FirstOrDefault(),
                                    ngayVaoLam = p.ngayVaoLam,
                                    tenChucDanh = p.chucDanh,
                                    tenPhongBan = p.phongBan,
                                }).ToList();
            ViewBag.isGet = "True";
            int currentPageIndex = page.HasValue ? page.Value : 1;
            ViewBag.Count = listNhanVien.Count();
            ViewBag.maPhanCa = maPhanCa;
            if (Request.IsAjaxRequest())
            {
                return PartialView("_ListNhanVien", listNhanVien.ToPagedList(currentPageIndex, 20));
            }
            return PartialView("_ListNhanVien", listNhanVien.ToPagedList(currentPageIndex, 20));
        }


        public ActionResult ListNhanVienTrongPhanCa(int? page, int? maPhanCa)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            var listNhanVien = (from p in linqNS.sp_DS_NhanVienTheoPhanCa(maPhanCa)
                                select new DMPhanCaNhanVien
                                {
                                    hinhThucTinhLuong = string.Empty,
                                    hoVaTen = HoVaTen(p.maNhanVien),
                                    maNhanVien = p.maNhanVien,
                                    ngayApDung = linqDM.tbl_NS_PhanCaNhanViens.Where(d => d.maNhanVien == p.maNhanVien && d.maPhanCa == maPhanCa).OrderByDescending(d => d.ngayApDung).Select(d => d.ngayApDung).FirstOrDefault(),
                                    ngayVaoLam = p.ngayVaoLam,
                                    tenChucDanh = p.chucDanh,
                                    tenPhongBan = p.phongBan,
                                }).ToList();

            int currentPageIndex = page.HasValue ? page.Value : 1;

            PagingLoaderController(string.Empty, listNhanVien.Count(), currentPageIndex, string.Empty);

            var tenCa = linqDM.tbl_NS_PhanCas.Where(d => d.maPhanCa == maPhanCa).Select(d => d.tenPhanCa).FirstOrDefault();
            ViewBag.TenCa = tenCa;
            ViewBag.maPhanCa = maPhanCa;

            if (Request.IsAjaxRequest())
            {
                return PartialView("ViewDanhSachNhanVienIndex", listNhanVien.Skip(start).Take(offset).ToList());
            }

            return PartialView("ViewDanhSachNhanVienIndex", listNhanVien.Skip(start).Take(offset).ToList());
        }


        [HttpPost]
        public ActionResult EditNhanVien(FormCollection coll)
        {

            try
            {
                List<tbl_NS_PhanCaNhanVien> listNhanVien = new List<tbl_NS_PhanCaNhanVien>();
                string[] maNhanVien = coll.GetValues("maNhanVien");
                if (maNhanVien != null && maNhanVien.Length > 0)
                {
                    for (int i = 0; i < maNhanVien.Length; i++)
                    {
                        var NhanVienCheck = linqDM.tbl_NS_PhanCaNhanViens.Where(d => d.maPhanCa == Convert.ToInt16(coll.Get("maPhanCa")) && d.maNhanVien == maNhanVien[i] && d.ngayApDung == DateTime.ParseExact(coll.GetValues("ngayApDung")[i], "dd/MM/yyyy", CultureInfo.InvariantCulture)).FirstOrDefault();
                        if (NhanVienCheck == null)
                        {
                            tbl_NS_PhanCaNhanVien NhanVien = new tbl_NS_PhanCaNhanVien();
                            NhanVien.maNhanVien = maNhanVien[i];
                            NhanVien.maPhanCa = Convert.ToInt16(coll.Get("maPhanCa"));
                            NhanVien.ngayApDung = DateTime.ParseExact(coll.GetValues("ngayApDung")[i], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                            NhanVien.nguoiLap = GetUser().userName;
                            NhanVien.ngayLap = DateTime.Now;
                            linqDM.tbl_NS_PhanCaNhanViens.InsertOnSubmit(NhanVien);
                            linqDM.SubmitChanges();
                        }


                    }

                }
                return RedirectToAction("EditNhanVien", new { id = coll.Get("maPhanCa") });
            }
            catch
            {
                return View("error");
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
                var chiTietPhanCa = linqDM.tbl_NS_PhanCaChiTiets.Where(d => d.maPhanCa == id).ToList();
                if (chiTietPhanCa != null && chiTietPhanCa.Count > 0)
                {
                    linqDM.tbl_NS_PhanCaChiTiets.DeleteAllOnSubmit(chiTietPhanCa);
                }

                var chiTietPhanCaNhanVien = linqDM.tbl_NS_PhanCaNhanViens.Where(d => d.maPhanCa == id).ToList();
                if (chiTietPhanCaNhanVien != null && chiTietPhanCaNhanVien.Count > 0)
                {
                    linqDM.tbl_NS_PhanCaNhanViens.DeleteAllOnSubmit(chiTietPhanCaNhanVien);
                }

                var PhanCa = linqDM.tbl_NS_PhanCas.Where(d => d.maPhanCa == id).FirstOrDefault();
                if (PhanCa != null)
                {
                    linqDM.tbl_NS_PhanCas.DeleteOnSubmit(PhanCa);
                }

                linqDM.SubmitChanges();

                return RedirectToAction("Index");
            }
            catch
            {
                return View("error");
            }
        }
        public string HoVaTen(string MaNV)
        {

            return linqNS.tbl_NS_NhanViens.Where(d => d.maNhanVien == MaNV).Select(d => d.ho + " " + d.ten).FirstOrDefault();
        }
        public ActionResult ChonNhanVien()
        {
            var buildTree = new StringBuilder();
            var phongBans = linqDM.tbl_DM_PhongBans.ToList();
            buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
            ViewBag.NVPB = buildTree.ToString();
            return PartialView("_ChonNhanVien");
        }

        public ActionResult LoadNhanVien(int? page, string searchString, string maPhongBan, int maPhanCa)
        {
            //DateTime? DateApDung = null;

            //if (!String.IsNullOrEmpty(ngayApDung))
            //    DateApDung = DateTime.ParseExact(ngayApDung, "dd/MM/yyyy", CultureInfo.InvariantCulture);

            var DaPhanCa = linqDM.tbl_NS_PhanCaNhanViens.Where(d => d.maPhanCa == maPhanCa).ToList();
            IList<sp_PB_DanhSachNhanVienResult> phongBan1s;
            phongBan1s = linqDM.sp_PB_DanhSachNhanVien(searchString, maPhongBan).Where(d => !DaPhanCa.Select(o => o.maNhanVien).Contains(d.maNhanVien)).ToList();
            ViewBag.isGet = "True";
            int currentPageIndex = page.HasValue ? page.Value : 1;
            ViewBag.Count = phongBan1s.Count();
            ViewBag.Search = searchString;
            ViewBag.maPhanCa = maPhanCa;
            ViewBag.maPhongBan = maPhongBan;
            return PartialView("_LoadChonNhanVien", phongBan1s.ToPagedList(currentPageIndex, 10));
        }
        public ActionResult DanhSachNhanVienChon(FormCollection coll, string[] MaNV)
        {
            try
            {
                //string[] maNhanVienDaAdd = coll.GetValues("maNhanVien");
                string[] splitStr = MaNV;

                var ListNhanVienChon = new List<DMPhanCaNhanVien>();
                //if (maNhanVienDaAdd != null && maNhanVienDaAdd.Length > 0)
                //{
                //    for (int i = 0; i < maNhanVienDaAdd.Length; i++)
                //    {
                //        var nhanViens = linqDM.sp_PB_DanhSachNhanVien(maNhanVienDaAdd[i], null).FirstOrDefault();
                //        DMPhanCaNhanVien NhanVien = new DMPhanCaNhanVien();
                //        NhanVien.maNhanVien = maNhanVienDaAdd[i];
                //        NhanVien.hoVaTen = nhanViens.hoVaTen;
                //        NhanVien.tenChucDanh = nhanViens.chucDanh;
                //        NhanVien.tenPhongBan = nhanViens.phongBan;
                //        NhanVien.ngayVaoLam = nhanViens.ngayVaoLam;
                //        NhanVien.ngayApDung = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                //        NhanVien.CapBac = nhanViens.soCapBac;
                //        ListNhanVienChon.Add(NhanVien);
                //    }
                //}

                if (splitStr != null && splitStr.Length > 0)
                {
                    for (int j = 0; j < splitStr.Length; j++)
                    {
                        if (!ListNhanVienChon.Select(d => d.maNhanVien).Contains(splitStr[j]))
                        {
                            var nhanViens = linqDM.sp_PB_DanhSachNhanVien(splitStr[j], null).FirstOrDefault();
                            DMPhanCaNhanVien NhanVien = new DMPhanCaNhanVien();
                            NhanVien.maNhanVien = splitStr[j];
                            NhanVien.hoVaTen = nhanViens.hoVaTen;
                            NhanVien.tenChucDanh = nhanViens.chucDanh;
                            NhanVien.tenPhongBan = nhanViens.phongBan;
                            NhanVien.ngayVaoLam = nhanViens.ngayVaoLam;
                            NhanVien.ngayApDung = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1); ;
                            NhanVien.CapBac = nhanViens.soCapBac;
                            ListNhanVienChon.Add(NhanVien);

                        }
                    }
                }
                return PartialView("DanhSachNhanVien", ListNhanVienChon);
            }
            catch
            {

                return View("error");
            }
        }
    }
}
