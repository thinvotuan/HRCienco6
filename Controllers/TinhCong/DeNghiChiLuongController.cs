using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Utils.Paging;
using BatDongSan.Helper.Utils;
using System.Data;
using System.Globalization;
using System.Text;
using System.IO;
using BatDongSan.Models.NhanSu;
using System.Configuration;
using System.Net.Mail;
using BatDongSan.Models.HeThong;

namespace BatDongSan.Controllers.TinhCong
{
    public class DeNghiChiLuongController : ApplicationController
    {
        //
        // GET: /DeNghiChiLuong/
        LinqNhanSuDataContext _dataContext = new LinqNhanSuDataContext();
        LinqHeThongDataContext ht = new LinqHeThongDataContext();
        private LinqThuanVietDataContext lqThuanViet = new LinqThuanVietDataContext();
        private readonly string MCV = "DENGHICHILUONG";
        private bool? permission;
        public DeNghiChiLuong deNghiCL;
        public List<DeNghiChiLuongChiTiet> ListChiTietDeNghi;
        public ActionResult Index(int? nam, string qSearch)
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
                nam = nam ?? DateTime.Now.Year;
                var listDN = lqThuanViet.tbl_DeNghiChiLuongs.Where(d => d.nam == nam);
                if (!string.IsNullOrEmpty(qSearch))
                {
                    listDN = listDN.Where(d => d.maPhieu.Contains(qSearch));
                }
                var listCuoi = (from p1 in listDN
                                select new DeNghiChiLuong {                                    
                                    maNguoiLap = p1.nguoiLap,
                                    nam = p1.nam,
                                    ngayLap = Convert.ToDateTime(p1.ngayLap),
                                    noiDung = p1.noiDung,
                                    soPhieu = p1.maPhieu,
                                    tenNguoiLap = p1.tenNguoiLap,
                                    thang = p1.thang,
                                    trangThai = (int?)lqThuanViet.HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == p1.maPhieu).Select(d => d.trangThai).FirstOrDefault() ?? 0,
                                    tenBuocDuyet = lqThuanViet.tbl_QuiTrinhDuyet_BuocDuyets.Where(d => d.Id == ((int?)lqThuanViet.HT_DMNguoiDuyets.OrderByDescending(c => c.ID).Where(c => c.maPhieu == p1.maPhieu).Select(c => c.trangThai).FirstOrDefault() ?? 0)).Select(d => d.TenBuocDuyet).FirstOrDefault() ?? "Tạo mới",
                                    tongTienThucChuyen = lqThuanViet.tbl_DeNghiChiLuongChiTiets.Where(d => d.maPhieu == p1.maPhieu).Sum(d => d.chuyenKhoan) ?? 0
                                }).OrderByDescending(d=>d.thang).ToList();
                if (Request.IsAjaxRequest())
                {
                    return PartialView("PartialIndex", listCuoi);
                }
                ViewBag.Search = qSearch;
                BindDataNam(DateTime.Now.Year);
                return View(listCuoi);
            }
            catch
            {
                return View("Error");
            }
        }
        public void BindDataNam(int? nam)
        {
            Dictionary<int, int> dict = new Dictionary<int, int>();
            var namCurrent = DateTime.Now.Year;
            for (int i = namCurrent - 5; i <= namCurrent+1; i++)
            {
                dict.Add(i, i);
            }
            ViewBag.Nams = new SelectList(dict, "Key", "Value", nam);
        }
        private void BindDataMonth(int? month)
        {
            Dictionary<int, int> dicYear = new Dictionary<int, int>();
            for (int i = 1; i <= 13; i++)
            {
                dicYear[i] = i;
            }
            ViewBag.Thangs = new SelectList(dicYear, "Key", "Value", month);
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

            BindDataMonth(DateTime.Now.Month - 1);
            BindDataNam(DateTime.Now.Year);

            deNghiCL = new DeNghiChiLuong();
            deNghiCL.maNguoiLap = GetUser().manv;
            deNghiCL.tenNguoiLap = HoVaTen(GetUser().manv);
            deNghiCL.ngayLap = DateTime.Now;
            deNghiCL.soPhieu = CheckLetterDNCL("DNCL", GetMaxSoPhieu(), 3);
            return View(deNghiCL);
        }
        [HttpPost]
        public ActionResult Create(FormCollection coll)
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
                tbl_DeNghiChiLuongChiTiet deNghiChiTiet;
                List<tbl_DeNghiChiLuongChiTiet> lsChiTiet = new List<tbl_DeNghiChiLuongChiTiet>();
                tbl_DeNghiChiLuong deNghi = new tbl_DeNghiChiLuong();
                deNghi.maPhieu = CheckLetterDNCL("DNCL", GetMaxSoPhieu(), 3);
                deNghi.thang = Convert.ToInt32(coll.Get("thang"));
                deNghi.nam = Convert.ToInt32(coll.Get("nam"));
                deNghi.ngayLap = DateTime.Now;
                deNghi.nguoiLap = GetUser().manv;
                deNghi.tenNguoiLap = HoVaTen(GetUser().manv);
                deNghi.noiDung = coll.Get("noiDung");

                lqThuanViet.tbl_DeNghiChiLuongs.InsertOnSubmit(deNghi);

                string[] chiTiet = coll.GetValues("TenBoPhanTinhLuong");
                if (chiTiet != null)
                {
                    for (int i = 0; i < chiTiet.Count(); i++)
                    {
                        deNghiChiTiet = new tbl_DeNghiChiLuongChiTiet();
                        deNghiChiTiet.maPhieu = deNghi.maPhieu;
                        deNghiChiTiet.tenBoPhanTinhLuong = coll.GetValues("TenBoPhanTinhLuong")[i];
                        deNghiChiTiet.luongThang = Convert.ToDecimal(coll.GetValues("LuongThang")[i]);
                        deNghiChiTiet.phuCapCongTrinh = Convert.ToDecimal(coll.GetValues("PhuCapCT")[i]);
                        deNghiChiTiet.phuCapKhac = Convert.ToDecimal(coll.GetValues("PhuCapKhac")[i]);
                        deNghiChiTiet.BHXH = Convert.ToDecimal(coll.GetValues("BHXH")[i]);
                        deNghiChiTiet.thueTNCN = Convert.ToDecimal(coll.GetValues("ThueTNCN")[i]);
                        deNghiChiTiet.truyLanhTruyThu = Convert.ToDecimal(coll.GetValues("TruyLanhTruyThu")[i]);
                        deNghiChiTiet.chuyenKhoan = Convert.ToDecimal(coll.GetValues("ChuyenKhoan")[i]);
                        deNghiChiTiet.maCongTrinh = lqThuanViet.tbl_DeNghiChiLuongChiTiets.Where(d => d.tenBoPhanTinhLuong.Contains(coll.GetValues("TenBoPhanTinhLuong")[i])).Select(d => d.maCongTrinh).FirstOrDefault();
                        deNghiChiTiet.phiCongDoan = Convert.ToDouble(coll.GetValues("doanPhi")[i]);

                        lsChiTiet.Add(deNghiChiTiet);
                    }
                    lqThuanViet.tbl_DeNghiChiLuongChiTiets.InsertAllOnSubmit(lsChiTiet);
                }
                lqThuanViet.SubmitChanges();

                return RedirectToAction("Edit", "DeNghiChiLuong", new { id = deNghi.maPhieu });
            }
            catch
            {
                return View("Error");
            }
        }
        public string GetMaxSoPhieu()
        {
            return lqThuanViet.tbl_DeNghiChiLuongs.OrderByDescending(d => d.ngayLap).Select(d => d.maPhieu).FirstOrDefault();
        }
        public ActionResult Edit(string id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            BindThongTin(id);

            return View(deNghiCL);
        }
        public void BindThongTin(string id)
        {
            deNghiCL = (from p in lqThuanViet.tbl_DeNghiChiLuongs
                        where p.maPhieu == id
                        select new DeNghiChiLuong
                        {
                            maNguoiLap = p.nguoiLap,
                            tenNguoiLap = p.tenNguoiLap,
                            nam = p.nam,
                            thang = p.thang,
                            ngayLap = Convert.ToDateTime(p.ngayLap),
                            noiDung = p.noiDung,
                            soPhieu = p.maPhieu,
                        }).FirstOrDefault();

            ListChiTietDeNghi = (from ct in lqThuanViet.tbl_DeNghiChiLuongChiTiets
                                 where ct.maPhieu == id
                                 select new DeNghiChiLuongChiTiet
                                 {
                                     BHXH = Convert.ToDouble(ct.BHXH),
                                     ChuyenKhoan = Convert.ToDouble(ct.chuyenKhoan),
                                     LuongThang = Convert.ToDouble(ct.luongThang),
                                     PhuCapCT = Convert.ToDouble(ct.phuCapCongTrinh),
                                     PhuCapKhac = Convert.ToDouble(ct.phuCapKhac),
                                     TenBoPhanTinhLuong = ct.tenBoPhanTinhLuong,
                                     ThueTNCN = Convert.ToDouble(ct.thueTNCN),
                                     TruyLanhTruyThu = Convert.ToDouble(ct.truyLanhTruyThu),
                                     DoanPhi = ct.phiCongDoan ?? 0
                                 }).ToList();

            deNghiCL.DeNghiChiLuongChiTiet = ListChiTietDeNghi;

            ViewBag.lisQuiTrinhDuyet = lqThuanViet.sp_QuiTrinhDuyetAS(MCV).Select(d => new QuiTrinhDuyet
            {
                Id = d.Id,
                TenQuiTrinh = d.TenQuiTrinh,
                ChuoiBuocDuyet = d.chuoiBuocDuyet,
            }).ToList();
            ViewBag.URL = Request.Url.AbsoluteUri.ToString();
            ViewBag.trangThai = (int?)lqThuanViet.HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == id).Select(d => d.trangThai).FirstOrDefault() ?? 0;            
        }
        [HttpPost]
        public ActionResult Edit(FormCollection coll)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            var DeNghi = lqThuanViet.tbl_DeNghiChiLuongs.Where(d => d.maPhieu == coll.Get("soPhieu")).FirstOrDefault();
            DeNghi.noiDung = coll.Get("noiDung");
            lqThuanViet.SubmitChanges();
            return RedirectToAction("Edit", new { id = DeNghi.maPhieu });
        }

        
        public ActionResult Details(string id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXemChiTiet);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            BindThongTin(id);
            return View(deNghiCL);
        }

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

                var DeNghiChiTiet = lqThuanViet.tbl_DeNghiChiLuongChiTiets.Where(s => s.maPhieu == id).ToList();
                lqThuanViet.tbl_DeNghiChiLuongChiTiets.DeleteAllOnSubmit(DeNghiChiTiet);
                var DeNghi = lqThuanViet.tbl_DeNghiChiLuongs.Where(s => s.maPhieu == id).FirstOrDefault();
                lqThuanViet.tbl_DeNghiChiLuongs.DeleteOnSubmit(DeNghi);
                lqThuanViet.SubmitChanges();

                var nguoiDuyet = lqThuanViet.HT_DMNguoiDuyets.Where(d => d.maPhieu == id).ToList();
                if (nguoiDuyet != null && nguoiDuyet.Count > 0)
                {
                    lqThuanViet.HT_DMNguoiDuyets.DeleteAllOnSubmit(nguoiDuyet);
                    lqThuanViet.SubmitChanges();
                }
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
                return RedirectToAction("Details", new { id = id });// detail de duyet
            }
            catch
            {
                return View("Error");
            }
        }
        public string CheckLetterDNCL(string preString, string maxValue, int length)
        {
            string yearCurrent = DateTime.Now.Year.ToString().Substring(2, 2);
            string monthCurrent = DateTime.Now.Month.ToString(); // "4"
            //khi thang hien tai nho hon 9 thi cong them "0" vao
            if (Convert.ToInt32(monthCurrent) <= 9)
            {
                monthCurrent = "0" + monthCurrent;
            }
            //Khi tham so select o database la null khoi tao so dau tien
            if (String.IsNullOrEmpty(maxValue))
            {
                string ret = "1";
                while (ret.Length < length)
                {
                    ret = "0" + ret;
                }
                return preString + yearCurrent + monthCurrent + "-" + ret;
            }
            else
            {
                string preStringMax = maxValue.Substring(0, maxValue.IndexOf("-") - 4);
                string maxNumber = maxValue.Substring(maxValue.IndexOf("-") + 1);
                string monthYear = maxValue.Substring(maxValue.IndexOf("-") - 4, 4);
                string monthDb = monthYear.Substring(2, 2); //as "04"

                string stringTemp = maxNumber;
                //Khi thang trong gia tri max bang voi thang create thi cong len cho 1
                if (monthDb == monthCurrent)
                {
                    int strToInt = Convert.ToInt32(maxNumber);
                    maxNumber = Convert.ToString(strToInt + 1);
                    while (maxNumber.Length < stringTemp.Length)
                        maxNumber = "0" + maxNumber;
                }
                else //reset
                {
                    maxNumber = "1";
                    while (maxNumber.Length < stringTemp.Length)
                    {
                        maxNumber = "0" + maxNumber;
                    }
                }

                return preStringMax + yearCurrent + monthCurrent + "-" + maxNumber;
            }
        }
    }
}
