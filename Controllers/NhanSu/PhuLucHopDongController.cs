using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Models.HeThong;
using BatDongSan.Models.NhanSu;

namespace BatDongSan.Controllers.NhanSu
{
    public class PhuLucHopDongController : ApplicationController
    {
        private PhanBoLuongModel phanBoLuong;
        private LinqNhanSuDataContext context = new LinqNhanSuDataContext();
        private IList<PhuLucHopDongModel> phuLucs;
        private PhuLucHopDongModel model;
        private tbl_NS_HopDongLaoDong hopDong;
        private tbl_NS_PhuLucHopDong phuLuc;
        private readonly string MCV = "PhuLucHopDongLaoDong";
        private bool? permission;
        public ActionResult Index(string id, string trangThai)
        {
            try
            {
                #region Role user
                permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion
                BindDataTrangThai(MCV);
                ViewBag.DSPhuLuc = context.sp_NS_PhuLucHopDong_Index(trangThai, id).ToList();
                ViewBag.SoHopDong = id;
                if (context.sp_NS_PhuLucHopDong_Index(trangThai, id).Where(d => d.maBuocDuyet == "DTN" || d.maBuocDuyet == "DD").Count() == context.sp_NS_PhuLucHopDong_Index(trangThai, id).Count())
                {
                    ViewBag.FlagPhuLuc = "true";
                }
                else
                {
                    ViewBag.FlagPhuLuc = "false";
                }
                return View();
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }
        }

        public ActionResult ViewIndex(string id, string trangThai)
        {
            try
            {
                #region Role user
                permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion
                BindDataTrangThai(MCV);
                ViewBag.DSPhuLuc = context.sp_NS_PhuLucHopDong_Index(trangThai, id).ToList();
                ViewBag.SoHopDong = id;
                return PartialView("ViewIndex");
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }
        }

        //
        // GET: /PhuLucHopDong/Details/5

        public ActionResult Details(string id)
        {
            try
            {
                #region Role user
                permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion
                var dsPhuLuc = context.tbl_NS_PhuLucHopDongs.Where(d => d.soPhuLuc == id).FirstOrDefault();
                model = new PhuLucHopDongModel();
                ThongTinPhuLucHopDong(id);
                GetAllDropdownList();
                //   TinhPhanBoMucLuong(model.mucLuongMoi, model.maNhanVien, false, id);
                return View(model);
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }
        }

        //
        // GET: /PhuLucHopDong/Create

        public ActionResult Create(string id)
        {
            try
            {
                #region Role user
                permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion
                context = new LinqNhanSuDataContext();
                hopDong = context.tbl_NS_HopDongLaoDongs.Where(s => s.soHopDong == id).FirstOrDefault();
                SetModelData();
                model.maLoaiHopDong = hopDong.maLoaiHopDong;
                GetAllDropdownList();
                return View(model);
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }
        }

        //
        // POST: /PhuLucHopDong/Create

        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {
                #region Role user
                permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion
                // TODO: Add insert logic here
                phuLuc = new tbl_NS_PhuLucHopDong();
                BindDataToSave(collection, true);
                context.tbl_NS_PhuLucHopDongs.InsertOnSubmit(phuLuc);
                context.SubmitChanges();
                return RedirectToAction("Edit", new { id = phuLuc.soPhuLuc });
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }
        }

        //
        // GET: /PhuLucHopDong/Edit/5

        public ActionResult Edit(string id)
        {
            try
            {
                #region Role user
                permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion
                var dsPhuLuc = context.tbl_NS_PhuLucHopDongs.Where(d => d.soPhuLuc == id).FirstOrDefault();
                model = new PhuLucHopDongModel();
                ThongTinPhuLucHopDong(id);
                GetAllDropdownList();
                //TinhPhanBoMucLuong(model.mucLuongMoi, model.maNhanVien, false, id);
                return View(model);
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }
        }

        //
        // POST: /PhuLucHopDong/Edit/5

        [HttpPost]
        public ActionResult Edit(FormCollection collection)
        {
            try
            {
                #region Role user
                permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
                if (!permission.HasValue)
                    return View("LogIn");
                if (!permission.Value)
                    return View("AccessDenied");
                #endregion
                // TODO: Add update logic here
                phuLuc = new tbl_NS_PhuLucHopDong();
                BindDataToSave(collection, false);
                context.SubmitChanges();
                return RedirectToAction("Edit", new { id = phuLuc.soPhuLuc });
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }
        }

        // POST: /HopDongLaoDong/Delete/5

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

                LinqHeThongDataContext ht = new LinqHeThongDataContext();
                var delPhuLuc = context.tbl_NS_PhuLucHopDongs.Where(d => d.soPhuLuc == id);
                context.tbl_NS_PhuLucHopDongs.DeleteAllOnSubmit(delPhuLuc);
                string soHopDong = delPhuLuc.Select(d => d.soHopDong).FirstOrDefault() ?? string.Empty;
                var delChiTietPL = context.tbl_NS_PhuLucHopDongLaoDong_PhuCapNhanViens.Where(d => d.soPhuLuc == id);
                context.tbl_NS_PhuLucHopDongLaoDong_PhuCapNhanViens.DeleteAllOnSubmit(delChiTietPL);

                var nd = ht.tbl_HT_DMNguoiDuyets.Where(d => d.maPhieu == id);
                ht.tbl_HT_DMNguoiDuyets.DeleteAllOnSubmit(nd);

                context.SubmitChanges();
                ht.SubmitChanges();
                return RedirectToAction("Index", "PhuLucHopDong", new { id = soHopDong });
            }
            catch
            {
                return View();
            }
        }

        public void SetModelData()
        {
            //Kiểm tra hợp đồng đó có phụ lục hay chưa nếu có thì lấy phụ lục mới nhất và phụ lục đó đã duyệt
            DMNguoiDuyetController nd = new DMNguoiDuyetController();
            string soPhuLuc = string.Empty;
            var dsPhuLuc = context.tbl_NS_PhuLucHopDongs.Where(d => d.soHopDong == hopDong.soHopDong).Select(d => new PhuLucHopDongModel
                {
                    soPhuLuc = d.soPhuLuc,
                    Duyet = nd.GetDetailByMaPhieuTheoQuiTrinhDong(d.soPhuLuc, d.maQuiTrinhDuyet ?? 0),
                }).ToList();
            if (dsPhuLuc != null && dsPhuLuc.Count() > 0)
            {
                soPhuLuc = dsPhuLuc.OrderByDescending(d => d.soPhuLuc).Where(d => d.Duyet.maBuocDuyet == "DTN" || d.Duyet.maBuocDuyet == "DD").Select(d => d.soPhuLuc).FirstOrDefault() ?? string.Empty;
            }
            if (soPhuLuc != "")
            {
                model = (from pl in context.tbl_NS_PhuLucHopDongs
                         join hd in context.tbl_NS_HopDongLaoDongs on pl.soHopDong equals hd.soHopDong
                         join nv in context.tbl_NS_NhanViens on hd.maNhanVien equals nv.maNhanVien
                         where pl.soPhuLuc == soPhuLuc
                         select new PhuLucHopDongModel
                         {
                             idHopDong = hd.id,
                             maNhanVien = hd.maNhanVien,
                             tenNhanVien = nv.ho + " " + nv.ten,
                             soHopDong = hd.soHopDong,
                             ngayHieuLuc = pl.ngayHieuLuc,
                             giaHanDen = pl.giaHanDen,
                             nguoiCapNhat = (string)Session["TenNhanVien"],
                             doanPhi = (double?)pl.doanPhi ?? 0,
                             dangPhi = (double?)pl.dangPhi ?? 0,
                             tienDienThoai = pl.tienDienThoai ?? 0,
                             tienAnGiuaCa = pl.tienAnGiuaCa ?? false,
                             luongDongBH = (decimal?)pl.luongDongBaoHiem ?? 0,
                             khoanBoSungLuong = (decimal?)pl.khoanBoSungLuong ?? 0,
                             luong = pl.luong ?? 0,
                             tongLuong = pl.tongLuong ?? 0,
                             phuCapLuong = (double?)pl.phuCapLuong ?? 0,
                             maLoaiHopDong = pl.maLoaiHopDong,
                             phuCapKhac=(double?)pl.phuCapKhac??0
                         }).FirstOrDefault();
            }
            else
            {
                model = (from hd in context.tbl_NS_HopDongLaoDongs
                         join lhd in context.tbl_DM_LoaiHopDongLaoDongs on hd.maLoaiHopDong equals lhd.maLoaiHopDong
                         join th in context.tbl_DM_ThoiHanHopDongLaoDongs on hd.idThoiHanHopDong equals th.id
                         join nv in context.tbl_NS_NhanViens on hd.maNhanVien equals nv.maNhanVien
                         where hd.id == hopDong.id
                         select new PhuLucHopDongModel
                         {
                             idHopDong = hd.id,
                             maNhanVien = hd.maNhanVien,
                             tenNhanVien = nv.ho + " " + nv.ten,
                             mucDieuChinh = 0,
                             soHopDong = hd.soHopDong,
                             ngayHieuLuc = DateTime.Now,
                             giaHanDen = DateTime.Now,
                             nguoiCapNhat = (string)Session["TenNhanVien"],
                             doanPhi = (double?)hd.doanPhi ?? 0,
                             dangPhi = (double?)hd.dangPhi ?? 0,
                             tienDienThoai = hd.tienDienThoai ?? 0,
                             tienAnGiuaCa = hd.tienAnGiuaCa ?? false,
                             luongDongBH = hd.luongDongBaoHiem ?? 0,
                             khoanBoSungLuong = hd.khoanBoSungLuong ?? 0,
                             phuCapLuong = (double?)hd.phuCapLuong ?? 0,
                             luong = hd.luongThoaThuan ?? 0,
                             tongLuong = hd.tongLuong ?? 0,
                             maLoaiHopDong = hd.maLoaiHopDong,
                             tenLoaiHopDong = lhd.tenLoaiHopDong,
                             phuCapKhac = (double?)hd.phuCapKhac ?? 0
                         }).FirstOrDefault();
            }
            var nhanVien = context.sp_NS_NhanVien_Index(null, hopDong.maNhanVien, null, null, null).FirstOrDefault();
            model.tenChucDanh = nhanVien.TenChucDanh;
            model.soPhuLuc = IdGenerator(hopDong.soHopDong);
        }

        public string IdGenerator(string soHopDong)
        {
            StringBuilder sb = new StringBuilder();
            var date = DateTime.Now;
            int soLanTaoPhuLuc = context.tbl_NS_PhuLucHopDongs.Where(s => s.soHopDong == soHopDong).Count() + 1;
            return soHopDong + "-PLHD-" + soLanTaoPhuLuc.ToString();
        }


        #region Insert, update details phụ lục hợp đồng lao động

        public void BindDataToSave(FormCollection col, bool isCreate)
        {
            var dsHopDong = context.tbl_NS_HopDongLaoDongs.Where(d => d.soHopDong == col.Get("soHopDong")).FirstOrDefault();
            if (isCreate == true)
            {
                phuLuc.soPhuLuc = IdGenerator(dsHopDong.soHopDong);
                phuLuc.ngayLap = DateTime.Now;
                phuLuc.nguoiLap = GetUser().manv;
            }
            else
            {
                phuLuc = context.tbl_NS_PhuLucHopDongs.Where(d => d.soPhuLuc == col.Get("soPhuLuc")).FirstOrDefault();
            }

            if (!string.IsNullOrEmpty(col.Get("idThoiHanHopDong")))
            {
                phuLuc.idThoiHanHopDong = Convert.ToByte(col.Get("idThoiHanHopDong"));
            }


            phuLuc.soHopDong = col.Get("soHopDong");
            phuLuc.noiDungThayDoi = col.Get("noiDungThayDoi");
            phuLuc.ghiChu = col.Get("ghiChu");
            phuLuc.ngayHieuLuc = String.IsNullOrEmpty(col["ngayHieuLuc"]) ? (DateTime?)null : DateTime.ParseExact(col["ngayHieuLuc"], "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
            phuLuc.giaHanDen = String.IsNullOrEmpty(col["giaHanDen"]) ? (DateTime?)null : DateTime.ParseExact(col["giaHanDen"], "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
            phuLuc.tongLuong = Convert.ToDecimal(col.Get("tongLuong"));
            phuLuc.doanPhi = string.IsNullOrEmpty(col["doanPhi"]) ? 0 : Convert.ToDecimal(col["doanPhi"].Replace('%', '0').Replace(' ', '0'));
            phuLuc.dangPhi = string.IsNullOrEmpty(col["dangPhi"]) ? 0 : Convert.ToDecimal(col["dangPhi"].Replace('%', '0').Replace(' ', '0'));
            phuLuc.tienDienThoai = string.IsNullOrEmpty(col["tienDienThoai"]) ? 0 : Convert.ToDecimal(col["tienDienThoai"]);
            phuLuc.tienAnGiuaCa = col["tienAnGiuaCa"].Contains("true");
            phuLuc.luongDongBaoHiem = string.IsNullOrEmpty(col["luongDongBH"]) ? 0 : Convert.ToDecimal(col["luongDongBH"]);
            phuLuc.khoanBoSungLuong = string.IsNullOrEmpty(col["khoanBoSungLuong"]) ? 0 : Convert.ToDecimal(col["khoanBoSungLuong"]);
            phuLuc.luong = string.IsNullOrEmpty(col["luong"]) ? 0 : Convert.ToDecimal(col["luong"]);
            phuLuc.phuCapLuong = string.IsNullOrEmpty(col["phuCapLuong"]) ? 0 : Convert.ToDecimal(col["phuCapLuong"]);
            phuLuc.phuCapKhac = string.IsNullOrEmpty(col["phuCapKhac"]) ? 0 : Convert.ToDecimal(col["phuCapKhac"]);
            phuLuc.maLoaiHopDong = col["maLoaiHopDong"];
        }
        public void GetAllDropdownList()
        {
            var loaiHopDongs = context.tbl_DM_LoaiHopDongLaoDongs.ToList();
            ViewBag.LoaiHopDongs = new SelectList(loaiHopDongs, "maLoaiHopDong", "tenLoaiHopDong", model != null ? model.maLoaiHopDong : "");

            var thoiHanHopDongs = context.tbl_DM_ThoiHanHopDongLaoDongs.ToList();
            thoiHanHopDongs.Insert(0, new BatDongSan.Models.NhanSu.tbl_DM_ThoiHanHopDongLaoDong { id = 0, tenThoiHanHopDong = "" });
            ViewBag.ThoiHanHopDongs = new SelectList(thoiHanHopDongs, "id", "tenThoiHanHopDong", model.idThoiHanHopDong);
        }
        public void ThongTinPhuLucHopDong(string soPhuLuc)
        {
            DMNguoiDuyetController nd = new DMNguoiDuyetController();
            model = context.tbl_NS_PhuLucHopDongs.Where(d => d.soPhuLuc == soPhuLuc).Select(d => new PhuLucHopDongModel
                {
                    nguoiLap = d.nguoiLap,
                    soPhuLuc = d.soPhuLuc,
                    soHopDong = d.soHopDong,
                    ngayHieuLuc = d.ngayHieuLuc,
                    giaHanDen = d.giaHanDen,
                    nguoiCapNhat = ((context.tbl_NS_NhanViens.Where(c => c.maNhanVien == d.nguoiLap).Select(c => c.ho).FirstOrDefault() ?? string.Empty) + " " + (context.tbl_NS_NhanViens.Where(c => c.maNhanVien == d.nguoiLap).Select(c => c.ho).FirstOrDefault() ?? string.Empty)),
                    noiDungThayDoi = d.noiDungThayDoi,
                    ghiChu = d.ghiChu,
                    maQuiTrinhDuyet = d.maQuiTrinhDuyet ?? 0,
                    Duyet = nd.GetDetailByMaPhieuTheoQuiTrinhDong(d.soPhuLuc, d.maQuiTrinhDuyet ?? 0),
                    doanPhi = (double?)d.doanPhi ?? 0,
                    dangPhi = (double?)d.dangPhi ?? 0,
                    tienDienThoai = d.tienDienThoai ?? 0,
                    tienAnGiuaCa = d.tienAnGiuaCa ?? false,
                    luongDongBH = (decimal?)d.luongDongBaoHiem ?? 0,
                    khoanBoSungLuong = (decimal?)d.khoanBoSungLuong ?? 0,
                    luong = d.luong ?? 0,
                    tongLuong = d.tongLuong ?? 0,
                    phuCapLuong = (double?)d.phuCapLuong ?? 0,
                    idThoiHanHopDong = d.idThoiHanHopDong ?? 0,
                    maLoaiHopDong = d.maLoaiHopDong ?? string.Empty,
                    phuCapKhac = (double?)d.phuCapKhac ?? 0
                }).FirstOrDefault();

            if (model != null)
            {
                var loaiHopDong = context.tbl_DM_LoaiHopDongLaoDongs.Where(c => c.maLoaiHopDong == model.maLoaiHopDong).FirstOrDefault();
                model.maLoaiHopDong = loaiHopDong == null ? string.Empty : loaiHopDong.maLoaiHopDong ?? string.Empty;
                model.tenLoaiHopDong = loaiHopDong == null ? string.Empty : loaiHopDong.tenLoaiHopDong ?? string.Empty;
            }
            try
            {
                var tblThoiHan = context.tbl_DM_ThoiHanHopDongLaoDongs.Where(c => c.id == model.idThoiHanHopDong).FirstOrDefault();
                if (tblThoiHan != null)
                {
                    model.soThang = tblThoiHan.soThang;
                    model.tenThoiHanHopDong = tblThoiHan.tenThoiHanHopDong;
                }
            }
            catch
            {
                model.soThang = 0;
            }
            model.maNhanVien = context.tbl_NS_HopDongLaoDongs.Where(c => c.soHopDong == model.soHopDong).Select(c => c.maNhanVien).FirstOrDefault() ?? string.Empty;
            var nhanVien = context.sp_NS_NhanVien_Index(null, model.maNhanVien, null, null, null).FirstOrDefault();
            model.tenChucDanh = nhanVien.TenChucDanh;
            model.soPhuLuc = soPhuLuc;
            model.tenNhanVien = context.sp_NS_NhanVien_Index(null, model.maNhanVien, null, null, null).Select(c => c.hoVaTen).FirstOrDefault() ?? string.Empty;

            LinqHeThongDataContext ht = new LinqHeThongDataContext();
            string hoTen = (ht.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == GetUser().manv).Select(d => d.ho).FirstOrDefault() ?? string.Empty) + " " + (ht.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == GetUser().manv).Select(d => d.ten).FirstOrDefault() ?? string.Empty);
            ViewBag.HoTen = hoTen;
            int trangThaiHT = (int?)ht.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == model.soPhuLuc).Select(d => d.trangThai).FirstOrDefault() ?? 0;
            ViewBag.TenTrangThaiDuyet = ht.tbl_HT_QuiTrinhDuyet_BuocDuyets.Where(d => d.id == trangThaiHT).Select(d => d.maBuocDuyet).FirstOrDefault() ?? string.Empty;
            ViewBag.URL = Request.Url.AbsoluteUri.ToString();
            var maBuocDuyet = ht.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == model.soPhuLuc).FirstOrDefault();
            ViewBag.TenBuocDuyet = ht.tbl_HT_QuiTrinhDuyet_BuocDuyets.Where(d => d.id == (maBuocDuyet == null ? 0 : maBuocDuyet.trangThai)).Select(d => d.maBuocDuyet).FirstOrDefault() ?? string.Empty;
        }

        #endregion

        #region Tính phân bổ mức lương của phụ lục
        /// <summary>
        /// Tính phân bổ mức lương
        /// </summary>
        /// <param name="luongThoaThuan"></param>
        /// <param name="maNhanVien"></param>
        /// <returns></returns>
        public ActionResult PhanBoLuong(decimal? luongThoaThuan, string maNhanVien)
        {
            TinhPhanBoMucLuong(luongThoaThuan, maNhanVien, true, string.Empty);
            if (phanBoLuong.flag == 0)
            {
                return Json(new { mucLuong = "Chưa có dữ liệu trong Bảng Phân Tách Lương cho mức lương này" }, JsonRequestBehavior.AllowGet);
            }
            return PartialView("PartialPhanBoLuong", phanBoLuong);

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
                return View();
            }
        }
        public void TinhPhanBoMucLuong(decimal? luongThoaThuan, string maNhanVien, bool phanBo, string soPhuLuc)
        {
            phanBoLuong = new PhanBoLuongModel();
            phanBoLuong.flag = 1;
            var nhanVien = context.tbl_NS_NhanViens.Where(s => s.maNhanVien == maNhanVien).FirstOrDefault();
            phanBoLuong.tenNhanVien = nhanVien.ho + " " + nhanVien.ten;
            phanBoLuong.maNhanVien = nhanVien.maNhanVien;
            var mucLuong = context.tbl_NS_BangPhanTachMucLuongs.Where(s => s.mucLuongTu <= luongThoaThuan && s.mucLuongDen >= luongThoaThuan).FirstOrDefault();
            if (mucLuong == null)
            {
                phanBoLuong.flag = 0;
            }
            else
            {
                phanBoLuong.luongCoBan = mucLuong.luongCoBan;
                phanBoLuong.mucTinhPhanBo = mucLuong.tenMucLuong;
                phanBoLuong.mucLuongThoaThuan = luongThoaThuan;
                phanBoLuong.ngayApDung = DateTime.Now;
                phanBoLuong.tongPhuCap = luongThoaThuan - mucLuong.luongCoBan;
                //if (phanBo == true)
                //{
                phanBoLuong.chiTiets = (from ct in context.tbl_NS_BangPhanTachMucLuongChiTiets
                                        join b in context.tbl_NS_BangPhanTachMucLuongs on ct.maMucLuong equals b.maMucLuong
                                        join pc in context.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_PhuCap>() on ct.maPhuCap equals pc.maPhuCap
                                        join l in context.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_LoaiPhuCap>() on pc.loaiPhuCap equals l.id
                                        where ct.maMucLuong == mucLuong.maMucLuong
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
                //}
                //else
                //{
                //    phanBoLuong.chiTiets = (from plct in context.tbl_NS_PhuLucHopDongLaoDong_PhuCapNhanViens
                //                            join ct in context.tbl_NS_BangPhanTachMucLuongChiTiets on plct.maPhuCap equals ct.maPhuCap
                //                            join b in context.tbl_NS_BangPhanTachMucLuongs on ct.maMucLuong equals b.maMucLuong
                //                            join pc in context.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_PhuCap>() on ct.maPhuCap equals pc.maPhuCap
                //                            join l in context.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_LoaiPhuCap>() on pc.loaiPhuCap equals l.id
                //                            where ct.maMucLuong == mucLuong.maMucLuong && plct.soPhuLuc == soPhuLuc && plct.loaiTyLe == ct.loaiTyLe
                //                            select new BangPhanTachLuongChiTietModel
                //                            {
                //                                ghiChu = plct.ghiChu,
                //                                id = ct.id,
                //                                loaiTyLe = ct.loaiTyLe,
                //                                idLoaiPhuCap = pc.loaiPhuCap,
                //                                tenLoaiPhuCap = l.tenLoaiPhuCap,
                //                                maMucLuong = ct.maMucLuong,
                //                                maPhuCap = ct.maPhuCap,
                //                                salaryTemplate = plct.soTien.ToString(),
                //                                tenPhuCap = pc.tenPhuCap,
                //                                tenMucLuong = b.tenMucLuong,
                //                                tyLe = ct.tyLe
                //                            }).ToList();
                //}

                phanBoLuong.tongPhuCapKhongTheoTyLe = phanBoLuong.chiTiets.Where(s => s.loaiTyLe == "I").Select(s => Convert.ToDecimal(s.salaryTemplate)).Sum();
                phanBoLuong.tongPhuCapTheoTyLe = phanBoLuong.tongPhuCap - phanBoLuong.tongPhuCapKhongTheoTyLe;
                foreach (var item in phanBoLuong.chiTiets.Where(s => s.loaiTyLe == "II"))
                {
                    item.salaryTemplate = (item.tyLe * phanBoLuong.tongPhuCapTheoTyLe / 100).ToString();
                }

                phanBoLuong.tongTyle = phanBoLuong.chiTiets.Select(s => s.tyLe).Sum();
                ViewBag.PhanBoLuong = phanBoLuong;
            }

        }
        #endregion
    }
}
