using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Helper.Utils;
using BatDongSan.Models.DanhMuc;
using BatDongSan.Models.HeThong;
using BatDongSan.Models.NhanSu;
using BatDongSan.Utils.Paging;
namespace BatDongSan.Controllers.NhanSu
{
    public class DanhGiaTinNhiemController : ApplicationController
    {
        private LinqNhanSuDataContext hr = new LinqNhanSuDataContext();
        private StringBuilder buildTree = null;
        private IList<BatDongSan.Models.NhanSu.sp_NS_NhanVien_IndexResult> nhanViens;
        private IList<tbl_DM_PhongBan> phongBans;
        private LinqDanhMucDataContext linqDanhMuc = new LinqDanhMucDataContext();
        private LinqHeThongDataContext contentHT = new LinqHeThongDataContext();

        private readonly string MCV = "DanhGiaTinNhiem";
        private bool? permission;
        public string soPhieu = string.Empty;
        public ActionResult Index(int? page, int? pageSize, string loaiHopDong, string searchString)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            hr = new LinqNhanSuDataContext();
            ViewBag.lsDanhSach = hr.sp_NS_PhieuDanhGia_Index(GetUser().manv).ToList();
            return View();
        }
        public ActionResult ThongBao(int? id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            var ds = (Sys_User)Session["User"];
            Session["countNhacViec"] = contentHT.sp_HT_NhacViec(ds == null ? string.Empty : ds.manv, null, null).Count();
            //Theme Color
            var listTheme = contentHT.tbl_DM_TienIchThemes.OrderByDescending(d => d.id).ToList();
            Session["listTheme"] = listTheme;

            //Check idColor in table user
            var mauMacDinhUser = contentHT.Sys_Users.Where(d => d.manv == GetUser().manv).Select(d => new { maMau = d.idColor }).ToArray();
            if (mauMacDinhUser != null)
            {
                if (mauMacDinhUser.FirstOrDefault() != null)
                {
                    Session["maMau"] = mauMacDinhUser.FirstOrDefault().maMau;
                }
            }
            if (Session["maMau"] == null)
            {
                var mauMacDinh = contentHT.tbl_DM_TienIchThemes.Where(d => d.trangThai == 1).Select(d => new { maMau = d.maMau, trangThai = d.trangThai }).ToArray();
                if (mauMacDinh != null)
                {
                    if (mauMacDinh.FirstOrDefault() != null)
                    {
                        Session["maMau"] = mauMacDinh.FirstOrDefault().maMau;
                    }
                    else
                    {
                        Session["maMau"] = "#1A9DCC";
                    }
                }
            }

            // End Theme
            // Bat dau thong bao
            var NgayBatDauThongBao1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayBatDauThongBao1").Select(d => d.giaTriNgayThang).FirstOrDefault();
            var NgayKetThucThongBao1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayKetThucThongBao1").Select(d => d.giaTriNgayThang).FirstOrDefault();

            //var NgayBatDauThongBao2 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayBatDauThongBao2").Select(d => d.giaTriNgayThang).FirstOrDefault();
            //var NgayKetThucThongBao2 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayKetThucThongBao2").Select(d => d.giaTriNgayThang).FirstOrDefault();

            // Bat dau danh gia
            var NgayBatDau1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayBatDau1").Select(d => d.giaTriNgayThang).FirstOrDefault();
            var NgayKetThuc1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayKetThuc1").Select(d => d.giaTriNgayThang).FirstOrDefault();

            //var NgayBatDau2 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayBatDau2").Select(d => d.giaTriNgayThang).FirstOrDefault();
            //var NgayKetThuc2 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayKetThuc2").Select(d => d.giaTriNgayThang).FirstOrDefault();
            // Ket thuc danh gia

            var NgayBatDauCongBo1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayBatDauCongBo1").Select(d => d.giaTriNgayThang).FirstOrDefault();
            var NgayKetThucCongBo1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayKetThucCongBo1").Select(d => d.giaTriNgayThang).FirstOrDefault();

            //var NgayBatDauCongBo2 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayBatDauCongBo2").Select(d => d.giaTriNgayThang).FirstOrDefault();
            //var NgayKetThucCongBo2 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayKetThucCongBo2").Select(d => d.giaTriNgayThang).FirstOrDefault();

            var TrangThaiDanhGia = 0;
            var TrangThaiMoDanhGia = 0;

            int checkNgayCB = 0;
            int StatusDanhGia = 0;
            int showThongBao = 0;

            // StatusDanhGia:
            //1: Ngay bat dau thong bao 1 -> Ngay ket thuc thong bao  1
            //2: Ngay bat dau danh gia 1 -> Ngay ket thuc danh gia 1
            //3: Ngay bat dau cong bo 1 -> Ngay ket thuc cong bo 1
            //4: Ngay bat dau thong bao  2 -> Ngay ket thuc thong bao  2
            //5: Ngay bat dau danh gia 2 -> Ngay ket thuc danh gia 2
            //6: Ngay bat dau cong bo 2 -> Ngay ket thuc cong bo 2

            string NgayCongBo = string.Empty;
            var DaTaoKy1 = hr.tbl_NS_PhieuDanhGias.Where(t => t.maNhanVien.Equals(GetUser().manv) && t.kyDanhGia == Convert.ToInt32(hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "KyDanhGia").Select(d => d.giaTri).FirstOrDefault()) && t.nam == DateTime.Now.Year).FirstOrDefault();
            //var DaTaoKy2 = hr.tbl_NS_PhieuDanhGias.Where(t => t.maNhanVien.Equals(GetUser().manv) && t.qui == 2 && t.nam == DateTime.Now.Year - 1).FirstOrDefault();


            if ((DateTime.Now >= NgayBatDauThongBao1 && DateTime.Now <= NgayKetThucThongBao1))
            {
                showThongBao = 1;
            }
            if (showThongBao == 0)
            {
                return RedirectToAction("Index", "Home");
            }
            if ((DateTime.Now >= NgayBatDau1 && DateTime.Now <= NgayKetThuc1))
            {
                StatusDanhGia = 1;
            }
            if ((DateTime.Now >= NgayBatDauCongBo1 && DateTime.Now <= NgayKetThucThongBao1))
            {
                StatusDanhGia = 2;
            }



            string NgayDanhGia = string.Empty;
            if (DateTime.Now >= NgayBatDau1 && DateTime.Now <= NgayKetThuc1)
            {

                var PhieuKy1 = hr.tbl_NS_PhieuDanhGias.Where(t => t.maNhanVien.Equals(GetUser().manv) && (NgayBatDau1 <= t.ngayLap && NgayKetThuc1 >= t.ngayLap) && t.kyDanhGia == Convert.ToInt32(hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "KyDanhGia").Select(d => d.giaTri).FirstOrDefault()) && t.nam == DateTime.Now.Year).FirstOrDefault();
                NgayDanhGia = "Đánh giá từ " + (NgayBatDau1.HasValue ? NgayBatDau1.Value.ToString("dd/MM/yyyy") : string.Empty) + " đến " + (NgayKetThuc1.HasValue ? NgayKetThuc1.Value.ToString("dd/MM/yyyy") : string.Empty);
                NgayCongBo = "Ngày công bố: " + (NgayBatDauCongBo1.HasValue ? NgayBatDauCongBo1.Value.ToString("dd/MM/yyyy") : string.Empty);
                if (PhieuKy1 != null)
                {
                    if (PhieuKy1.trangThai == 1)
                    {
                        TrangThaiDanhGia = 2;
                    }
                    else
                    {
                        TrangThaiDanhGia = 1;
                    }
                    ViewData["maPhieuDanhGia"] = PhieuKy1.maPhieuDanhGia;
                }

            }
            //Kỳ 2
            //if (DateTime.Now >= NgayBatDau2 && DateTime.Now <= NgayKetThuc2)
            //{
            //    TrangThaiMoDanhGia = 1;
            //    var PhieuKy2 = hr.tbl_NS_PhieuDanhGias.Where(t => t.maNhanVien.Equals(GetUser().manv) && (NgayBatDau2 <= t.ngayLap && NgayKetThuc2 >= t.ngayLap) && t.qui == 2 && t.nam == DateTime.Now.Year - 1).FirstOrDefault();
            //    NgayDanhGia = "Đánh giá từ " + (NgayBatDau2.HasValue ? NgayBatDau2.Value.ToString("dd/MM/yyyy") : string.Empty) + " đến " + (NgayKetThuc2.HasValue ? NgayKetThuc2.Value.ToString("dd/MM/yyyy") : string.Empty);
            //    NgayCongBo = "Ngày công bố: " + (NgayBatDauCongBo2.HasValue ? NgayBatDauCongBo2.Value.ToString("dd/MM/yyyy") : string.Empty);
            //    if (PhieuKy2 != null)
            //    {

            //        if (PhieuKy2.trangThai == 1)
            //        {
            //            TrangThaiDanhGia = 2;
            //        }
            //        else
            //        {
            //            TrangThaiDanhGia = 1;
            //        }
            //        ViewData["maPhieuDanhGia"] = PhieuKy2.maPhieuDanhGia;
            //    }

            //}

            ViewBag.TrangThaiDanhGia = TrangThaiDanhGia;
            ViewBag.TrangThaiMoDanhGia = TrangThaiMoDanhGia;
            ViewBag.checkNgayCB = checkNgayCB;
            ViewBag.StatusDanhGia = StatusDanhGia;
            ViewBag.NgayDanhGia = NgayDanhGia;
            ViewBag.NgayCongBo = NgayCongBo;
            ViewBag.showThongBao = showThongBao;
            //}

            return View();

        }
        public int KiemTraTrangThaiDanhGia()
        {

            var NgayBatDau1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayBatDau1").Select(d => d.giaTriNgayThang).FirstOrDefault();
            var NgayKetThuc1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayKetThuc1").Select(d => d.giaTriNgayThang).FirstOrDefault();

            //var NgayBatDau2 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayBatDau2").Select(d => d.giaTriNgayThang).FirstOrDefault();
            //var NgayKetThuc2 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayKetThuc2").Select(d => d.giaTriNgayThang).FirstOrDefault();


            var NgayCB1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayCongBoKQ1").Select(d => d.giaTriNgayThang).FirstOrDefault();
            //var NgayCB2 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayCongBoKQ2").Select(d => d.giaTriNgayThang).FirstOrDefault();

            var TrangThaiDanhGia = 0;
            var TrangThaiMoDanhGia = 0;

            int checkNgayCB = 0;
            string NgayCongBo = string.Empty;
            var DaTaoKy1 = hr.tbl_NS_PhieuDanhGias.Where(t => t.maNhanVien.Equals(GetUser().manv) && t.kyDanhGia == Convert.ToInt32(hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "KyDanhGia").Select(d => d.giaTri).FirstOrDefault()) && t.nam == DateTime.Now.Year).FirstOrDefault();



            if (DateTime.Now >= NgayCB1)
            {
                checkNgayCB = 1;
            }




            //Kỳ 1

            string NgayDanhGia = string.Empty;
            if (DateTime.Now >= NgayBatDau1 && DateTime.Now <= NgayKetThuc1)
            {
                TrangThaiMoDanhGia = 1;
                var PhieuKy1 = hr.tbl_NS_PhieuDanhGias.Where(t => t.maNhanVien.Equals(GetUser().manv) && (NgayBatDau1 <= t.ngayLap && NgayKetThuc1 >= t.ngayLap) && t.kyDanhGia == Convert.ToInt32(hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "KyDanhGia").Select(d => d.giaTri).FirstOrDefault()) && t.nam == DateTime.Now.Year).FirstOrDefault();
                NgayDanhGia = "Đánh giá từ " + (NgayBatDau1.HasValue ? NgayBatDau1.Value.ToString("dd/MM/yyyy") : string.Empty) + " đến " + (NgayKetThuc1.HasValue ? NgayKetThuc1.Value.ToString("dd/MM/yyyy") : string.Empty);
                NgayCongBo = "Ngày công bố: " + (NgayCB1.HasValue ? NgayCB1.Value.ToString("dd/MM/yyyy") : string.Empty);
                if (PhieuKy1 != null)
                {
                    if (PhieuKy1.trangThai == 1)
                    {
                        TrangThaiDanhGia = 2;
                    }
                    else
                    {
                        TrangThaiDanhGia = 1;
                    }
                    ViewData["maPhieuDanhGia"] = PhieuKy1.maPhieuDanhGia;
                }

            }

            ViewBag.TrangThaiDanhGia = TrangThaiDanhGia;

            return TrangThaiDanhGia;

        }

        // Cau hinh
        public ActionResult Cauhinh()
        {
            #region Role user
            permission = GetPermission("CauHinhPhieuDanhGia", BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            hr = new LinqNhanSuDataContext();
            var lsCauHinh = hr.tbl_NS_Cauhinhs.ToList();
            var ValueKyDanhGia = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("kyDanhGia")).Select(d => d.giaTri).FirstOrDefault();

            var kqKyDG = Convert.ToInt32(ValueKyDanhGia);

            kyDanhGias(kqKyDG);
            ViewBag.lsCauHinh = lsCauHinh;
            return View();
        }
        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Cauhinh(FormCollection collection)
        {
            #region Role user
            permission = GetPermission("CauHinhPhieuDanhGia", BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            hr = new LinqNhanSuDataContext();
            var LoaiXuatSac = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("LoaiXuatSac")).FirstOrDefault();
            var LoaiA = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("LoaiA")).FirstOrDefault();
            var LoaiB = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("LoaiB")).FirstOrDefault();
            var LoaiC = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("LoaiC")).FirstOrDefault();
            var KyDanhGia = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("kyDanhGia")).FirstOrDefault();
            var NgangCap = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("NgangCap")).FirstOrDefault();
            var CapTren = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("CapTren")).FirstOrDefault();
            var CapDuoi = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("CapDuoi")).FirstOrDefault();
            var HeSoQuiDoi = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("HeSoQuiDoi")).FirstOrDefault();

            var NgayBatDauThongBao1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("NgayBatDauThongBao1")).FirstOrDefault();
            var NgayKetThucThongBao1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("NgayKetThucThongBao1")).FirstOrDefault();
            var NgayBatDau1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("NgayBatDau1")).FirstOrDefault();
            var NgayKetThuc1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("NgayKetThuc1")).FirstOrDefault();
            var NgayBatDauCongBo1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("NgayBatDauCongBo1")).FirstOrDefault();
            //var NgayKetThucCongBo1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("NgayKetThucCongBo1")).FirstOrDefault();

            //var NgayBatDau2 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("NgayBatDau2")).FirstOrDefault();
            //var NgayKetThuc2 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("NgayKetThuc2")).FirstOrDefault();
            //var NgayBatDauThongBao2 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("NgayBatDauThongBao2")).FirstOrDefault();
            //var NgayKetThucThongBao2 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("NgayKetThucThongBao2")).FirstOrDefault();
            //var NgayBatDauCongBo2 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("NgayBatDauCongBo2")).FirstOrDefault();
            //var NgayKetThucCongBo2 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai.Equals("NgayKetThucCongBo2")).FirstOrDefault();

            try
            { NgayBatDauThongBao1.giaTriNgayThang = DateTime.ParseExact(collection.GetValues("NgayBatDauThongBao1")[0], "dd/MM/yyyy", null); }
            catch { NgayBatDauThongBao1.giaTriNgayThang = null; }

            try
            { NgayKetThucThongBao1.giaTriNgayThang = DateTime.ParseExact(collection.GetValues("NgayKetThucThongBao1")[0], "dd/MM/yyyy", null); }
            catch { NgayKetThucThongBao1.giaTriNgayThang = null; }
            try
            { NgayBatDau1.giaTriNgayThang = DateTime.ParseExact(collection.GetValues("NgayBatDau1")[0], "dd/MM/yyyy", null); }
            catch { NgayBatDau1.giaTriNgayThang = null; }

            try
            { NgayKetThuc1.giaTriNgayThang = DateTime.ParseExact(collection.GetValues("NgayKetThuc1")[0], "dd/MM/yyyy", null); }
            catch { NgayKetThuc1.giaTriNgayThang = null; }

            try
            { NgayBatDauCongBo1.giaTriNgayThang = DateTime.ParseExact(collection.GetValues("NgayBatDauCongBo1")[0], "dd/MM/yyyy", null); }
            catch { NgayBatDauCongBo1.giaTriNgayThang = null; }
            //try
            //{ NgayKetThucCongBo1.giaTriNgayThang = DateTime.ParseExact(collection.GetValues("NgayKetThucCongBo1")[0], "dd/MM/yyyy", null); }
            //catch { NgayKetThucCongBo1.giaTriNgayThang = null; }

            //Kỳ 2
            //try
            //{ NgayBatDauThongBao2.giaTriNgayThang = DateTime.ParseExact(collection.GetValues("NgayBatDauThongBao2")[0], "dd/MM/yyyy", null); }
            //catch { NgayBatDauThongBao2.giaTriNgayThang = null; }

            //try
            //{ NgayKetThucThongBao2.giaTriNgayThang = DateTime.ParseExact(collection.GetValues("NgayKetThucThongBao2")[0], "dd/MM/yyyy", null); }
            //catch { NgayKetThucThongBao2.giaTriNgayThang = null; }
            //try
            //{ NgayBatDau2.giaTriNgayThang = DateTime.ParseExact(collection.GetValues("NgayBatDau2")[0], "dd/MM/yyyy", null); }
            //catch { NgayBatDau2.giaTriNgayThang = null; }

            //try
            //{ NgayKetThuc2.giaTriNgayThang = DateTime.ParseExact(collection.GetValues("NgayKetThuc2")[0], "dd/MM/yyyy", null); }
            //catch { NgayKetThuc2.giaTriNgayThang = null; }

            //try
            //{ NgayBatDauCongBo2.giaTriNgayThang = DateTime.ParseExact(collection.GetValues("NgayBatDauCongBo2")[0], "dd/MM/yyyy", null); }
            //catch { NgayBatDauCongBo2.giaTriNgayThang = null; }
            //try
            //{ NgayKetThucCongBo2.giaTriNgayThang = DateTime.ParseExact(collection.GetValues("NgayKetThucCongBo2")[0], "dd/MM/yyyy", null); }
            //catch { NgayKetThucCongBo2.giaTriNgayThang = null; }

            try
            {
                LoaiXuatSac.giaTri = Convert.ToDouble(collection.GetValues("LoaiXuatSac")[0]);

            }
            catch { LoaiXuatSac.giaTri = 0; }

            try
            {
                LoaiA.giaTri = Convert.ToDouble(collection.GetValues("LoaiA")[0]);
            }
            catch
            {
                LoaiA.giaTri = 0;
            }

            try
            {
                LoaiB.giaTri = Convert.ToDouble(collection.GetValues("LoaiB")[0]);
            }
            catch { LoaiB.giaTri = 0; }

            try
            {
                LoaiC.giaTri = Convert.ToDouble(collection.GetValues("LoaiC")[0]);
            }
            catch { LoaiC.giaTri = 0; }
            try
            {
                KyDanhGia.giaTri = Convert.ToDouble(collection.GetValues("KyDanhGia")[0]);
            }
            catch { KyDanhGia.giaTri = 0; }

            try
            {
                NgangCap.giaTri = Convert.ToDouble(collection.GetValues("NgangCap")[0]);
            }
            catch { NgangCap.giaTri = 0; }

            try
            {
                CapTren.giaTri = Convert.ToDouble(collection.GetValues("CapTren")[0]);
            }
            catch { CapTren.giaTri = 0; }

            try
            {
                CapDuoi.giaTri = Convert.ToDouble(collection.GetValues("CapDuoi")[0]);
            }
            catch { CapDuoi.giaTri = 0; }

            try
            {
                HeSoQuiDoi.giaTri = Convert.ToDouble(collection.GetValues("HeSoQuiDoi")[0]);
            }
            catch
            {
                HeSoQuiDoi.giaTri = 0;
            }
            hr.SubmitChanges();
            ViewBag.lsCauHinh = hr.tbl_NS_Cauhinhs.ToList();
            return RedirectToAction("Cauhinh");
            //return View();
        }
        // Buoc 1: Load danh sach nhan vien de chon:
        public ActionResult ChonDanhSachNhanVien()
        {
            // Bat dau thong bao
            var NgayBatDau1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayBatDau1").Select(d => d.giaTriNgayThang).FirstOrDefault();
            var NgayKetThuc1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayKetThuc1").Select(d => d.giaTriNgayThang).FirstOrDefault();


            // Check time danh gia
            if (DateTime.Now > NgayKetThuc1 || DateTime.Now < NgayBatDau1)
            {
                return RedirectToAction("Index", "Home");
            }

            if (GetUser() == null)
            {
                return RedirectToAction("LogOn", "Account");
            }
            hr = new LinqNhanSuDataContext();
            if (KiemTraTrangThaiDanhGia() == 2)
            {

                return RedirectToAction("Index", "DanhGiaTinNhiem");
            }

            PhieuKyNay();

            var danhSachMacDinh = hr.sp_NS_DanhSachNhanVienMacDinh(GetUser().manv).Where(d => d.maNhanVien != GetUser().manv).ToList();
            var DanhSachDaChon = hr.tbl_NS_ChonNhanVienDanhGias.Where(d => d.maNguoiDanhGia == GetUser().manv).ToList();
            List<DanhSachNhanVienDaChonModel> NhanVienChon = new List<DanhSachNhanVienDaChonModel>();
            if (DanhSachDaChon != null && DanhSachDaChon.Count > 0)
            {
                var listChon = (from p in hr.sp_NS_DanhSachNhanVienDaChon(GetUser().manv)
                                select new DanhSachNhanVienDaChonModel
                                {
                                    MaNhanVien = p.maNhanVien,
                                    TenNhanVien = p.hoTen,
                                    MaPhongBan = p.maPhongBan,
                                    TenPhongBan = p.tenPhongBan,
                                    MaChucDanh = p.maChucDanh,
                                    TenChucDanh = p.TenChucDanh,
                                    CapBac = p.soCapBac,
                                    NgaySinh = p.ngaySinh
                                    //Avatar = p.fileDinhKemAnhDaiDien,

                                }).ToList();
                NhanVienChon = listChon;
            }
            else
            {
                var listChon = (from p in danhSachMacDinh
                                select new DanhSachNhanVienDaChonModel
                                {
                                    MaNhanVien = p.maNhanVien,
                                    TenNhanVien = p.hoTen,
                                    MaPhongBan = p.maPhongBan,
                                    TenPhongBan = p.tenPhongBan,
                                    MaChucDanh = p.maChucDanh,
                                    TenChucDanh = p.TenChucDanh,
                                    CapBac = p.SoCapBac,
                                    NgaySinh = p.ngaySinh
                                }).ToList();
                NhanVienChon = listChon;
            }

            var CT = NhanVienChon.Select(d => d.MaPhongBan).Distinct().ToList();
            var listPhongBan = (from p in CT
                                select new PhongBanModel
                                {
                                    MaPhongBan = p,
                                    Ten = linqDanhMuc.tbl_DM_PhongBans.Where(d => d.maPhongBan.Equals(p)).Select(d => d.maPhongBan).FirstOrDefault()
                                }).Distinct().ToList();
            ViewBag.ListNhanVien = NhanVienChon;
            ViewBag.ListPhongBan = listPhongBan;
            return View();
        }
        //Phieu ky nay
        public string PhieuKyNay()
        {
            var NgayBatDau1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayBatDau1").Select(d => d.giaTriNgayThang).FirstOrDefault();
            var NgayKetThuc1 = hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "NgayKetThuc1").Select(d => d.giaTriNgayThang).FirstOrDefault();


            string DanhGiaKy = "1";
            string maPhieuKyNay = string.Empty;
            DanhGiaKy = Convert.ToString(hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "KyDanhGia").Select(d => d.giaTri).FirstOrDefault());
            if (DateTime.Now >= NgayBatDau1 && DateTime.Now <= NgayKetThuc1)
            {

                var PhieuKy1 = hr.tbl_NS_PhieuDanhGias.Where(t => t.maNhanVien.Equals(GetUser().manv) && (NgayBatDau1 <= t.ngayLap && NgayKetThuc1 >= t.ngayLap) && t.kyDanhGia == Convert.ToInt32(hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "KyDanhGia").Select(d => d.giaTri).FirstOrDefault()) && t.nam == DateTime.Now.Year).FirstOrDefault();

                if (PhieuKy1 != null)
                {
                    maPhieuKyNay = PhieuKy1.maPhieuDanhGia;
                    soPhieu = maPhieuKyNay;
                }

            }

            ViewBag.maPhieuTinNhiemEdit = maPhieuKyNay;
            return DanhGiaKy;
        }
        //Load danh sach nhan vien khi click button add them:
        public ActionResult DanhSachNhanVien()
        {
            try
            {
                buildTree = new StringBuilder();
                phongBans = linqDanhMuc.tbl_DM_PhongBans.ToList();
                buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
                ViewBag.departments = buildTree.ToString();
                ViewBag.shiftType = linqDanhMuc.tbl_NS_PhanCas.Where(t => t.tenPhanCa != "").ToList();
                ViewBag.page = 0;
                ViewBag.total = 0;
                return View(phongBans);
            }
            catch (Exception ex)
            {

                ViewBag.Message = ex.Message;
                return View("Error");
            }
        }
        public ActionResult LoadNhanVien(string qSearch, int _page, string parrentId)
        {
            try
            {

                string parentID = linqDanhMuc.tbl_DM_PhongBans.Where(d => d.maPhongBan == parrentId).Select(d => d.maPhongBan).FirstOrDefault() ?? string.Empty;
                if (String.IsNullOrEmpty(parentID))
                {
                    parrentId = string.Empty;
                }
                int page = _page == 0 ? 1 : _page;
                int pIndex = page;
                int total = linqDanhMuc.sp_PB_DanhSachNhanVien(qSearch, parrentId).Count();
                PagingLoaderController("/DanhGiaTinNhiem/Index/", total, page, "?qsearch=" + qSearch + "&parrentId=" + parrentId);
                ViewBag.nhanVien = linqDanhMuc.sp_PB_DanhSachNhanVien(qSearch, parrentId).Skip(start).Take(offset).ToList();
                ViewBag.parrentId = parrentId;
                ViewBag.qSearch = qSearch ?? string.Empty;
                ViewBag.currentMaNV = GetUser().manv;
                return PartialView("LoadNhanVien");
            }
            catch (Exception e)
            {
                ViewData["Message"] = e.Message;
                return View("Error");
            }
        }

        public ActionResult GetDanhSachNhanVien(FormCollection collection, string nhanVien, string phongBanCT)
        {
            hr = new LinqNhanSuDataContext();
            //Add danh sách nhân viên ở dưới lưới
            List<DanhSachNhanVienDaChonModel> nhanVienDC = new List<DanhSachNhanVienDaChonModel>();
            DanhSachNhanVienDaChonModel ds;
            string[] phongBan = collection.GetValues("maPhongBan");
            string[] maNhanViens = nhanVien.Split(';');
            string[] maPhongBans = phongBanCT.Split(';');
            if (phongBan != null)
            {
                for (int j = 0; j < phongBan.Count(); j++)
                {
                    DateTime? NgaySinh = null;

                    ds = new DanhSachNhanVienDaChonModel();
                    ds.MaPhongBan = phongBan[j];

                    ds.MaNhanVien = collection.GetValues("maNhanVien")[j] ?? null;


                    ds.TenNhanVien = collection.GetValues("tenNhanVien")[j] ?? null;


                    ds.TenPhongBan = collection.GetValues("tenPhongBan")[j] ?? null;


                    ds.TenChucDanh = collection.GetValues("tenChucDanh")[j] ?? null;
                    ds.CapBac = 1;
                    if (!String.IsNullOrEmpty(collection.GetValues("capBac")[j]))
                    {
                        ds.CapBac = int.Parse(collection.GetValues("capBac")[j]);
                    }
                    if (!String.IsNullOrEmpty(collection.GetValues("ngaySinh")[j]))
                        ds.NgaySinh = DateTime.ParseExact(collection.GetValues("ngaySinh")[j], "dd/MM/yyyy", null);
                    nhanVienDC.Add(ds);
                }
            }

            //Add những nhân viên đã chọn           
            if (maNhanViens != null)
            {
                for (int j = 0; j < maNhanViens.Count(); j++)
                {
                    if (maNhanViens[j] != null && maNhanViens[j] != "")
                    {
                        ds = new DanhSachNhanVienDaChonModel();
                        ds.MaPhongBan = maPhongBans[j];
                        ds.MaNhanVien = maNhanViens[j];
                        ds.TenNhanVien = hr.tbl_NS_NhanViens.Where(d => d.maNhanVien == maNhanViens[j]).Select(d => d.ho + " " + d.ten).FirstOrDefault() ?? string.Empty;
                        ds.TenPhongBan = linqDanhMuc.tbl_DM_PhongBans.Where(d => d.maPhongBan == maPhongBans[j]).Select(d => d.tenPhongBan).FirstOrDefault() ?? string.Empty;
                        ds.TenChucDanh = hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == maNhanViens[j]).Select(d => d.TenChucDanh).FirstOrDefault() ?? string.Empty;
                        ds.CapBac = hr.Sys_ChucDanhs.Where(d => d.MaChucDanh == hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(x => x.maNhanVien == maNhanViens[j]).Select(r => r.maChucDanh).FirstOrDefault()).Select(d => d.SoCapBac).FirstOrDefault() ?? 1;
                        ds.NgaySinh = hr.tbl_NS_NhanViens.Where(d => d.maNhanVien == maNhanViens[j]).Select(d => d.ngaySinh).FirstOrDefault() ?? null;
                        nhanVienDC.Add(ds);
                    }
                }
            }
            //Gom nhóm phòng ban
            var listPhongBan = nhanVienDC.Select(p => new PhongBanModel
            {
                MaPhongBan = p.MaPhongBan ?? string.Empty,
                Ten = linqDanhMuc.tbl_DM_PhongBans.Where(d => d.maPhongBan == p.MaPhongBan).Select(d => d.tenPhongBan).FirstOrDefault() ?? string.Empty
            }).GroupBy(s => new { s.MaPhongBan }).Select(g => new PhongBanModel
            {
                MaPhongBan = g.Key.MaPhongBan,
                Ten = linqDanhMuc.tbl_DM_PhongBans.Where(d => d.maPhongBan == g.Key.MaPhongBan).Select(d => d.tenPhongBan).FirstOrDefault() ?? string.Empty
            }).ToList();

            Session["ListNhanVien"] = nhanVienDC.OrderByDescending(d => d.CapBac).ToList();
            Session["ListPhongBan"] = listPhongBan.ToList();
            return Json("true");
        }
        public ActionResult LoadDanhSachNhanVien()
        {
            List<DanhSachNhanVienDaChonModel> NhanVienChon = Session["ListNhanVien"] as List<DanhSachNhanVienDaChonModel>;
            List<PhongBanModel> listPhongBan = Session["ListPhongBan"] as List<PhongBanModel>;
            Session.Remove("ListNhanVien");
            Session.Remove("ListPhongBan");
            ViewData["ListNhanVien"] = NhanVienChon;
            ViewData["ListPhongBan"] = listPhongBan;
            return PartialView("_ChonDanhSachNhanVien");
        }
        //B2. Click button Tiep tuc:
        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult ChonDanhSachNhanVien(FormCollection collection)
        {

            hr = new LinqNhanSuDataContext();
            var listOld = hr.tbl_NS_ChonNhanVienDanhGias.Where(d => d.maNguoiDanhGia == GetUser().manv).ToList();
            if (listOld != null && listOld.Count > 0)
            {
                hr.tbl_NS_ChonNhanVienDanhGias.DeleteAllOnSubmit(listOld);
            }
            List<tbl_NS_ChonNhanVienDanhGia> list = new List<tbl_NS_ChonNhanVienDanhGia>();
            tbl_NS_ChonNhanVienDanhGia nhanVien;
            string[] listNhanVien = collection.GetValues("maNhanVien");
            if (listNhanVien != null && listNhanVien.Length > 0)
            {
                for (int i = 0; i < listNhanVien.Length; i++)
                {
                    nhanVien = new tbl_NS_ChonNhanVienDanhGia();
                    nhanVien.maNguoiDanhGia = GetUser().manv;
                    nhanVien.maNhanVienChon = collection.GetValues("maNhanVien")[i];
                    list.Add(nhanVien);
                }
                hr.tbl_NS_ChonNhanVienDanhGias.InsertAllOnSubmit(list);
                hr.SubmitChanges();
            }

            PhieuKyNay();
            if (!string.IsNullOrEmpty(soPhieu))
            {
                return Json("edit");
            }
            else
            {
                return Json("create");
            }
        }
        //Load Form Create
        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            Session["BuocDanhGiaNew"] = "2";

            int thang = DateTime.Now.Month;
            ViewData["nam"] = DateTime.Now.Year;

            ViewData["quy"] = Convert.ToString(hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "KyDanhGia").Select(d => d.giaTri).FirstOrDefault());

            ViewData["maPhieuDanhGia"] = GenerateUtil.CheckLetter("DGTN", hr.tbl_NS_PhieuDanhGias.OrderByDescending(t => t.id).FirstOrDefault() != null ? hr.tbl_NS_PhieuDanhGias.OrderByDescending(t => t.id).FirstOrDefault().maPhieuDanhGia : "");
            //lấy ra tên người lập

            var DanhSachDaChon = hr.tbl_NS_ChonNhanVienDanhGias.Where(d => d.maNguoiDanhGia == GetUser().manv).ToList();
            List<DanhSachNhanVienDaChonModel> NhanVienChon = new List<DanhSachNhanVienDaChonModel>();
            if (DanhSachDaChon != null && DanhSachDaChon.Count > 0)
            {
                var listChon = (from p in hr.sp_NS_DanhSachNhanVienDaChon(GetUser().manv)
                                select new DanhSachNhanVienDaChonModel
                                {
                                    MaNhanVien = p.maNhanVien,
                                    TenNhanVien = p.hoTen,
                                    MaPhongBan = p.maPhongBan,
                                    TenPhongBan = p.tenPhongBan,
                                    MaChucDanh = p.maChucDanh,
                                    TenChucDanh = p.TenChucDanh,
                                    //Avatar = p.fileDinhKemAnhDaiDien,
                                    CapBac = hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maChucDanh == p.maChucDanh).Select(d => d.SoCapBac).FirstOrDefault(),
                                }).OrderByDescending(d => d.CapBac).ToList();
                NhanVienChon = listChon;
            }
            var CapBacUser = hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == GetUser().manv).FirstOrDefault();
            ViewData["CapBacUser"] = hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maChucDanh == CapBacUser.maChucDanh).Select(d => d.SoCapBac).FirstOrDefault();

            var listCapBac = NhanVienChon.Select(p => new DanhSachNhanVienDaChonModel
            {
                CapBac = p.CapBac,
            }).GroupBy(s => new { s.CapBac }).Select(g => new DanhSachNhanVienDaChonModel
            {
                CapBac = g.Key.CapBac,

            }).ToList();

            ViewData["lsCapBac"] = listCapBac;
            ViewData["lsNhanVienPhongBan"] = NhanVienChon;
            ViewData["lsDanhMucTieuChi"] = hr.tbl_NS_TieuChis.ToList();
            ViewData["tenNguoiLap"] = layRaTenNguoilap();
            return PartialView("_Create");
            //return View();
        }
        // POST: /DanhGiaTinNhiem/Create

        [AcceptVerbs(HttpVerbs.Post)]
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

                // TODO: Add insert logic here

                var maPhieuTinNhiem = GenerateUtil.CheckLetter("DGTN",
                hr.tbl_NS_PhieuDanhGias.OrderByDescending(t => t.id).FirstOrDefault() != null ? hr.tbl_NS_PhieuDanhGias.OrderByDescending(t => t.id).FirstOrDefault().maPhieuDanhGia : "");

                int qui = 0;

                //if (PhieuKyNay() == "1")
                //{
                //    qui = 1;
                //}
                //else
                //{
                //    qui = 2;
                //}
                qui = Convert.ToInt32(hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "KyDanhGia").Select(d => d.giaTri).FirstOrDefault());
                foreach (var i in hr.sp_NS_DanhSachNhanVienDaChon(GetUser().manv).ToList())
                //hr.sp_PhieuDanhGia_NhanVien(GetUser().MaNV).ToList())
                {
                    tbl_NS_PhieuDanhGia p_DanhGia = new tbl_NS_PhieuDanhGia();
                    p_DanhGia.maNhanVien = GetUser().manv;
                    p_DanhGia.maPhieuDanhGia = maPhieuTinNhiem;
                    p_DanhGia.nam = DateTime.Now.Year;

                    p_DanhGia.kyDanhGia = qui;
                    p_DanhGia.ngayLap = DateTime.Now;
                    p_DanhGia.maNhanVienDanhGia = i.maNhanVien;//i.maNhanVien;
                    p_DanhGia.trangThai = 0;
                    if (collection.Get("TrangThaiLuu") == "1")
                    {
                        p_DanhGia.trangThai = 1;
                    }
                    p_DanhGia.tongDiem = Convert.ToDouble(collection.GetValues(i.maNhanVien + "_diemTong")[0]);
                    p_DanhGia.heSoCap = 1;
                    if (!String.IsNullOrEmpty(collection.GetValues(i.maNhanVien + "_heSoCap")[0]))
                    {
                        p_DanhGia.heSoCap = Convert.ToInt32(collection.GetValues(i.maNhanVien + "_heSoCap")[0]);
                    }

                    p_DanhGia.nhanXet = collection.GetValues(i.maNhanVien + "_nhanXet")[0];
                    //p_DanhGia.tongDiem = Convert.ToDouble(collection.GetValues(i.maNhanVien + "_diemTong")[0]);
                    //p_DanhGia.heSoCap = Convert.ToInt32(collection.GetValues(i.maNhanVien + "_heSoCap")[0]);
                    hr.tbl_NS_PhieuDanhGias.InsertOnSubmit(p_DanhGia);
                    hr.SubmitChanges();
                    var idPhieuDanhGia = hr.tbl_NS_PhieuDanhGias.Where(d=>d.maPhieuDanhGia == maPhieuTinNhiem).OrderByDescending(t => t.ngayLap).Select(d => d.id).FirstOrDefault();
                    //Add List
                    tbl_NS_PhieuDanhGiaChiTiet danhGia_ChiTiet = null;
                    List<tbl_NS_PhieuDanhGiaChiTiet> lst_tbl_NS_PhieuDanhGiaChiTiet = new List<tbl_NS_PhieuDanhGiaChiTiet>();
                    foreach (var ct in hr.tbl_NS_TieuChis.ToList())
                    {
                        danhGia_ChiTiet = new tbl_NS_PhieuDanhGiaChiTiet();
                        danhGia_ChiTiet.maPhieuDanhGia = maPhieuTinNhiem;
                        danhGia_ChiTiet.idPhieuDanhGia = idPhieuDanhGia;
                        danhGia_ChiTiet.idTieuChi = ct.id;
                        var diemSo = collection.GetValues(i.maNhanVien + "_xepLoai_" + ct.id)[0];
                        //collection.GetValues(i.maNhanVien + "_xepLoai_" + ct.id)[0];
                        try
                        {
                            danhGia_ChiTiet.diemSo = Convert.ToInt32(diemSo.ToString());
                        }
                        catch { danhGia_ChiTiet.diemSo = 0; }
                        lst_tbl_NS_PhieuDanhGiaChiTiet.Add(danhGia_ChiTiet);
                       
                    }
                    hr.tbl_NS_PhieuDanhGiaChiTiets.InsertAllOnSubmit(lst_tbl_NS_PhieuDanhGiaChiTiet);
                    hr.SubmitChanges();

                }

                //return RedirectToAction("Edit", new { id = maPhieuTinNhiem });
                return Json(maPhieuTinNhiem);
            }
            catch
            {
                return View();
            }
        }
        public string layRaTenNguoilap()
        {

            if (GetUser().manv != null)
            {
                string MaNV = GetUser().manv.ToString();
                string tenNV = hr.tbl_NS_NhanViens.Where(t => t.maNhanVien == MaNV).Select(d => d.ho + " " + d.ten).FirstOrDefault();
                return tenNV;
            }
            return "";
        }
        public ActionResult UpdateBuocDuyet()
        {
            Session["BuocDanhGiaNew"] = "1";
            return Json(string.Empty);
        }

        // Khi bam nut luu tam
        // POST: /DanhGiaTinNhiem/Edit/5

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Edit(FormCollection collection)
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


                var maPhieuTinNhiemCu = collection.Get("maPhieuTinNhiem");
                var lsPhieu = hr.tbl_NS_PhieuDanhGias.Where(t => t.maPhieuDanhGia.Equals(maPhieuTinNhiemCu) && t.maNhanVien.Equals(GetUser().manv));
                var lsPhieuChiTiet = hr.tbl_NS_PhieuDanhGiaChiTiets.Where(t => t.maPhieuDanhGia.Equals(maPhieuTinNhiemCu));

                var maPhieuTinNhiem = GenerateUtil.CheckLetter("DGTN",
                hr.tbl_NS_PhieuDanhGias.OrderByDescending(t => t.id).FirstOrDefault() != null ? hr.tbl_NS_PhieuDanhGias.OrderByDescending(t => t.id).FirstOrDefault().maPhieuDanhGia : "");

                int qui = 1;

                //if (PhieuKyNay() == "1")
                //{
                //    qui = 1;
                //}
                //else
                //{
                //    qui = 2;
                //}
                qui = Convert.ToInt32(hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "KyDanhGia").Select(d => d.giaTri).FirstOrDefault());
                foreach (var i in hr.sp_NS_DanhSachNhanVienDaChon(GetUser().manv).ToList())
                {
                    tbl_NS_PhieuDanhGia p_DanhGia = new tbl_NS_PhieuDanhGia();
                    p_DanhGia.maNhanVien = GetUser().manv;
                    p_DanhGia.maPhieuDanhGia = maPhieuTinNhiem;
                    p_DanhGia.nam = DateTime.Now.Year;

                    p_DanhGia.kyDanhGia = qui;
                    p_DanhGia.ngayLap = DateTime.Now;
                    p_DanhGia.maNhanVienDanhGia = i.maNhanVien;
                    //i.maNhanVien;
                    p_DanhGia.trangThai = 0;
                    p_DanhGia.tongDiem = Convert.ToDouble(collection.GetValues(i.maNhanVien + "_diemTong")[0]);
                    p_DanhGia.heSoCap = 1;
                    if (!String.IsNullOrEmpty(collection.GetValues(i.maNhanVien + "_heSoCap")[0]))
                    {
                        p_DanhGia.heSoCap = Convert.ToInt32(collection.GetValues(i.maNhanVien + "_heSoCap")[0]);
                    }

                    p_DanhGia.nhanXet = collection.GetValues(i.maNhanVien + "_nhanXet")[0];
                    hr.tbl_NS_PhieuDanhGias.InsertOnSubmit(p_DanhGia);
                    hr.SubmitChanges();
                    var idPhieuDanhGia = hr.tbl_NS_PhieuDanhGias.Where(d=>d.maPhieuDanhGia == maPhieuTinNhiem).OrderByDescending(t => t.ngayLap).Select(d => d.id).FirstOrDefault();
                    
                    //Add List
                    tbl_NS_PhieuDanhGiaChiTiet danhGia_ChiTiet = null;
                    List<tbl_NS_PhieuDanhGiaChiTiet> lst_tbl_NS_PhieuDanhGiaChiTiet = new List<tbl_NS_PhieuDanhGiaChiTiet>();
                    foreach (var ct in hr.tbl_NS_TieuChis.ToList())
                    {
                        danhGia_ChiTiet = new tbl_NS_PhieuDanhGiaChiTiet();
                        danhGia_ChiTiet.maPhieuDanhGia = maPhieuTinNhiem;
                        danhGia_ChiTiet.idPhieuDanhGia = idPhieuDanhGia;
                        danhGia_ChiTiet.idTieuChi = ct.id;
                        var diemSo = collection.GetValues(i.maNhanVien + "_xepLoai_" + ct.id)[0];
                        //collection.GetValues(i.maNhanVien + "_xepLoai_" + ct.id)[0];
                        try
                        {
                            danhGia_ChiTiet.diemSo = Convert.ToInt32(diemSo.ToString());
                        }
                        catch { danhGia_ChiTiet.diemSo = 0; }
                        lst_tbl_NS_PhieuDanhGiaChiTiet.Add(danhGia_ChiTiet);

                    }
                    hr.tbl_NS_PhieuDanhGiaChiTiets.InsertAllOnSubmit(lst_tbl_NS_PhieuDanhGiaChiTiet);
                    hr.SubmitChanges();
                }
                hr.tbl_NS_PhieuDanhGiaChiTiets.DeleteAllOnSubmit(lsPhieuChiTiet);
                hr.tbl_NS_PhieuDanhGias.DeleteAllOnSubmit(lsPhieu);
                hr.SubmitChanges();

                return Json(maPhieuTinNhiem);
            }
            catch
            {
                return View();
            }
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
            Session["BuocDanhGiaNew"] = "3";

            if (GetUser() == null)
                return RedirectToAction("Index");

            int thang = DateTime.Now.Month;
            ViewData["nam"] = DateTime.Now.Year;
            //if (thang == 1 || thang == 2 || thang == 3)
            //{
            //    ViewData["quy"] = "I";
            //}
            //else if (thang == 4 || thang == 5 || thang == 6)
            //{
            //    ViewData["quy"] = "II";
            //}
            //else if (thang == 7 || thang == 8 || thang == 9)
            //{
            //    ViewData["quy"] = "III";
            //}
            //else
            //{
            //    ViewData["quy"] = "IV";
            //}
            ViewData["quy"] = Convert.ToString(hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "KyDanhGia").Select(d => d.giaTri).FirstOrDefault());
            //lấy ra tên người lập


            var trangThaiHoanThanh = "0";
            var TrangThai = hr.tbl_NS_PhieuDanhGias.Where(d => d.maPhieuDanhGia == id).FirstOrDefault();
            if (TrangThai.trangThai == 1)
            {
                trangThaiHoanThanh = "1";
            }
            ViewData["TrangThaiHoanThanh"] = trangThaiHoanThanh;
            var DanhSachDaChon = hr.tbl_NS_ChonNhanVienDanhGias.Where(d => d.maNguoiDanhGia == GetUser().manv).ToList();
            var DanhSachDaChonLuu = hr.tbl_NS_PhieuDanhGias.Where(t => t.maPhieuDanhGia.Equals(id)).ToList();
            List<DanhSachNhanVienDaChonModel> NhanVienChon = new List<DanhSachNhanVienDaChonModel>();
            if (trangThaiHoanThanh == "0")
            {
                if (DanhSachDaChon != null && DanhSachDaChon.Count > 0)
                {
                    var listChon = (from p in hr.sp_NS_DanhSachNhanVienDaChon(GetUser().manv)
                                    select new DanhSachNhanVienDaChonModel
                                    {
                                        MaNhanVien = p.maNhanVien,
                                        TenNhanVien = p.hoTen,
                                        MaPhongBan = p.maPhongBan,
                                        TenPhongBan = p.tenPhongBan,
                                        MaChucDanh = p.maChucDanh,
                                        TenChucDanh = p.TenChucDanh,
                                        //Avatar = p.fileDinhKemAnhDaiDien,
                                        CapBac = p.soCapBac
                                    }).OrderByDescending(d => d.CapBac).ToList();
                    NhanVienChon = listChon;
                }
            }
            else
            {
                if (DanhSachDaChonLuu != null && DanhSachDaChonLuu.Count > 0)
                {
                    var listChon = (from p in DanhSachDaChonLuu
                                    select new DanhSachNhanVienDaChonModel
                                    {
                                        MaNhanVien = p.maNhanVienDanhGia,
                                        TenNhanVien = hr.tbl_NS_NhanViens.Where(d => d.maNhanVien == p.maNhanVienDanhGia).Select(d => d.ho + " " + d.ten).FirstOrDefault(),
                                        TenPhongBan = hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == p.maNhanVienDanhGia).Select(d => d.tenPhongBan).FirstOrDefault(),
                                        CapBac = hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maChucDanh == hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(x => x.maNhanVien == p.maNhanVienDanhGia).Select(y => y.maChucDanh).FirstOrDefault()).Select(d => d.SoCapBac).FirstOrDefault(),
                                    }).OrderByDescending(d => d.CapBac).ToList();
                    NhanVienChon = listChon;
                }
            }
            var CapBacUser = hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == GetUser().manv).FirstOrDefault();
            ViewData["CapBacUser"] = hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maChucDanh == CapBacUser.maChucDanh).Select(d => d.SoCapBac).FirstOrDefault();

            var listCapBac = NhanVienChon.Select(p => new DanhSachNhanVienDaChonModel
            {
                CapBac = p.CapBac,
            }).GroupBy(s => new { s.CapBac }).Select(g => new DanhSachNhanVienDaChonModel
            {
                CapBac = g.Key.CapBac,

            }).ToList();
            ViewData["lsCapBac"] = listCapBac;
            ViewData["lsNhanVienPhongBan"] = NhanVienChon;
            ViewData["lsDanhMucTieuChi"] = hr.tbl_NS_TieuChis.ToList();
            ViewData["tenNguoiLap"] = layRaTenNguoilap();
            ViewData["lsDanhGiaChiTiet"] = hr.sp_NS_PhieuDanhGia_Edit(id, GetUser().manv).ToList();

            if (hr.sp_NS_PhieuDanhGia_Edit(id, GetUser().manv).Count() == 0)
            {
                //return RedirectToAction("Index");
            }
            else if (hr.sp_NS_PhieuDanhGia_Edit(id, GetUser().manv).FirstOrDefault().maNhanVien.Equals(GetUser().manv) == false)
            {
                //return RedirectToAction("Index");
            }
            //return View();
            return PartialView("_Edit");
        }
        public ActionResult HoanThanhDanhGia(string id)
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

                var lsTN = hr.tbl_NS_PhieuDanhGias.Where(t => t.maPhieuDanhGia.Equals(id)).ToList();
                foreach (var t in lsTN)
                {
                    var lsCT = hr.tbl_NS_PhieuDanhGias.Where(f => f.id == t.id).FirstOrDefault();
                    lsCT.trangThai = 1;
                    hr.SubmitChanges();
                }

                return Json(id);
            }
            catch
            {
                return Json("error");
            }
        }
        public ActionResult Details(string id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            if (GetUser() == null)
                return RedirectToAction("Index");

            var phieuNay = hr.tbl_NS_PhieuDanhGias.Where(t => t.maPhieuDanhGia.Equals(id)).FirstOrDefault();
            int thang = phieuNay.ngayLap.Value.Month;
            ViewData["nam"] = phieuNay.ngayLap.Value.Year;
            //if (thang == 1 || thang == 2 || thang == 3)
            //{
            //    ViewData["quy"] = "I";
            //}
            //else if (thang == 4 || thang == 5 || thang == 6)
            //{
            //    ViewData["quy"] = "II";
            //}
            //else if (thang == 7 || thang == 8 || thang == 9)
            //{
            //    ViewData["quy"] = "III";
            //}
            //else
            //{
            //    ViewData["quy"] = "IV";
            //}
            ViewData["quy"] = Convert.ToString(hr.tbl_NS_Cauhinhs.Where(d => d.maLoai == "KyDanhGia").Select(d => d.giaTri).FirstOrDefault());
            //lấy ra tên người lập

            List<DanhSachNhanVienDaChonModel> NhanVienChon = new List<DanhSachNhanVienDaChonModel>();
            var DanhSachDaChon = hr.tbl_NS_ChonNhanVienDanhGias.Where(d => d.maNguoiDanhGia == GetUser().manv).ToList();
            var DanhSachDaChonLuu = hr.tbl_NS_PhieuDanhGias.Where(t => t.maPhieuDanhGia.Equals(id)).ToList();
            var trangThaiHoanThanh = "0";
            if (phieuNay.trangThai == 1)
            {
                trangThaiHoanThanh = "1";
            }
            ViewData["TrangThaiHoanThanh"] = trangThaiHoanThanh;
            if (trangThaiHoanThanh == "0")
            {
                if (DanhSachDaChon != null && DanhSachDaChon.Count > 0)
                {
                    var listChon = (from p in hr.sp_NS_DanhSachNhanVienDaChon(GetUser().manv)
                                    select new DanhSachNhanVienDaChonModel
                                    {
                                        MaNhanVien = p.maNhanVien,
                                        TenNhanVien = p.hoTen,
                                        MaPhongBan = p.maPhongBan,
                                        TenPhongBan = p.tenPhongBan,
                                        MaChucDanh = p.maChucDanh,
                                        TenChucDanh = p.TenChucDanh,

                                        CapBac = p.soCapBac
                                    }).OrderByDescending(d => d.CapBac).ToList();
                    NhanVienChon = listChon;
                }
            }
            else
            {
                if (DanhSachDaChonLuu != null && DanhSachDaChonLuu.Count > 0)
                {
                    var listChon = (from p in DanhSachDaChonLuu
                                    select new DanhSachNhanVienDaChonModel
                                    {
                                        MaNhanVien = p.maNhanVienDanhGia,
                                        TenNhanVien = hr.tbl_NS_NhanViens.Where(d => d.maNhanVien == p.maNhanVienDanhGia).Select(d => d.ho + " " + d.ten).FirstOrDefault(),
                                        TenPhongBan = hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == p.maNhanVienDanhGia).Select(d => d.tenPhongBan).FirstOrDefault(),
                                        CapBac = hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maChucDanh == hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(x => x.maNhanVien == p.maNhanVienDanhGia).Select(y => y.maChucDanh).FirstOrDefault()).Select(d => d.SoCapBac).FirstOrDefault(),
                                        MaChucDanh = hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maChucDanh == hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(x => x.maNhanVien == p.maNhanVienDanhGia).Select(y => y.maChucDanh).FirstOrDefault()).Select(d => d.maChucDanh).FirstOrDefault(),
                                        TenChucDanh = hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maChucDanh == hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(x => x.maNhanVien == p.maNhanVienDanhGia).Select(y => y.maChucDanh).FirstOrDefault()).Select(d => d.TenChucDanh).FirstOrDefault(),
                                    }).OrderByDescending(d => d.CapBac).ToList();
                    NhanVienChon = listChon;
                }
            }
            var CapBacUser = hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == GetUser().manv).FirstOrDefault();
            ViewData["CapBacUser"] = hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maChucDanh == CapBacUser.maChucDanh).Select(d => d.SoCapBac).FirstOrDefault();

            var listCapBac = NhanVienChon.Select(p => new DanhSachNhanVienDaChonModel
            {
                CapBac = p.CapBac,
            }).GroupBy(s => new { s.CapBac }).Select(g => new DanhSachNhanVienDaChonModel
            {
                CapBac = g.Key.CapBac,

            }).ToList();
            ViewData["lsCapBac"] = listCapBac;
            ViewData["lsNhanVienPhongBan"] = NhanVienChon;
            ViewData["lsDanhMucTieuChi"] = hr.tbl_NS_TieuChis.ToList();
            ViewData["tenNguoiLap"] = layRaTenNguoilap();
            ViewData["lsDanhGiaChiTiet"] = hr.sp_NS_PhieuDanhGia_Edit(id, GetUser().manv).ToList();

            if (hr.sp_NS_PhieuDanhGia_Edit(id, GetUser().manv).Count() == 0)
            {
                return RedirectToAction("Index");
            }
            else if (hr.sp_NS_PhieuDanhGia_Edit(id, GetUser().manv).FirstOrDefault().maNhanVien.Equals(GetUser().manv) == false)
            {
                return RedirectToAction("Index");
            }
            return View();
        }
        public ActionResult XemKetQua()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            if (GetUser() == null)
                return RedirectToAction("Index");


            var lists = hr.sp_NS_KetQuaDanhGiaNhanVien(GetUser().manv).ToList();
            ViewData["lists"] = lists;

            return View();
        }
        public ActionResult BaoCaoXepLoaiNhanVien(string maNV, int?nam)
        {
            #region Role user
            permission = GetPermission("BCKQDGTN", BangPhanQuyen.QuyenDuyet);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            var listChiTiet = hr.sp_NS_KetQuaDanhGiaNhanVien_ChiTiet(maNV,nam);
            ViewData["ThongTinNhanVien"] = hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maNhanVien == maNV).FirstOrDefault();
            ViewData["ListNhanVien"] = listChiTiet.ToList();
            return View();
        }


        public ActionResult BaoCaoXepLoai(string maPhongBan, int? mucLevel, int? qui, int? nam, string qSearch, int? page, int? pageSize)
        {
            #region Role user
            permission = GetPermission("BCKQDGTN", BangPhanQuyen.QuyenDuyet);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            int currentPageIndex = page.HasValue ? page.Value : 1;
            pageSize = pageSize ?? 30;
            int? tongSoDong = 0;


            //
            buildTree = new StringBuilder();
            phongBans = linqDanhMuc.tbl_DM_PhongBans.ToList();
            buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
            ViewBag.PhongBans = buildTree.ToString();
            //
            try
            {
                ViewBag.Count = nhanViens[0].tongSoDong;
                tongSoDong = nhanViens[0].tongSoDong;
            }
            catch
            {
                ViewBag.Count = 0;
            }
            int? NewNam = nam == null ? DateTime.Now.Year : Convert.ToInt32(nam);
            if (qui == null)
            {
                qui = 1;
            }
            ViewData["qui"] = qui;
            ViewData["nam"] = nam;
            ViewData["mucLevel"] = mucLevel;
            ViewData["maPhongBan"] = maPhongBan;
            ViewData["tenPhongBan"] = hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maPhongBan == maPhongBan).Select(d => d.tenPhongBan).FirstOrDefault();
            ViewData["qSearch"] = qSearch;



            ViewData["yearWorkings"] = new SelectList(GetYearLimits((int)NewNam, 5), NewNam);

            var listLevel = hr.Sys_ChucDanhs.ToList();
            Dictionary<string, string> capBac = new Dictionary<string, string>();
            capBac.Add("", "--Tất cả--");
            foreach (var item in listLevel.OrderByDescending(d => d.SoCapBac))
            {

                capBac.Add(item.MaChucDanh, item.TenChucDanh);
            }
            ViewData["ListCapBac"] = new SelectList(capBac, "Key", "Value", mucLevel);
            return View("");
        }
        public ActionResult BaoCaoXepLoaiNam(int? nam)
        {
            #region Role user
            permission = GetPermission("BCXLNam", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion


            int? NewNam = nam == null ? DateTime.Now.Year : Convert.ToInt32(nam);

            ViewData["yearWorkings"] = new SelectList(GetYearLimits((int)NewNam, 5), NewNam);
            var listLoaiDGs = linqDanhMuc.tbl_DM_XepLoaiDGs.OrderBy(d => d.thuTu).ToList();
            ViewBag.ListLoaiDGs = listLoaiDGs;
            // Select ma nhom danh gia
            var listNhomDanhGiaNew = linqDanhMuc.sp_DG_XepLoaiHang(nam).ToList();
            ViewBag.ListNhomDanhGiasNews = listNhomDanhGiaNew;
            // End select ma nhom danh gia
            return View("");
        }
        public ActionResult LoadBaoCaoXepLoaiNam(int? nam)
        {
            #region Role user
            permission = GetPermission("BCXLNam", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            // Check voi nam nay da tao bao cao xep loai nam chua
            var listNam = linqDanhMuc.tbl_XepLoaiTongs.Where(d => d.nam == nam).ToList();
            List<tbl_XepLoaiTong> lst_tblXepLoaiTong = new List<tbl_XepLoaiTong>();

            var listLoai2 = linqDanhMuc.tbl_DM_XepLoaiDGs.ToList();
            foreach (var item in listLoai2)
            {
                var chek = linqDanhMuc.tbl_XepLoaiTongs.Where(d => d.nam == nam && d.maLoai == item.maXepLoai).FirstOrDefault();
                if (chek == null)
                {
                    tbl_XepLoaiTong tblXepLoaiTong = new tbl_XepLoaiTong();
                    tblXepLoaiTong.nam = nam;
                    tblXepLoaiTong.maLoai = item.maXepLoai;
                    tblXepLoaiTong.soLuong = 0;
                    tblXepLoaiTong.tiLe = 0;
                    lst_tblXepLoaiTong.Add(tblXepLoaiTong);
                }
            }
            linqDanhMuc.tbl_XepLoaiTongs.InsertAllOnSubmit(lst_tblXepLoaiTong);
            linqDanhMuc.SubmitChanges();




            var listXL = linqDanhMuc.sp_DG_XepLoaiTong(nam).ToList();
            var toTal = listXL.Count();
            ViewBag.listXLTong = listXL;
            ViewBag.toTal = toTal;
            // end check

            // Ti le khen thuong theo xep hang
            var lstXepLoaiHang = linqDanhMuc.sp_DG_XepLoaiHang(nam).ToList();
            List<tbl_XepLoaiTheoHang> lst_tblXepLoaiTheoHang = new List<tbl_XepLoaiTheoHang>();
            foreach (var itemHang in lstXepLoaiHang)
            {
                var lstDanhGia = linqDanhMuc.tbl_DM_XepLoaiDGs.ToList();
                foreach (var itemLDG in lstDanhGia)
                {
                    var checkLDG = linqDanhMuc.tbl_XepLoaiTheoHangs.Where(d => d.nam == nam && d.thuTuHang == Convert.ToDouble(itemHang.thuHang ?? 0)).FirstOrDefault();
                    if (checkLDG == null)
                    {
                        tbl_XepLoaiTheoHang tblXepLoaiTheohang = new tbl_XepLoaiTheoHang();
                        tblXepLoaiTheohang.maLoai = itemLDG.maXepLoai;
                        tblXepLoaiTheohang.nam = Convert.ToInt32(nam);
                        tblXepLoaiTheohang.thuTuHang = Convert.ToDouble(itemHang.thuHang);
                        tblXepLoaiTheohang.soLuongNV = Convert.ToDouble(itemHang.sL);
                        tblXepLoaiTheohang.tiLeChoi = 0;
                        tblXepLoaiTheohang.phanTram = 0;
                        tblXepLoaiTheohang.soLuong = 0;
                        lst_tblXepLoaiTheoHang.Add(tblXepLoaiTheohang);
                    }

                }

            }
            linqDanhMuc.tbl_XepLoaiTheoHangs.InsertAllOnSubmit(lst_tblXepLoaiTheoHang);
            linqDanhMuc.SubmitChanges();
            // End ti le khen thuong theo xep hang
            //Load loại theo hạng
            var lstTheoHangIndex = linqDanhMuc.sp_DG_XepLoaiTheoHang_Index(nam).ToList();
            ViewBag.totalHang = lstTheoHangIndex.Count();
            ViewBag.lstTheoHangIndex = lstTheoHangIndex;
            // Load xếp loại theo hạng
            var listXLTheoBoPhan = linqDanhMuc.sp_DG_XepLoaiTheoBoPhan(nam).ToList();
            ViewBag.listXLTheoBoPhan = listXLTheoBoPhan;
            // End xep loai theo hạng
            // Select ma nhom danh gia
            var listNhomDanhGiaNew = linqDanhMuc.sp_DG_XepLoaiHang(nam).ToList();
            ViewBag.ListNhomDanhGiasNews = listNhomDanhGiaNew;
            // End select ma nhom danh gia
            int? NewNam = nam == null ? DateTime.Now.Year : Convert.ToInt32(nam);


            ViewData["yearWorkings"] = new SelectList(GetYearLimits((int)NewNam, 5), NewNam);


            return PartialView("_LoadBaoCaoXepLoaiNam");
        }
        public ActionResult LoadBaoCaoChiTietNhanVienPB(int? nam, string maNhomPhongBan, string maNhomDanhGia, string qSearch, int pageIndex, int pageSize)
        {
            #region Role user
            permission = GetPermission("BCXLNam", BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            // Check voi nam nay da tao bao cao xep loai nam chua
            try
            {

                var list = linqDanhMuc.sp_NS_BaoCaoDanhGiaXepLoai_Index(nam, maNhomPhongBan, maNhomDanhGia, qSearch, pageIndex, pageSize).ToList();


                if (list.Count == 0)
                {
                    list = null;
                }

                return Json(new { count = list.Count(), lists = list });

            }
            catch
            {
                return Json(string.Empty);
            }
        }

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult SaveKhenThuongTH(FormCollection form)
        {

            #region Role user
            permission = GetPermission("BCXLNam", BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            try
            {

                // Ti le khen thuong theo xep hang
                var deleteThuHang = linqDanhMuc.tbl_XepLoaiTheoHangs.Where(d => d.nam == Convert.ToInt32(form.Get("namThuHang"))).ToList();
                if (deleteThuHang != null)
                {
                    var tongSoThuHang = deleteThuHang.Count();
                    linqDanhMuc.tbl_XepLoaiTheoHangs.DeleteAllOnSubmit(deleteThuHang);
                    linqDanhMuc.SubmitChanges();
                }
                // Create New;
                tbl_XepLoaiTheoHang tblXepLoaiTheoHang = null;
                List<tbl_XepLoaiTheoHang> lst_tblXepLoaiTheoHang = new List<tbl_XepLoaiTheoHang>();

                string[] thuHang = form.GetValues("thuTuHang");


                if (thuHang != null)
                {
                    for (int i = 0; i < thuHang.Length; i++)
                    {

                        tblXepLoaiTheoHang = new tbl_XepLoaiTheoHang();
                        tblXepLoaiTheoHang.nam = Convert.ToInt32(form.Get("namThuHang"));
                        tblXepLoaiTheoHang.thuTuHang = Convert.ToDouble(form.GetValues("thuTuHang")[i]);
                        tblXepLoaiTheoHang.soLuongNV = Convert.ToInt32(form.GetValues("soLuongNV")[i]);
                        tblXepLoaiTheoHang.tiLeChoi = Convert.ToDouble(form.GetValues("tiLeChoi")[i]);
                        tblXepLoaiTheoHang.soLuong = Convert.ToDouble(form.GetValues("soLuongHang")[i]);
                        tblXepLoaiTheoHang.phanTram = Convert.ToDouble(form.GetValues("tilephantram")[i]);
                        tblXepLoaiTheoHang.maLoai = form.GetValues("maLoaiHang")[i];



                        lst_tblXepLoaiTheoHang.Add(tblXepLoaiTheoHang);

                    }
                    linqDanhMuc.tbl_XepLoaiTheoHangs.InsertAllOnSubmit(lst_tblXepLoaiTheoHang);
                }

                linqDanhMuc.SubmitChanges();
                // End ti le khen thuong theo xep hang
                return RedirectToAction("BaoCaoXepLoaiNam");


            }
            catch
            {
                return View();
            }
        }
        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult SaveXepLoaiTheoBP(FormCollection form)
        {

            #region Role user
            permission = GetPermission("BCXLNam", BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            try
            {

                // Ti le khen thuong theo xep hang
                var deleteThuHang = linqDanhMuc.tbl_XepLoaiTheoBoPhans.Where(d => d.nam == Convert.ToInt32(form.Get("namThuHangXLBP"))).ToList();
                if (deleteThuHang != null)
                {
                    var tongSoThuHang = deleteThuHang.Count();
                    linqDanhMuc.tbl_XepLoaiTheoBoPhans.DeleteAllOnSubmit(deleteThuHang);
                    linqDanhMuc.SubmitChanges();
                }
                // Create New;
                tbl_XepLoaiTheoBoPhan tblXepLoaiTheoHang = null;
                List<tbl_XepLoaiTheoBoPhan> lst_tblXepLoaiTheoHang = new List<tbl_XepLoaiTheoBoPhan>();

                string[] thuHang = form.GetValues("thuTuHangXLBP");


                if (thuHang != null)
                {
                    for (int i = 0; i < thuHang.Length; i++)
                    {

                        tblXepLoaiTheoHang = new tbl_XepLoaiTheoBoPhan();
                        tblXepLoaiTheoHang.nam = Convert.ToInt32(form.Get("namThuHangXLBP"));
                        tblXepLoaiTheoHang.thuTuHang = Convert.ToDouble(form.GetValues("thuTuHangXLBP")[i]);
                        tblXepLoaiTheoHang.maPhongBan = form.GetValues("maPhongBanXLBP")[i];
                        tblXepLoaiTheoHang.maLoai = form.GetValues("maLoaiHangXLBP")[i];
                        tblXepLoaiTheoHang.soLuong = Convert.ToDouble(form.GetValues("soLuongXLBP")[i]);



                        lst_tblXepLoaiTheoHang.Add(tblXepLoaiTheoHang);

                    }
                    linqDanhMuc.tbl_XepLoaiTheoBoPhans.InsertAllOnSubmit(lst_tblXepLoaiTheoHang);
                }

                linqDanhMuc.SubmitChanges();
                // End ti le khen thuong theo xep hang
                // Call store sp_BaoCaoTongHopDanhGia2017;
                int namTongHop = string.IsNullOrEmpty(form.Get("namThuHangXLBP")) ? 0 : Convert.ToInt32(form.Get("namThuHangXLBP"));
                linqDanhMuc.sp_BaoCaoTongHopDanhGia2017(namTongHop).ToString();
                // End call sp_BaoCaoTongHopDanhGia2017
                return RedirectToAction("BaoCaoXepLoaiNam");


            }
            catch
            {
                return View();
            }
        }
        public string UpdateMucLuongThuong(int? nam, string soLuongs, string maLoais)
        {
            #region Role user
            permission = GetPermission("BCXLNam", BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return "Lỗi";
            if (!permission.Value)
                return "Lỗi";
            #endregion
            try
            {
                 var checkEx = hr.tbl_DuyetXepLoaiNVs.Where(t => t.nam == nam).FirstOrDefault();
                 if (checkEx == null)
                 {
                     // Update
                     string[] SoLuong = soLuongs.Split(',');
                     string[] MaLoai = maLoais.Split(',');

                     if (SoLuong != null)
                     {
                         for (int i = 0; i < SoLuong.Length; i++)
                         {
                             tbl_XepLoaiTong mucLuongThuong = new tbl_XepLoaiTong();
                             mucLuongThuong = linqDanhMuc.tbl_XepLoaiTongs.Where(d => d.maLoai == MaLoai[i] && d.nam == nam).FirstOrDefault();
                             mucLuongThuong.maLoai = MaLoai[i];
                             mucLuongThuong.nam = mucLuongThuong.nam;
                             mucLuongThuong.soLuong = Convert.ToDouble(SoLuong[i]);
                             mucLuongThuong.tiLe = mucLuongThuong.tiLe;
                             linqDanhMuc.SubmitChanges();

                         }

                     }



                     return "true";
                 }
                 return "false";


            }
            catch (Exception e)
            {
                return "Lỗi: " + e.Message;
            }
        }
        public string UpdateTiLe(int? nam, string tiLes, string maLoais)
        {
            #region Role user
            permission = GetPermission("BCXLNam", BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return "Lỗi";
            if (!permission.Value)
                return "Lỗi";
            #endregion

            try
            {
                 var checkEx = hr.tbl_DuyetXepLoaiNVs.Where(t => t.nam == nam).FirstOrDefault();
                 if (checkEx == null)
                 {
                     // Update
                     string[] TiLe = tiLes.Split(',');
                     string[] MaLoai = maLoais.Split(',');

                     if (TiLe != null)
                     {
                         for (int i = 0; i < TiLe.Length; i++)
                         {
                             tbl_XepLoaiTong mucLuongThuong = new tbl_XepLoaiTong();
                             mucLuongThuong = linqDanhMuc.tbl_XepLoaiTongs.Where(d => d.maLoai == MaLoai[i] && d.nam == nam).FirstOrDefault();
                             mucLuongThuong.maLoai = MaLoai[i];
                             mucLuongThuong.nam = mucLuongThuong.nam;
                             mucLuongThuong.soLuong = mucLuongThuong.soLuong;
                             mucLuongThuong.tiLe = Convert.ToDouble(TiLe[i]);
                             linqDanhMuc.SubmitChanges();

                         }

                     }




                     return "true";
                 }
                 return "false";


            }
            catch (Exception e)
            {
                return "Lỗi: " + e.Message;
            }
        }
        public string UpdateDieuChinhXL(int? nam, string maNhanVien, string maXLDieuChinh)
        {
            #region Role user
            permission = GetPermission("BCXLNam", BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return "Lỗi";
            if (!permission.Value)
                return "Lỗi";
            #endregion

            try
            {
                var checkEx = hr.tbl_DuyetXepLoaiNVs.Where(t => t.nam == nam).FirstOrDefault();
                if (checkEx == null)
                {
                    // Update



                    tbl_NS_KetQuaDanhGiaXepLoai KetQuaDanhGiaXepLoai = new tbl_NS_KetQuaDanhGiaXepLoai();
                    KetQuaDanhGiaXepLoai = linqDanhMuc.tbl_NS_KetQuaDanhGiaXepLoais.Where(d => d.maNhanVien == maNhanVien && d.nam == nam).FirstOrDefault();
                    if (KetQuaDanhGiaXepLoai != null)
                    {
                        KetQuaDanhGiaXepLoai.dieuChinhXL = maXLDieuChinh;
                        linqDanhMuc.SubmitChanges();
                    }


                    SaveActiveHistory("Điều chỉnh XLNV, Mã NV: " + maNhanVien + " năm: " + nam + " điều chỉnh sang: " + maXLDieuChinh);

                    return "true";
                }
                return "false";


            }
            catch (Exception e)
            {
                return "Lỗi: " + e.Message;
            }
        }

        public string DuyetXepLoaiNV(int? nam)
        {
            #region Role user
            permission = GetPermission("BCXLNam", BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return "Lỗi";
            if (!permission.Value)
                return "Lỗi";
            #endregion
           
            try
            {

                // Update

                //Check 
                var result = new { kq = false };
                var checkEx = hr.tbl_DuyetXepLoaiNVs.Where(t => t.nam == nam).FirstOrDefault();
                if (checkEx != null)
                {
                    return "false";
                }
                //End check
                // Insert Row 
                tbl_DuyetXepLoaiNV tblDuyetBL = new tbl_DuyetXepLoaiNV();
                tblDuyetBL.nam = nam??0;
               
                tblDuyetBL.ngayDuyet = DateTime.Now;
                tblDuyetBL.nguoiDuyet = GetUser().manv;
                hr.tbl_DuyetXepLoaiNVs.InsertOnSubmit(tblDuyetBL);
                hr.SubmitChanges();
                // End Insert Row
                // Check Exist
                SaveActiveHistory("Duyệt báo cáo xếp loại NV năm: " + nam);
                return "true";


            }
            catch (Exception e)
            {
                return "Lỗi: " + e.Message;
            }
        }

        public string CheckDuyetXepLoaiNV(int? nam)
        {
            #region Role user
            permission = GetPermission("BCXLNam", BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return "Lỗi";
            if (!permission.Value)
                return "Lỗi";
            #endregion

            try
            {

                // Update

                //Check 
              
                var checkEx = hr.tbl_DuyetXepLoaiNVs.Where(t => t.nam == nam).FirstOrDefault();
                if (checkEx != null)
                {
                    return "true";
                }
               
                return "false";


            }
            catch (Exception e)
            {
                return "Lỗi: " + e.Message;
            }
        }

        public string ResetTiLeKhenThuong(int? nam)
        {
            #region Role user
            permission = GetPermission("BCXLNam", BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return "Lỗi";
            if (!permission.Value)
                return "Lỗi";
            #endregion
            try
            {
                // Delete truoc khi insert new
                var delXLTh = linqDanhMuc.tbl_XepLoaiTheoHangs.Where(d => d.nam == nam).ToList();
                if (delXLTh != null)
                {
                    linqDanhMuc.tbl_XepLoaiTheoHangs.DeleteAllOnSubmit(delXLTh);
                    linqDanhMuc.SubmitChanges();
                }
                // End delete truoc khi insert
                // Ti le khen thuong theo xep hang
                var lstXepLoaiHang = linqDanhMuc.sp_DG_XepLoaiHang(nam).ToList();
                List<tbl_XepLoaiTheoHang> lst_tblXepLoaiTheoHang = new List<tbl_XepLoaiTheoHang>();
                foreach (var itemHang in lstXepLoaiHang)
                {
                    var lstDanhGia = linqDanhMuc.tbl_DM_XepLoaiDGs.ToList();
                    foreach (var itemLDG in lstDanhGia)
                    {
                        var checkLDG = linqDanhMuc.tbl_XepLoaiTheoHangs.Where(d => d.nam == nam && d.thuTuHang == Convert.ToDouble(itemHang.thuHang ?? 0)).FirstOrDefault();
                        if (checkLDG == null)
                        {
                            tbl_XepLoaiTheoHang tblXepLoaiTheohang = new tbl_XepLoaiTheoHang();
                            tblXepLoaiTheohang.maLoai = itemLDG.maXepLoai;
                            tblXepLoaiTheohang.nam = Convert.ToInt32(nam);
                            tblXepLoaiTheohang.thuTuHang = Convert.ToDouble(itemHang.thuHang);
                            tblXepLoaiTheohang.soLuongNV = Convert.ToDouble(itemHang.sL);
                            tblXepLoaiTheohang.tiLeChoi = 0;
                            tblXepLoaiTheohang.phanTram = 0;
                            tblXepLoaiTheohang.soLuong = 0;
                            lst_tblXepLoaiTheoHang.Add(tblXepLoaiTheohang);
                        }

                    }

                }
                linqDanhMuc.tbl_XepLoaiTheoHangs.InsertAllOnSubmit(lst_tblXepLoaiTheoHang);
                linqDanhMuc.SubmitChanges();
                // End ti le khen thuong theo xep hang
                // Ti le khen thuong theo xep hang
                var deleteThuHang = linqDanhMuc.tbl_XepLoaiTheoBoPhans.Where(d => d.nam == Convert.ToInt32(nam)).ToList();
                if (deleteThuHang != null)
                {
                    var tongSoThuHang = deleteThuHang.Count();
                    linqDanhMuc.tbl_XepLoaiTheoBoPhans.DeleteAllOnSubmit(deleteThuHang);
                    linqDanhMuc.SubmitChanges();
                }
                SaveActiveHistory("Khởi tạo lại giá trị khen thưởng năm: " + nam);
                return "true";
            }
            catch (Exception e)
            {
                return "Lỗi: " + e.Message;
            }
        }
        public string TinhLuongNVTheoXepLoai(int? nam)
        {
            #region Role user
            permission = GetPermission("BCXLNam", BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return "Lỗi";
            if (!permission.Value)
                return "Lỗi";
            #endregion

            try
            {

                // Update

                //Check 
              
                var checkEx = hr.sp_NS_TinhLuongThuong(nam).ToString();
             
                SaveActiveHistory("Tính lương thưởng theo xếp loại NV năm: " + nam);
                return "true";


            }
            catch (Exception e)
            {
                return "Lỗi: " + e.Message;
            }
        }
        public ActionResult LoadBaoCaoXepLoai(string maPhongBan, int? mucLevel, int? qui, int? nam, string qSearch, int _page = 0)
        {
            #region Role user
            permission = GetPermission("BCKQDGTN", BangPhanQuyen.QuyenDuyet);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion


            int? NewNam = nam == null ? DateTime.Now.Year : Convert.ToInt32(nam);
            if (qui == null)
            {
                qui = 1;
            }
            int page = _page == 0 ? 1 : _page;
            int pIndex = page;
            int total = hr.sp_NS_BaoCaoDanhGiaNhanVien(qui, NewNam, maPhongBan, mucLevel, qSearch).Count();
            PagingLoaderController("/DanhGiaTinNhiem/BaoCaoXepLoai/", total, page, "?qsearch=" + qSearch + "&qui=" + qui + "&NewNam=" + nam + "&maPhongBan=" + maPhongBan + "&mucLevel=" + mucLevel);
            ViewData["lsDanhSach"] = hr.sp_NS_BaoCaoDanhGiaNhanVien(qui, NewNam, maPhongBan, mucLevel, qSearch).Skip(start).Take(offset).ToList();

            ViewData["qui"] = qui;
            ViewData["nam"] = nam;
            ViewData["mucLevel"] = mucLevel;
            ViewData["maPhongBan"] = maPhongBan;
            ViewData["tenPhongBan"] = hr.vw_NS_DanhSachNhanVienTheoPhongBans.Where(d => d.maPhongBan == maPhongBan).Select(d => d.tenPhongBan).FirstOrDefault();
            ViewData["qSearch"] = qSearch;



            ViewData["yearWorkings"] = new SelectList(GetYearLimits((int)NewNam, 5), NewNam);

            var listLevel = hr.Sys_ChucDanhs.ToList();
            Dictionary<string, string> capBac = new Dictionary<string, string>();
            capBac.Add("", "--Tất cả--");
            foreach (var item in listLevel.OrderByDescending(d => d.SoCapBac))
            {

                capBac.Add(item.MaChucDanh, item.TenChucDanh);
            }
            ViewData["ListCapBac"] = new SelectList(capBac, "Key", "Value", mucLevel);
            return PartialView("_LoadBaoCaoXepLoai");
        }
        public static List<int> GetYearLimits(int currentYear, int limmit)
        {
            List<int> years = new List<int>();
            for (int i = currentYear - limmit; i <= currentYear + limmit; i++)
            {
                years.Add(i);
            }

            return years;
        }
        private void kyDanhGias(int value)
        {
            Dictionary<int, string> dics = new Dictionary<int, string>();
            for (int i = 1; i < 5; i++)
            {
                dics[i] = i.ToString();
            }
            ViewBag.KyDanhGias = new SelectList(dics, "Key", "Value", value);

        }
    }
}
