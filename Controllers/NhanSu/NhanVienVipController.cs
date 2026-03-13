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
using BatDongSan.Models.VIP;
using BatDongSan.Utils.Paging;

namespace BatDongSan.Controllers.NhanSu
{
    public class NhanVienVipController : ApplicationController
    {
        LinqVIPDataContext linqVIP = new LinqVIPDataContext();
        LinqDanhMucDataContext linqDanhMuc = new LinqDanhMucDataContext();
        LinqHeThongDataContext lqHT = new LinqHeThongDataContext();
        BatDongSan.Models.NhanSu.LinqNhanSuDataContext context = new BatDongSan.Models.NhanSu.LinqNhanSuDataContext();
        IList<sp_NhanVienVip_IndexResult> nhanVienVips;
        NhanVienVIPModel modelNV;
        BatDongSan.Models.NhanSu.HopDongLaoDongModel model;
        BatDongSan.Models.VIP.tbl_NhanVienVIP choViec;
        LinqDanhMucDataContext linqDM = new LinqDanhMucDataContext();
        IList<tbl_DM_PhongBan> phongBans;
        StringBuilder buildTree = null;
        readonly string MCV = "NhanVienVip";
        bool? permission;

        public ActionResult Index(int? page, int? pageSize, string loaiHopDong, string searchString, string maPhongBan)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            BindDataTrangThai(MCV);
            buildTree = new StringBuilder();
            phongBans = linqDM.tbl_DM_PhongBans.ToList();
            buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
            ViewBag.PhongBans = buildTree.ToString();
            //
            int currentPageIndex = page.HasValue ? page.Value : 1;
            pageSize = pageSize ?? 200;

            var loaiHopDongs = context.tbl_DM_LoaiHopDongLaoDongs.ToList();
            loaiHopDongs.Insert(0, new BatDongSan.Models.NhanSu.tbl_DM_LoaiHopDongLaoDong { maLoaiHopDong = "", tenLoaiHopDong = "[Tất cả loại hợp đồng]" });
            ViewBag.LoaiHopDongs = new SelectList(loaiHopDongs, "maLoaiHopDong", "tenLoaiHopDong");
            nhanVienVips = linqVIP.sp_NhanVienVip_Index(searchString, currentPageIndex, 200).ToList();

            int? tongSoDong = 0;
            try
            {
                ViewBag.Count = nhanVienVips[0].tongSoDong;
                tongSoDong = nhanVienVips[0].tongSoDong;
            }
            catch
            {
                ViewBag.Count = 0;
            }
            if (Request.IsAjaxRequest())
            {
                return PartialView("PartialIndex", nhanVienVips.ToPagedList(currentPageIndex, 200, true, tongSoDong));
            }
            ViewBag.searchString = searchString;
            ViewBag.loaiHopDong = loaiHopDong;
            ViewBag.maPhongBan = maPhongBan;
            return View(nhanVienVips.ToPagedList(currentPageIndex, 200, true, tongSoDong));
        }

        #region Create, Edit, Details

        // GET: /HopDongLaoDong/Create

        public ActionResult Create()
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
                thang(DateTime.Now.Month);
                nam(DateTime.Now.Year);
                modelNV = new NhanVienVIPModel();
                modelNV.maNhanVien = IdGenerator();
                var chiNhanhNganHangs = context.tbl_DM_ChiNhanhNganHangs.ToList();
                chiNhanhNganHangs.Insert(0, new BatDongSan.Models.NhanSu.tbl_DM_ChiNhanhNganHang { maChiNhanh = "", tenChiNhanh = "" });
                ViewBag.ChiNhanhNHs = new SelectList(chiNhanhNganHangs, "maChiNhanh", "tenChiNhanh", modelNV.maChiNhanhNganHang);
                modelNV.ngayLap = DateTime.Now;
              
                return View(modelNV);
            }
            catch
            {
                return View("Error");
            }
        }

        // GET: /HopDongLaoDong/Create

        public ActionResult Edit(string id)
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
                
                var listChoViec = linqVIP.tbl_NhanVienVIPs.Where(d => d.MaNhanVien == id).FirstOrDefault();
                modelNV = new NhanVienVIPModel();

                modelNV.maNhanVien = listChoViec.MaNhanVien;
                modelNV.tenNhanVien = listChoViec.TenNhanVien;
                modelNV.email = listChoViec.Email;
                modelNV.ngayVaoLam = listChoViec.NgayVaoLam;
                modelNV.ngayLap = listChoViec.ngayLap;
                modelNV.ghiChu = listChoViec.GhiChu;
                modelNV.luongCoBan = (decimal?)listChoViec.LuongCoBan ?? 0;
                modelNV.tongLuong = (decimal?)listChoViec.TongLuong ?? 0;
                modelNV.thoiGianCongTac = (float?)listChoViec.ThoiGianCongTac ?? 0;
                modelNV.phuCapThamNien = (decimal?)listChoViec.PhuCapThamNien ?? 0;
                modelNV.phuCapChucVu = (decimal?)listChoViec.PhuCapChucVu ?? 0;
                modelNV.phuCapCongTrinh = (decimal?)listChoViec.PhuCapCongTrinh ?? 0;
                modelNV.phuCapThuHut = (decimal?)listChoViec.PhuCapThuHut ?? 0;
                modelNV.phuCapPhatSinh = (decimal?)listChoViec.PhuCapPhatSinh ?? 0;
                modelNV.phuCapDacBiet = (decimal?)listChoViec.PhuCapDacBiet ?? 0;
                modelNV.truyLanh = (decimal?)listChoViec.TruyLanh ?? 0;
                modelNV.truyThu = (decimal?)listChoViec.TruyThu ?? 0;
                modelNV.truyLanh = (decimal?)listChoViec.TruyLanh ?? 0;
                modelNV.Net = listChoViec.Net ?? false;
                modelNV.khongChamCong = listChoViec.khongChamCong ?? false;
                modelNV.maChiNhanhNganHang = listChoViec.nganHang;
                modelNV.soTaiKhoan = listChoViec.soTaiKhoan;
                var chiNhanhNganHangs = context.tbl_DM_ChiNhanhNganHangs.ToList();
                chiNhanhNganHangs.Insert(0, new BatDongSan.Models.NhanSu.tbl_DM_ChiNhanhNganHang { maChiNhanh = "", tenChiNhanh = "" });
                ViewBag.ChiNhanhNHs = new SelectList(chiNhanhNganHangs, "maChiNhanh", "tenChiNhanh", listChoViec.nganHang);

                return View(modelNV);
            }
            catch
            {
                return View("Error");
            }
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
            try
            {
                
                var listChoViec = linqVIP.tbl_NhanVienVIPs.Where(d => d.MaNhanVien == id).FirstOrDefault();
                modelNV = new NhanVienVIPModel();
                modelNV.maNhanVien = listChoViec.MaNhanVien;
                modelNV.tenNhanVien = listChoViec.TenNhanVien;
                modelNV.email = listChoViec.Email;
                modelNV.ngayVaoLam = listChoViec.NgayVaoLam;
                modelNV.ngayLap = listChoViec.ngayLap;
                modelNV.ghiChu = listChoViec.GhiChu;
                modelNV.luongCoBan = (decimal?)listChoViec.LuongCoBan ?? 0;
                modelNV.tongLuong = (decimal?)listChoViec.TongLuong ?? 0;
                modelNV.thoiGianCongTac = (float?)listChoViec.ThoiGianCongTac ?? 0;
                modelNV.phuCapThamNien = (decimal?)listChoViec.PhuCapThamNien ?? 0;
                modelNV.phuCapChucVu = (decimal?)listChoViec.PhuCapChucVu ?? 0;
                modelNV.phuCapCongTrinh = (decimal?)listChoViec.PhuCapCongTrinh ?? 0;
                modelNV.phuCapThuHut = (decimal?)listChoViec.PhuCapThuHut ?? 0;
                modelNV.phuCapPhatSinh = (decimal?)listChoViec.PhuCapPhatSinh ?? 0;
                modelNV.phuCapDacBiet = (decimal?)listChoViec.PhuCapDacBiet ?? 0;
                modelNV.truyLanh = (decimal?)listChoViec.TruyLanh ?? 0;
                modelNV.truyThu = (decimal?)listChoViec.TruyThu ?? 0;
                modelNV.truyLanh = (decimal?)listChoViec.TruyLanh ?? 0;
                modelNV.Net = listChoViec.Net ?? false;
                modelNV.khongChamCong = listChoViec.khongChamCong ?? false;
                modelNV.maChiNhanhNganHang = listChoViec.nganHang;
                modelNV.soTaiKhoan = listChoViec.soTaiKhoan;

                var chiNhanhNganHangs = context.tbl_DM_ChiNhanhNganHangs.ToList();
                chiNhanhNganHangs.Insert(0, new BatDongSan.Models.NhanSu.tbl_DM_ChiNhanhNganHang { maChiNhanh = "", tenChiNhanh = "" });
                ViewBag.ChiNhanhNHs = new SelectList(chiNhanhNganHangs, "maChiNhanh", "tenChiNhanh", listChoViec.nganHang);
                return View(modelNV);
            }
            catch
            {
                return View("Error");
            }
        }
        // POST: /HopDongLaoDong/Create

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
                choViec = new BatDongSan.Models.VIP.tbl_NhanVienVIP();
                GetDataFromView(collection, true);
                linqVIP.tbl_NhanVienVIPs.InsertOnSubmit(choViec);
                linqVIP.SubmitChanges();
                
                // Cap nhat quyen mac dinh VIP
                var userId = linqDM.GetTable<Sys_User>().Where(d => d.manv == choViec.MaNhanVien).Select(d => d.userId).FirstOrDefault();
                LinqHeThongDataContext linqHT = new LinqHeThongDataContext();
                Sys_UserThuocNhom userThuocNhom = new Sys_UserThuocNhom();
                userThuocNhom.userId = userId;
                userThuocNhom.maNhomUser = "XemVIP";
                linqHT.Sys_UserThuocNhoms.InsertOnSubmit(userThuocNhom);
                linqHT.SubmitChanges();
                SaveActiveHistory("Cập nhật nhóm user từ create nhân viên VIP, userID: " + userId + " nhóm: " + userThuocNhom.maNhomUser);
                // End cap nhat
                return RedirectToAction("Edit", new { id = choViec.MaNhanVien });
            }
            catch
            {
                return View("Error");
            }
        }

        [HttpPost]
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
                
                tbl_BangLuongVIP nhanVienVipEdit = new tbl_BangLuongVIP();
                var getIdHienTai = linqVIP.tbl_NhanVienVIPs.Where(d => d.MaNhanVien == collection["maNhanVienOld"]).ToList();
                linqVIP.tbl_NhanVienVIPs.DeleteAllOnSubmit(getIdHienTai);
                linqVIP.SubmitChanges();
                //Bind data
                choViec = new BatDongSan.Models.VIP.tbl_NhanVienVIP();
                GetDataFromView(collection, true);
                linqVIP.tbl_NhanVienVIPs.InsertOnSubmit(choViec);
                linqVIP.SubmitChanges();
                //End
                return RedirectToAction("Edit", new { id = choViec.MaNhanVien });
            }
            catch
            {
                return View("Error");
            }
        }

        public void GetDataFromView(FormCollection collection, bool isCreate)
        {
        
                choViec.MaNhanVien = collection["maNhanVien"];
           
           
            
            choViec.TenNhanVien = collection["tenNhanVien"];
            choViec.ngayLap = DateTime.Now;
            choViec.nguoiLap = GetUser().manv;
            choViec.LuongCoBan = string.IsNullOrEmpty(collection["luongCoBan"]) ? 0 : Convert.ToDecimal(collection["luongCoBan"]);
            choViec.NgayVaoLam = String.IsNullOrEmpty(collection["ngayVaoLam"]) ? (DateTime?)null : DateTime.ParseExact(collection["ngayVaoLam"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
            choViec.ThoiGianCongTac = Convert.ToDouble(collection["thoiGianCongTac"]);
            choViec.Email = collection["email"];
            //Phu cap
            choViec.PhuCapThamNien = string.IsNullOrEmpty(collection["phuCapThamNien"]) ? 0 : Convert.ToDecimal(collection["phuCapThamNien"]);
            choViec.PhuCapChucVu= string.IsNullOrEmpty(collection["phuCapChucVu"]) ? 0 : Convert.ToDecimal(collection["phuCapChucVu"]);
            choViec.PhuCapCongTrinh = string.IsNullOrEmpty(collection["phuCapCongTrinh"]) ? 0 : Convert.ToDecimal(collection["phuCapCongTrinh"]);
            choViec.PhuCapThuHut = string.IsNullOrEmpty(collection["phuCapThuHut"]) ? 0 : Convert.ToDecimal(collection["phuCapThuHut"]);
            choViec.PhuCapDacBiet = string.IsNullOrEmpty(collection["phuCapDacBiet"]) ? 0 : Convert.ToDecimal(collection["phuCapDacBiet"]);
            choViec.PhuCapPhatSinh = string.IsNullOrEmpty(collection["phuCapPhatSinh"]) ? 0 : Convert.ToDecimal(collection["phuCapPhatSinh"]);
            choViec.TruyLanh = string.IsNullOrEmpty(collection["truyLanh"]) ? 0 : Convert.ToDecimal(collection["truyLanh"]);
            choViec.TruyThu = string.IsNullOrEmpty(collection["truyThu"]) ? 0 : Convert.ToDecimal(collection["truyThu"]);
            choViec.TongLuong = string.IsNullOrEmpty(collection["tongLuong"]) ? 0 : Convert.ToDecimal(collection["tongLuong"]);
            choViec.Net = collection["Net"].Contains("true") ? true : false;
            choViec.khongChamCong = collection["khongChamCong"].Contains("true") ? true : false;
            choViec.GhiChu = collection.Get("ghiChu");
            choViec.soTaiKhoan = collection["soTaiKhoan"];
            choViec.nganHang = collection["maChiNhanhNganHang"];
        }
        public void SetDataModelNhanVienVip(BatDongSan.Models.VIP.tbl_NhanVienVIP nhanVienVip)
        {
            modelNV.maNhanVien = nhanVienVip.MaNhanVien;
            modelNV.ngayLap = nhanVienVip.ngayLap;
            modelNV.ghiChu = nhanVienVip.GhiChu;
            modelNV.luongCoBan = (decimal?)nhanVienVip.LuongCoBan ?? 0;
            modelNV.tongLuong = (decimal?)nhanVienVip.TongLuong ?? 0;
            modelNV.thoiGianCongTac = (float?)nhanVienVip.ThoiGianCongTac ?? 0;
            modelNV.phuCapThamNien = (decimal?)nhanVienVip.PhuCapThamNien ?? 0;
            modelNV.phuCapChucVu = (decimal?)nhanVienVip.PhuCapChucVu ?? 0;
            modelNV.email = nhanVienVip.Email;
        }


        //
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
                choViec = linqVIP.tbl_NhanVienVIPs.Where(s => s.MaNhanVien == id).FirstOrDefault();
                linqVIP.tbl_NhanVienVIPs.DeleteOnSubmit(choViec);
                linqVIP.SubmitChanges();
                // Check xóa trong user vip
                LinqHeThongDataContext linqHT = new LinqHeThongDataContext();
                var userID = linqHT.Sys_Users.Where(d => d.manv == id).Select(d => d.userId).FirstOrDefault();
                if (userID != null)
                {
                    var userThuocNhom = linqHT.Sys_UserThuocNhoms.Where(d => d.maNhomUser == "XemVIP" && d.userId == userID).ToList();
                    if (userThuocNhom != null)
                    {
                        linqHT.Sys_UserThuocNhoms.DeleteAllOnSubmit(userThuocNhom);
                        linqHT.SubmitChanges();
                    }
                }
                return RedirectToAction("Index");
            }
            catch
            {
                return View("Error");
            }
        }
        public int CheckMaNhanVien(string maNhanVien)
        {
            var checkList = linqVIP.tbl_NhanVienVIPs.Where(d => d.MaNhanVien == maNhanVien).FirstOrDefault();
            if (checkList != null)
            {
                return 1;
            }
            return 0;
        }
        private void thang(int value)
        {
            Dictionary<int, string> dics = new Dictionary<int, string>();
            for (int i = 0; i < 13; i++)
            {
                dics[i] = i.ToString();
            }
            ViewData["thang"] = new SelectList(dics, "Key", "Value", value);
            ViewData["thangtc"] = new SelectList(dics, "Key", "Value", value);
        }
        private void nam(int value)
        {
            Dictionary<int, string> dics = new Dictionary<int, string>();
            for (int i = 2015; i < 2031; i++)
            {
                dics[i] = i.ToString();
            }
            ViewData["nam"] = new SelectList(dics, "Key", "Value", value);
            ViewData["namtc"] = new SelectList(dics, "Key", "Value", value);
        }

        //public void SetDataModelChoViec(tbl_NS_NhanVienVip choViec)
        //{
        //    model.maNhanVienVip = choViec.maNhanVienVip;
        //    model.ngayLap = choViec.ngayLap;
        //    model.ghiChu = choViec.ghiChu;
        //    model.tyLeNhanVienVip = (double)choViec.tyLe;
        //    model.congChoViec = (double)choViec.congChoViec;
        //    model.luongChoViec = (decimal?)choViec.luongChoViec ?? 0;
        //    model.tenHopDong = context.tbl_NS_HopDongLaoDongs.Where(d => d.soHopDong == choViec.soHopDong).Select(d => d.tenHopDong).FirstOrDefault() ?? string.Empty;
        //    model.soPhuLuc = choViec.soPhuLuc;
        //    model.NguoiLap = new NhanVienModel(choViec.nguoiLap, HoVaTen(choViec.nguoiLap));
        //    model.NhanVien = new NhanVienModel(choViec.maNhanVien, HoVaTen(choViec.maNhanVien));
        //    model.luongThoaThuan = choViec.luong;
        //}

        #endregion

        public string IdGenerator()
        {
            return GenerateUtil.CheckLetter("LCV", GetMax());
        }

        public string GetMax()
        {
            return linqVIP.tbl_NhanVienVIPs.OrderByDescending(d=>d.ngayLap).Select(d => d.MaNhanVien).FirstOrDefault();
        }

        #region  Load danh sách nhân viên

        public ActionResult NhanVienDuyet()
        {
            try
            {
                buildTree = new StringBuilder();
                phongBans = linqDanhMuc.tbl_DM_PhongBans.ToList();
                buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
                ViewBag.NVPB = buildTree.ToString();
                return PartialView("_NhanVienPhongBan");
            }
            catch
            {
                return View("Error");
            }
        }

        public ActionResult LoadNhanVienDuyet(int? page, string searchString, string maPhongBan)
        {
            try
            {
                IList<sp_PB_DanhSachNhanVienResult> phongBan1s;
                phongBan1s = linqDanhMuc.sp_PB_DanhSachNhanVien(searchString, maPhongBan).ToList();
                ViewBag.isGet = "True";
                int currentPageIndex = page.HasValue ? page.Value : 1;
                ViewBag.Count = phongBan1s.Count();
                ViewBag.Search = searchString;
                ViewBag.MaPhongBan = maPhongBan;
                return PartialView("LoadNhanVienDuyet", phongBan1s.ToPagedList(currentPageIndex, 10));
            }
            catch
            {
                return View("Error");
            }
        }
        #endregion

        #region Thông tin  thỏa thuận mức lương khi ký hợp đồng của nhân viên đó


        public JsonResult ThongTinNhanVien(string maNhanVien)
        {
            try
            {
                model = new BatDongSan.Models.NhanSu.HopDongLaoDongModel();
                string maHopDong = context.tbl_NS_HopDongLaoDongs.Where(d => d.maNhanVien == maNhanVien).Select(d => d.soHopDong).FirstOrDefault() ?? string.Empty;
                model = SetModelData(maHopDong);
                return Json(model);
            }
            catch
            {
                return Json("False");
            }
        }

        public BatDongSan.Models.NhanSu.HopDongLaoDongModel SetModelData(string maHopDong)
        {
            //Kiểm tra hợp đồng đó có phụ lục hay chưa nếu có thì lấy phụ lục mới nhất và phụ lục đó đã duyệt
            DMNguoiDuyetController nd = new DMNguoiDuyetController();
            string soPhuLuc = string.Empty;
            var dsPhuLuc = context.tbl_NS_PhuLucHopDongs.Where(d => d.soHopDong == maHopDong).Select(d => new BatDongSan.Models.NhanSu.PhuLucHopDongModel
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
                model = (from hd in context.tbl_NS_PhuLucHopDongs
                         where hd.soPhuLuc == soPhuLuc
                         select new BatDongSan.Models.NhanSu.HopDongLaoDongModel
                         {
                             soHopDong = hd.soHopDong,
                             tenHopDong = context.tbl_NS_HopDongLaoDongs.Where(d => d.soHopDong == maHopDong).Select(d => d.tenHopDong).FirstOrDefault() ?? string.Empty,
                             soPhuLuc = soPhuLuc,
                             luongDongBH = hd.luongDongBaoHiem ?? 0,
                             khoanBoSungLuong = hd.khoanBoSungLuong ?? 0,
                             phuCapLuong = (double?)hd.phuCapLuong ?? 0,
                             tongLuong = hd.tongLuong ?? 0,
                             luongHopDong = hd.luong,
                             
                         }).FirstOrDefault();
            }
            else
            {
                model = (from hd in context.tbl_NS_HopDongLaoDongs
                         where hd.soHopDong == maHopDong
                         select new BatDongSan.Models.NhanSu.HopDongLaoDongModel
                         {
                             soHopDong = hd.soHopDong,
                             tenHopDong = hd.tenHopDong,
                             soPhuLuc = string.Empty,
                             luongDongBH = hd.luongDongBaoHiem ?? 0,
                             khoanBoSungLuong = hd.khoanBoSungLuong ?? 0,
                             phuCapLuong = (double?)hd.phuCapLuong ?? 0,
                             tongLuong = hd.tongLuong ?? 0,
                             luongHopDong = hd.luongThoaThuan,
                         }).FirstOrDefault();
            }
            return model;
        }
        #endregion
    }
}
