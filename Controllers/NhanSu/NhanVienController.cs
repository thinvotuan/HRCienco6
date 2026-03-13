using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper;
using BatDongSan.Helper.Common;
using BatDongSan.Helper.Utils;
using BatDongSan.Models.ChamCong;
using BatDongSan.Models.DanhMuc;
using BatDongSan.Models.NhanSu;
using BatDongSan.Utils;
using BatDongSan.Utils.Paging;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.UserModel.Contrib;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using Worldsoft.Mvc.Web.Util;

namespace BatDongSan.Controllers.NhanSu
{
    public class NhanVienController : ApplicationController
    {
        private LinqNhanSuDataContext context;
        LinqDanhMucDataContext linqDM = new LinqDanhMucDataContext();
        private IList<sp_NS_NhanVien_IndexResult> nhanViens;
        private IList<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan> phongBans;
        private IList<tbl_NS_NhanVienChungChiNgoaiNgu> ngoaiNgus;
        private tbl_NS_NhanVien nhanVien;
        private tbl_NS_NhanVienChiNhanh nhanVienChiNhanh;
        private tbl_NS_NhanVienChucDanh nhanVienChucDanh;
        private tbl_NS_NhanVienPhongBan nhanVienPhongBan;
        //private tbl_NS_BoPhanTinhLuong nhanVienTinhLuong;
        private tbl_NS_PhanCaNhanVien nhanVienPhanCa;
        private NhanVienModel model;
        private StringBuilder buildTree;
        private readonly string MCV = "NhanVien";
        private bool? permission;

        public ActionResult Index(int? page, int? pageSize, string searchString, string maPhongBan, int? trangThai)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            int currentPageIndex = page.HasValue ? page.Value : 1;
            pageSize = pageSize ?? 30;
            int? tongSoDong = 0;
            using (context = new LinqNhanSuDataContext())
            {
                buildTree = new StringBuilder();
                phongBans = context.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan>().ToList();
                buildTree = TreePhongBanStyle.BuildTreeDepartment(phongBans);
                var nhanVienList = context.sp_NS_NhanVien_Index2(searchString, null, maPhongBan, trangThai, currentPageIndex, pageSize).ToList();

                try
                {
                    ViewBag.Count = nhanVienList[0].tongSoDong;
                    tongSoDong = nhanVienList[0].tongSoDong;
                }
                catch
                {
                    ViewBag.Count = 0;
                }
                ViewBag.PhongBans = buildTree.ToString();
                TempData["Params"] = searchString + "," + maPhongBan + "," + trangThai;

                if (Request.IsAjaxRequest())
                {
                    return PartialView("PartialIndex", nhanVienList.ToPagedList(currentPageIndex, 30, true, tongSoDong));
                }
                return View(nhanVienList.ToPagedList(currentPageIndex, 30, true, tongSoDong));
            }
        }

        //
        // GET: /NhanVien/Details/5

        public ActionResult Details(string id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            context = new LinqNhanSuDataContext();
            nhanVien = (from nv in context.tbl_NS_NhanViens
                        where nv.maNhanVien == id
                        select nv).FirstOrDefault();
            SetModelData();

            var nhanVienDetails = context.sp_NS_NhanVien_Index(null, id, null, null, null).FirstOrDefault();
            if (nhanVienDetails == null) nhanVienDetails = new sp_NS_NhanVien_IndexResult();
            model.maPhongBan = nhanVienDetails.maPhongBan;
            model.maChucDanh = nhanVienDetails.maChucDanh;
            model.maChiNhanhVanPhong = nhanVienDetails.maChiNhanh;
            model.tenPhongBan = nhanVienDetails.tenPhongBan;
            model.tenCapBacQL = nhanVienDetails.tenCapBac;
            model.tenChiNhanhNganHang = model.tenChiNhanhNganHang;
            model.tenChiNhanhVanPhong = model.tenChiNhanhVanPhong;
            model.tenKhoiTinhLuong = nhanVienDetails.tenKhoiTinhLuong;
            model.maChamCong = nhanVienDetails.maChamCong;
            model.tenChucDanh = nhanVienDetails.TenChucDanh;
            model.fileDinhKems = context.Sys_FileDinhKems.Where(s => s.identification == model.maNhanVien).ToList();
            GetThongTinQuanHeGD(model.maNhanVien);
            //model.hinhAnh = ImageHelper.byteArrayToImage(model.anhDaiDienBinary.ToArray());
            GetAllDropdownlist();
            // Check Phan ca: co 2 dong khong cho sua.
            int resultPhanCa = 1;
            var listPC = linqDM.tbl_NS_PhanCaNhanViens.Where(d => d.maNhanVien == nhanVienDetails.maNhanVien).ToList();
            if (listPC != null && listPC.Count >= 2)
            {
                resultPhanCa = 0;
            }
            ViewBag.resultPC = resultPhanCa;
            if (resultPhanCa == 1)
            {
                var listPCFirst = linqDM.tbl_NS_PhanCaNhanViens.Where(d => d.maNhanVien == nhanVienDetails.maNhanVien).FirstOrDefault();
                if (listPCFirst != null)
                {
                    model.maPhanCa = listPCFirst.maPhanCa;
                }
            }
            // Check Phong ban: co 2 dong khong cho sua
            int resultPhongBan = 1;
            var listPB = context.tbl_NS_NhanVienPhongBans.Where(d => d.maNhanVien == nhanVienDetails.maNhanVien).ToList();
            if (listPB != null && listPB.Count >= 2)
            {
                resultPhongBan = 0;
            }
            ViewBag.resultPB = resultPhongBan;
            //End check phong ban
            // Check chuc danh: co 2 dong khong cho sua
            int resultChucDanh = 1;
            var listCD = context.tbl_NS_NhanVienChucDanhs.Where(d => d.maNhanVien == nhanVienDetails.maNhanVien).ToList();
            if (listCD != null && listCD.Count >= 2)
            {
                resultChucDanh = 0;
            }
            ViewBag.resultCD = resultChucDanh;
            //End check phong ban

            return View(model);
        }

        //
        public ActionResult CapNhatLuyKeNB(string maNhanVien, float luyKeNghiBu)
        {
            try
            {
                #region Role user
                permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
                if (!permission.HasValue)
                    return Json("LogIn");
                if (!permission.Value)
                    return Json("AccessDenied");
                #endregion
                context = new LinqNhanSuDataContext();
                var tblNhanVienLK = context.tbl_NS_NhanViens.Where(d => d.maNhanVien == maNhanVien).FirstOrDefault();
                tblNhanVienLK.luyKeNghiBu = luyKeNghiBu;
                context.SubmitChanges();
                SaveActiveHistory("Cập nhật số nghỉ bù cho nhân viên: " + HoVaTen(maNhanVien) + " số ngày: " + luyKeNghiBu);
                return Json(new { messageError = string.Empty });
            }
            catch (Exception ex)
            {
                //write log
                Log4Net.WriteLog(log4net.Core.Level.Error, ex.Message);

                return Json(new { messageError = "Có lỗi xảy ra, vui lòng liên hệ bộ phận IT." });
            }
        }
        // GET: /NhanVien/Create

        public ActionResult Create()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenThem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            context = new LinqNhanSuDataContext();
            model = new NhanVienModel();
            model.maChucDanh = "NV";
            model.maQuocTich = "VN";
            model.trangThai = 1;
            GetAllDropdownlist();
            model.dangBaoHiemNoiKhac = false;
            model.maNhanVien = IdGeneratorMaNV();

            return View(model);
        }

        //
        // POST: /NhanVien/Create

        [HttpPost]
        public ActionResult Create(FormCollection collection, IEnumerable<HttpPostedFileBase> fileDinhKems, HttpPostedFileBase hinhAnh)
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
                //BatDongSan.Models.ChamCong.LinqChamCongServerDataContext contextCC = new BatDongSan.Models.ChamCong.LinqChamCongServerDataContext();
                //Check Ton tai
                // check ton tai cong nhan
                if (collection["maChamCong"] != null && collection["maChamCong"] != "")
                {
                    var checkListNV = linqDM.tbl_NS_DanhSachCongNhans.Where(d => d.maChamCong == collection["maChamCong"]).FirstOrDefault();
                    if (checkListNV != null)
                    {

                        return View("error");
                    }
                    var checkListNS = context.tbl_NS_NhanViens.Where(d => d.maChamCong == collection["maChamCong"]).FirstOrDefault();
                    if (checkListNS != null)
                    {

                        return View("error");
                    }


                    //var checkListNVCC = contextCC.WS_tblUserinfos.Where(d => d.badgenumber == collection["maChamCong"]).FirstOrDefault();
                    //if (checkListNVCC != null)
                    //{
                    //    return View("error");
                    //}
                }
                // End check ton tai

                nhanVien = new tbl_NS_NhanVien();
                nhanVien.ngayLap = DateTime.Now;
                nhanVien.nguoiLap = GetUser().manv;
                nhanVien.maNhanVien = collection["maNhanVien"];
                //Lưu dữ liệu nhân viên
                GetDataFromView(collection);

                //Lưu ảnh đại diện
                UploadAvartar(hinhAnh);

                if (String.IsNullOrEmpty(nhanVien.anhDaiDienURL)) nhanVien.anhDaiDienURL = "/Images/card.gif";

                //Lưu tập tin đính kèm
                UploadFileDinhKem(collection, fileDinhKems);

                //Lưu các chứng chỉ ngoại ngữ
                //InsertChungChiNgoaiNgu(collection);

                var checkList = context.tbl_NS_NhanViens.Where(s => s.CMNDSo == nhanVien.CMNDSo || s.maNhanVien == nhanVien.maNhanVien).ToList();

                if (checkList.Count() > 0)
                {
                    if (checkList.Where(s => s.maNhanVien == nhanVien.maNhanVien).Count() > 0)
                    {
                        TempData["MaNhanVien"] = "Mã nhân viên đã tồn tại";
                    }
                    if (checkList.Where(s => s.CMNDSo == nhanVien.CMNDSo).Count() > 0)
                    {
                        TempData["CMND"] = "Số CMND đã tồn tại";
                    }
                    SetModelData();
                    GetAllDropdownlist();
                    return View(model);
                }


                context.tbl_NS_NhanViens.InsertOnSubmit(nhanVien);
                context.tbl_NS_NhanVienPhongBans.InsertOnSubmit(nhanVienPhongBan);

                //context.tbl_NS_NhanVienChiNhanhs.InsertOnSubmit(nhanVienChiNhanh);
                context.tbl_NS_NhanVienChucDanhs.InsertOnSubmit(nhanVienChucDanh);

                context.SubmitChanges();
                // Bo phan tinh luong
                //var lstBoPhanTL = linqDM.tbl_NS_KhoiTinhLuongs.Where(d => d.maKhoiTinhLuong == Convert.ToInt32(collection["idKhoiTinhLuong"])).FirstOrDefault();
                //if (lstBoPhanTL != null)
                //{
                //    nhanVienTinhLuong = new tbl_NS_BoPhanTinhLuong();
                //    nhanVienTinhLuong.maNhanVien = nhanVien.maNhanVien;
                //    nhanVienTinhLuong.maBoPhan = collection["idKhoiTinhLuong"];
                //    nhanVienTinhLuong.tenBoPhan = lstBoPhanTL.tenKhoiTinhLuong;
                //    context.tbl_NS_BoPhanTinhLuongs.InsertOnSubmit(nhanVienTinhLuong);


                //    context.SubmitChanges();
                //}
                // End bo phan tinh luong
                //Inset Phan Ca
                nhanVienPhanCa = new tbl_NS_PhanCaNhanVien();
                nhanVienPhanCa.maNhanVien = nhanVien.maNhanVien;

                nhanVienPhanCa.maPhanCa = short.Parse(!string.IsNullOrEmpty(collection["maPhanCa"]) ? collection["maPhanCa"] : "0");
                nhanVienPhanCa.ngayLap = DateTime.Now;
                nhanVienPhanCa.nguoiLap = GetUser().manv;
                nhanVienPhanCa.ngayApDung = DateTime.ParseExact(collection["ngayVaoLam"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                linqDM.tbl_NS_PhanCaNhanViens.InsertOnSubmit(nhanVienPhanCa);
                linqDM.SubmitChanges();
                // End Phan ca

                //// Insert vao tbl nhan vien
                //if (collection["maChamCong"] != null && collection["maChamCong"] != "")
                //{
                //    // Add Moi tblNhanvien cham cong
                //    BatDongSan.Models.ChamCong.WS_tblUserinfo tblNhanVienMoi = new BatDongSan.Models.ChamCong.WS_tblUserinfo();
                //    tblNhanVienMoi.badgenumber = nhanVien.maChamCong;
                //    tblNhanVienMoi.fullName = nhanVien.ho + ' ' + nhanVien.ten;
                //    tblNhanVienMoi.employeeId = nhanVien.maNhanVien;
                //    contextCC.WS_tblUserinfos.InsertOnSubmit(tblNhanVienMoi);
                //    contextCC.SubmitChanges();
                //}

                SaveActiveHistory("Thêm mới nhân viên: " + nhanVien.maNhanVien);
                //linqChamCongServer.SubmitChanges();
                return RedirectToAction("Edit", new { id = nhanVien.maNhanVien });
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /NhanVien/Edit/5

        public ActionResult Edit(string id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            context = new LinqNhanSuDataContext();
            nhanVien = (from nv in context.tbl_NS_NhanViens
                        where nv.maNhanVien == id
                        select nv).FirstOrDefault();
            SetModelData();

            var nhanVienDetails = context.sp_NS_NhanVien_Index(null, id, null, null, null).FirstOrDefault();
            if (nhanVienDetails == null) nhanVienDetails = new sp_NS_NhanVien_IndexResult();
            model.maPhongBan = nhanVienDetails.maPhongBan;
            model.maChucDanh = nhanVienDetails.maChucDanh;
            model.tenChucDanh = nhanVienDetails.TenChucDanh;
            model.maChiNhanhVanPhong = nhanVienDetails.maChiNhanh;
            model.tenPhongBan = nhanVienDetails.tenPhongBan;
            model.tenCapBacQL = nhanVienDetails.tenCapBac;
            model.tenChiNhanhNganHang = model.tenChiNhanhNganHang;
            model.tenChiNhanhVanPhong = model.tenChiNhanhVanPhong;
            model.tenKhoiTinhLuong = nhanVienDetails.tenKhoiTinhLuong;
            model.maChamCong = nhanVienDetails.maChamCong;
            model.fileDinhKems = context.Sys_FileDinhKems.Where(s => s.identification == model.maNhanVien).ToList();
            GetThongTinQuanHeGD(model.maNhanVien);
            //model.hinhAnh = ImageHelper.byteArrayToImage(model.anhDaiDienBinary.ToArray());
            GetAllDropdownlist();
            // Check Phan ca: co 2 dong khong cho sua.
            int resultPhanCa = 1;
            var listPC = linqDM.tbl_NS_PhanCaNhanViens.Where(d => d.maNhanVien == nhanVienDetails.maNhanVien).ToList();
            if (listPC != null && listPC.Count >= 2)
            {
                resultPhanCa = 0;
            }
            ViewBag.resultPC = resultPhanCa;
            if (resultPhanCa == 1)
            {
                var listPCFirst = linqDM.tbl_NS_PhanCaNhanViens.Where(d => d.maNhanVien == nhanVienDetails.maNhanVien).FirstOrDefault();
                if (listPCFirst != null)
                {
                    model.maPhanCa = listPCFirst.maPhanCa;
                }
            }
            // Check Phong ban: co 2 dong khong cho sua
            int resultPhongBan = 1;
            var listPB = context.tbl_NS_NhanVienPhongBans.Where(d => d.maNhanVien == nhanVienDetails.maNhanVien).ToList();
            if (listPB != null && listPB.Count >= 2)
            {
                resultPhongBan = 0;
            }
            ViewBag.resultPB = resultPhongBan;
            //End check phong ban
            // Check chuc danh: co 2 dong khong cho sua
            int resultChucDanh = 1;
            var listCD = context.tbl_NS_NhanVienChucDanhs.Where(d => d.maNhanVien == nhanVienDetails.maNhanVien).ToList();
            if (listCD != null && listCD.Count >= 2)
            {
                resultChucDanh = 0;
            }
            ViewBag.resultCD = resultChucDanh;
            //End check phong ban

            return View(model);
        }

        //
        // POST: /NhanVien/Edit/5

        [HttpPost]
        public ActionResult Edit(string id, FormCollection collection, HttpPostedFileBase hinhAnh, IEnumerable<HttpPostedFileBase> fileDinhKems)
        {
            try
            {
                //LinqChamCongServerDataContext linqChamCongServer = new LinqChamCongServerDataContext();//Get dữ liệu từ mấy chấm công
                context = new LinqNhanSuDataContext();
                nhanVien = context.tbl_NS_NhanViens.Where(s => s.maNhanVien == id).FirstOrDefault();
                //Lưu dữ liệu nhân viên
                GetDataFromView(collection);

                //Lưu ảnh đại diện
                UploadAvartar(hinhAnh);

                if (String.IsNullOrEmpty(nhanVien.anhDaiDienURL)) nhanVien.anhDaiDienURL = "/Images/card.gif";

                //Lưu tập tin đính kèm
                UploadFileDinhKem(collection, fileDinhKems);

                //Lưu các chứng chỉ ngoại ngữ
                // InsertChungChiNgoaiNgu(collection);

                //Get dữ liệu phòng ban nhan viên hợp đồng để kiếm tra điều kiện hợp lệ
                var checkList = context.tbl_NS_NhanViens.Where(s => s.maNhanVien != id && (s.CMNDSo == nhanVien.CMNDSo || s.maNhanVien == nhanVien.maNhanVien)).ToList();
                //var phongBan = context.tbl_NS_NhanVienPhongBans.Where(s => s.maNhanVien == nhanVien.maNhanVien).OrderByDescending(o => o.id).FirstOrDefault();
                //var hopDong = context.tbl_NS_HopDongLaoDongs.Where(s => s.maNhanVien == nhanVien.maNhanVien).FirstOrDefault();

                //Get thông tin chi tiết đầy đủ của nhân viên
                var nhanVienDetails = context.sp_NS_NhanVien_Index(null, id, null, null, null).FirstOrDefault();
                if (nhanVienDetails == null) nhanVienDetails = new sp_NS_NhanVien_IndexResult();

                //if (hopDong != null && phongBan.maPhongBan != nhanVienPhongBan.maPhongBan)
                //{
                //    TempData["PhongBan"] = "Không thể thay đổi phòng ban khi nhân viên đã ký hợp đồng <br /> Thay đổi thành: " + nhanVienDetails.tenPhongBan;
                //}

                if (checkList.Count() > 0)
                {
                    if (checkList.Where(s => s.maNhanVien == nhanVien.maNhanVien).Count() > 0)
                    {
                        TempData["MaNhanVien"] = "Mã nhân viên đã tồn tại";
                    }
                    if (checkList.Where(s => s.CMNDSo == nhanVien.CMNDSo).Count() > 0)
                    {
                        TempData["CMND"] = "Số CMND đã tồn tại";
                    }
                    SetModelData();
                    model.maPhongBan = nhanVienDetails.maPhongBan;
                    model.maChucDanh = nhanVienDetails.maChucDanh;
                    model.maChiNhanhVanPhong = nhanVienDetails.maChiNhanh;
                    model.tenPhongBan = collection["tenPhongBan"];

                    model.fileDinhKems = context.Sys_FileDinhKems.Where(s => s.identification == model.maNhanVien).ToList();
                    GetThongTinQuanHeGD(model.maNhanVien);
                    GetAllDropdownlist();
                    return View(model);
                }


                context.SubmitChanges();
                //linqChamCongServer.SubmitChanges();
                var listPC = linqDM.tbl_NS_PhanCaNhanViens.Where(d => d.maNhanVien == id).ToList();
                if ((listPC != null && listPC.Count == 1) || listPC.Count == 0)
                {
                    var dataPC = linqDM.tbl_NS_PhanCaNhanViens.Where(d => d.maNhanVien == id).FirstOrDefault();
                    if (dataPC != null)
                    {
                        linqDM.tbl_NS_PhanCaNhanViens.DeleteOnSubmit(dataPC);
                        linqDM.SubmitChanges();
                    }
                    //Inset Phan Ca
                    nhanVienPhanCa = new tbl_NS_PhanCaNhanVien();
                    short maPhanCaShort;
                    maPhanCaShort = short.Parse(collection["maPhanCa"]);
                    nhanVienPhanCa.maNhanVien = nhanVien.maNhanVien;
                    nhanVienPhanCa.maPhanCa = maPhanCaShort;
                    nhanVienPhanCa.ngayLap = DateTime.Now;
                    nhanVienPhanCa.nguoiLap = GetUser().manv;
                    nhanVienPhanCa.ngayApDung = DateTime.ParseExact(collection["ngayVaoLam"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    linqDM.tbl_NS_PhanCaNhanViens.InsertOnSubmit(nhanVienPhanCa);
                    linqDM.SubmitChanges();
                }
                // End Phan ca
                var listPB = context.tbl_NS_NhanVienPhongBans.Where(d => d.maNhanVien == id).ToList();
                if ((listPB != null && listPB.Count == 1) || listPB.Count == 0)
                {
                    var dataPB = context.tbl_NS_NhanVienPhongBans.Where(d => d.maNhanVien == id).FirstOrDefault();
                    if (dataPB != null)
                    {
                        context.tbl_NS_NhanVienPhongBans.DeleteOnSubmit(dataPB);
                        context.SubmitChanges();
                    }
                    //Inset Phong Ban
                    nhanVienPhongBan = new tbl_NS_NhanVienPhongBan();

                    nhanVienPhongBan.maNhanVien = nhanVien.maNhanVien;
                    nhanVienPhongBan.maPhongBan = collection["maPhongBan"];
                    nhanVienPhongBan.ngayLap = DateTime.Now;
                    nhanVienPhongBan.nguoiLap = GetUser().manv;
                    context.tbl_NS_NhanVienPhongBans.InsertOnSubmit(nhanVienPhongBan);

                }
                // End Phong Ban
                var listCD = context.tbl_NS_NhanVienChucDanhs.Where(d => d.maNhanVien == id).ToList();
                if ((listCD != null && listCD.Count == 1) || listCD.Count == 0)
                {
                    var dataCD = context.tbl_NS_NhanVienChucDanhs.Where(d => d.maNhanVien == id).FirstOrDefault();
                    if (dataCD != null)
                    {
                        context.tbl_NS_NhanVienChucDanhs.DeleteOnSubmit(dataCD);
                        context.SubmitChanges();
                    }
                    //Inset Chuc Danh
                    nhanVienChucDanh = new tbl_NS_NhanVienChucDanh();

                    nhanVienChucDanh.maNhanVien = nhanVien.maNhanVien;
                    nhanVienChucDanh.maChucDanh = collection["maChucDanh"];
                    nhanVienChucDanh.ngayLap = DateTime.Now;
                    nhanVienChucDanh.nguoiLap = GetUser().manv;
                    context.tbl_NS_NhanVienChucDanhs.InsertOnSubmit(nhanVienChucDanh);

                }
                // End Chuc Danh
                //var dataBPTL = context.tbl_NS_BoPhanTinhLuongs.Where(d => d.maNhanVien == id).ToList();
                //if (dataBPTL != null)
                //{
                //    context.tbl_NS_BoPhanTinhLuongs.DeleteAllOnSubmit(dataBPTL);

                //}
                ////Inset tính lương
                //var lstBoPhanTL = linqDM.tbl_NS_KhoiTinhLuongs.Where(d => d.maKhoiTinhLuong == Convert.ToInt32(collection["idKhoiTinhLuong"])).FirstOrDefault();
                //if (lstBoPhanTL != null)
                //{
                //    nhanVienTinhLuong = new tbl_NS_BoPhanTinhLuong();
                //    nhanVienTinhLuong.maNhanVien = nhanVien.maNhanVien;
                //    nhanVienTinhLuong.maBoPhan = collection["idKhoiTinhLuong"];
                //    nhanVienTinhLuong.tenBoPhan = lstBoPhanTL.tenKhoiTinhLuong;
                //    context.tbl_NS_BoPhanTinhLuongs.InsertOnSubmit(nhanVienTinhLuong);



                //}
                // End Tinh Luong
                // Delete trong tblNhanVien DBChamCong

                //    BatDongSan.Models.ChamCong.LinqChamCongServerDataContext contextCC = new BatDongSan.Models.ChamCong.LinqChamCongServerDataContext();
                //    BatDongSan.Models.ChamCong.WS_tblUserinfo tblNhanVienCC = new BatDongSan.Models.ChamCong.WS_tblUserinfo();
                //    if (collection["maCCBanDau"] != null && collection["maCCBanDau"] != "")
                //    {
                //        tblNhanVienCC = contextCC.WS_tblUserinfos.Where(d => d.badgenumber == collection["maCCBanDau"]).FirstOrDefault();
                //        if (tblNhanVienCC != null)
                //        {
                //            contextCC.WS_tblUserinfos.DeleteOnSubmit(tblNhanVienCC);
                //            contextCC.SubmitChanges();
                //        }
                //    }
                //    if (nhanVien.maChamCong != null && nhanVien.maChamCong != "")
                //    {
                //    // Add Moi tblNhanvien cham cong
                //    BatDongSan.Models.ChamCong.WS_tblUserinfo tblNhanVienMoi = new BatDongSan.Models.ChamCong.WS_tblUserinfo();
                //    tblNhanVienMoi.badgenumber = nhanVien.maChamCong;
                //    tblNhanVienMoi.fullName = nhanVien.ho + ' ' + nhanVien.ten;
                //    tblNhanVienMoi.employeeId = nhanVien.maNhanVien;
                //    contextCC.WS_tblUserinfos.InsertOnSubmit(tblNhanVienMoi);
                //    contextCC.SubmitChanges();
                //}
                context.SubmitChanges();
                SaveActiveHistory("Sửa nhân viên: " + nhanVien.maNhanVien);
                return RedirectToAction("Edit", new { id = nhanVien.maNhanVien });
            }
            catch
            {
                return View();
            }
        }

        //
        // POST: /NhanVien/Delete/5

        [HttpPost]
        public ActionResult Delete(string id)
        {
            try
            {
                using (context = new LinqNhanSuDataContext())
                {
                    BatDongSan.Models.HeThong.LinqHeThongDataContext contextHeThong = new BatDongSan.Models.HeThong.LinqHeThongDataContext();
                    // Check Exist In User.
                    var checkUser = contextHeThong.Sys_Users.Where(d => d.manv == id).FirstOrDefault();
                    if (checkUser != null)
                    {

                        return View();
                    }
                    var chiNhanhs = context.tbl_NS_NhanVienChiNhanhs.Where(s => s.maNhanVien == id);
                    context.tbl_NS_NhanVienChiNhanhs.DeleteAllOnSubmit(chiNhanhs);
                    //var boPhanTLs = context.tbl_NS_BoPhanTinhLuongs.Where(s => s.maNhanVien == id);
                    //context.tbl_NS_BoPhanTinhLuongs.DeleteAllOnSubmit(boPhanTLs);
                    var nhanVienPBs = context.tbl_NS_NhanVienPhongBans.Where(s => s.maNhanVien == id);
                    context.tbl_NS_NhanVienPhongBans.DeleteAllOnSubmit(nhanVienPBs);
                    var nhanVienPhanCas = linqDM.tbl_NS_PhanCaNhanViens.Where(s => s.maNhanVien == id);
                    linqDM.tbl_NS_PhanCaNhanViens.DeleteAllOnSubmit(nhanVienPhanCas);
                    var fileDinhKems = context.Sys_FileDinhKems.Where(s => s.identification == id && s.controller == HtmlRequestHelper.GetRouteData().Controller);
                    context.Sys_FileDinhKems.DeleteAllOnSubmit(fileDinhKems);

                    nhanVien = context.tbl_NS_NhanViens.Where(s => s.maNhanVien == id).FirstOrDefault();
                    context.tbl_NS_NhanViens.DeleteOnSubmit(nhanVien);

                    // Delete trong tblNhanVien DBChamCong
                    //if (nhanVien.maChamCong != null && nhanVien.maChamCong != "")
                    //{
                    //    BatDongSan.Models.ChamCong.LinqChamCongServerDataContext contextCC = new BatDongSan.Models.ChamCong.LinqChamCongServerDataContext();
                    //    BatDongSan.Models.ChamCong.WS_tblUserinfo tblNhanVienCC = new BatDongSan.Models.ChamCong.WS_tblUserinfo();
                    //    tblNhanVienCC = contextCC.WS_tblUserinfos.Where(d => d.badgenumber == nhanVien.maChamCong).FirstOrDefault();
                    //    if (tblNhanVienCC != null)
                    //    {
                    //        contextCC.WS_tblUserinfos.DeleteOnSubmit(tblNhanVienCC);
                    //        contextCC.SubmitChanges();
                    //    }
                    //}
                    context.SubmitChanges();
                }
                SaveActiveHistory("Sửa nhân viên: " + id);
                return RedirectToAction("Index");

            }
            catch
            {
                return View();
            }
        }

        public string IdGeneratorMaNV()
        {
            StringBuilder sb = new StringBuilder();
            var date = DateTime.Now;
            string lastID = context.tbl_NS_NhanViens.OrderByDescending(d => d.ngayLap).Select(d => d.maNhanVien).FirstOrDefault();
            //string nam = date.Year.ToString();
            //nam = nam.Remove(0, 2);
            //string thang = string.Empty;
            //if (date.Month < 10)
            //{
            //    thang = "0" + date.Month;
            //}
            //else
            //{
            //    thang = date.Month.ToString();
            //}
            if (String.IsNullOrEmpty(lastID))
            {
                return "NV-" + "001";
            }
            else
            {
                int? idSuffix = Convert.ToInt32(lastID.Substring(lastID.Length - 3)) + 1;
                if (idSuffix <= 0 || idSuffix == null)
                    return "NV-" + "001";
                else
                {
                    sb.Append(idSuffix.ToString());
                    while (sb.Length < 3)
                    {
                        sb.Insert(0, "0");
                    }
                    return "NV-" + sb.ToString();
                }
            }
        }
        public string IdGenerator()
        {

            StringBuilder sb = new StringBuilder();
            var date = DateTime.Now;
            string lastID = context.tbl_NS_NhanViens.OrderByDescending(d => d.maNhanVien).Select(d => d.maNhanVien).FirstOrDefault();
            string nam = date.Year.ToString();
            nam = nam.Remove(0, 2);
            string thang = string.Empty;
            if (date.Month < 10)
            {
                thang = "0" + date.Month;
            }
            else
            {
                thang = date.Month.ToString();
            }
            if (String.IsNullOrEmpty(lastID))
            {
                return "NV-" + nam + thang + "001";
            }
            else
            {
                int? idSuffix = Convert.ToInt32(lastID.Substring(lastID.Length - 3)) + 1;
                if (idSuffix <= 0 || idSuffix == null)
                    return "NV-" + nam + thang + "001";
                else
                {
                    sb.Append(idSuffix.ToString());
                    while (sb.Length < 3)
                    {
                        sb.Insert(0, "0");
                    }
                    return "NV-" + nam + thang + sb.ToString();
                }
            }
        }


        public ActionResult PhongBanListTree()
        {

            using (context = new LinqNhanSuDataContext())
            {
                buildTree = new StringBuilder();
                phongBans = context.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_PhongBan>().ToList();
                buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
            }
            return PartialView("PartialPhongBanTree", buildTree);
        }


        public ActionResult GetQuanHuyen(string id)
        {
            try
            {
                context = new LinqNhanSuDataContext();
                var quanHuyens = context.Sys_QuanHuyens.Where(s => s.maTinhThanh == id)
                                        .Select(s => new { s.id, s.tenQuanHuyen }).ToList();
                return Json(quanHuyens);
            }
            catch
            {
                return View();
            }
        }

        public void GetDataFromView(FormCollection collection)
        {
            nhanVien.CMNDNgayCap = String.IsNullOrEmpty(collection["CMNDNgayCap"]) ? (DateTime?)null : DateTime.ParseExact(collection["CMNDNgayCap"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
            nhanVien.CMNDNoiCap = collection["CMNDNoiCap"];
            nhanVien.CMNDSo = collection["CMNDSo"];
            nhanVien.dangBaoHiemNoiKhac = collection["dangBaoHiemNoiKhac"].Contains("true") ? true : false;
            nhanVien.diaChiLienLacKhanCap = collection["diaChiLienLacKhanCap"];
            nhanVien.email = collection["email"];
            nhanVien.emailInd = collection["emailInd"];
            nhanVien.gioiTinh = collection["gioiTinh"] == "True" ? true : false;
            nhanVien.ho = collection["ho"];
            nhanVien.hoanTatHoSo = collection["hoanTatHoSo"].Contains("true") ? true : false;
            nhanVien.idDanToc = Convert.ToInt32(collection["idDanToc"]);
            nhanVien.idLoaiNhanVien = Convert.ToByte(collection["idLoaiNhanVien"]);
            nhanVien.idTonGiao = Convert.ToInt32(collection["idTonGiao"]);
            nhanVien.idTrinhDoChuyenMon = Convert.ToInt32(collection["idTrinhDoChuyenMon"]);
            nhanVien.laNguoiNgoaiQuoc = collection["laNguoiNgoaiQuoc"] == "True" ? true : false;
            //nhanVien.loaiThue = Convert.ToInt32(collection["loaiThue"]);
            nhanVien.maCapBacQL = collection["maCapBacQL"];
            nhanVien.maChamCong = collection["maChamCong"];
            nhanVien.maChiNhanhNganHang = collection["maChiNhanhNganHang"];
            nhanVien.idKhoiTinhLuong = String.IsNullOrEmpty(collection["idKhoiTinhLuong"]) ? (int?)null : Convert.ToInt32(collection["idKhoiTinhLuong"]);
            nhanVien.maQuocTich = collection["maQuocTich"];
            nhanVien.maSoThue = collection["maSoThue"];
            nhanVien.ngaySinh = String.IsNullOrEmpty(collection["ngaySinh"]) ? (DateTime?)null : DateTime.ParseExact(collection["ngaySinh"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
            nhanVien.ngayVaoLam = String.IsNullOrEmpty(collection["ngayVaoLam"]) ? (DateTime?)null : DateTime.ParseExact(collection["ngayVaoLam"], "dd/MM/yyyy", CultureInfo.InvariantCulture);

            nhanVien.ngayBatDauTinhPhep = String.IsNullOrEmpty(collection["ngayBatDauTinhPhepNam"]) ? (DateTime?)null : DateTime.ParseExact(collection["ngayBatDauTinhPhepNam"], "dd/MM/yyyy", CultureInfo.InvariantCulture);

            nhanVien.nguoiLienLacKhanCap = collection["nguoiLienLacKhanCap"];
            nhanVien.nguyenQuan = collection["nguyenQuan"];
            nhanVien.noiSinh = collection["noiSinh"];
            nhanVien.phoneLienLacKhanCap = collection["phoneLienLacKhanCap"];
            nhanVien.phoneNumber1 = collection["phoneNumber1"];
            nhanVien.phoneNumber2 = collection["phoneNumber2"];
            nhanVien.soBaoHiem = collection["soBaoHiem"];
            nhanVien.soSoLaoDong = collection["soSoLaoDong"];
            nhanVien.soTaiKhoan = collection["soTaiKhoan"];
            nhanVien.tamTruPhuongXa = collection["tamTruPhuongXa"];
            nhanVien.tamTruQuanHuyen = String.IsNullOrEmpty(collection["tamTruQuanHuyen"]) ? (int?)null : Convert.ToInt32(collection["tamTruQuanHuyen"]);
            nhanVien.tamTruSoNha = collection["tamTruSoNha"];
            nhanVien.tamTruTenDuong = collection["tamTruTenDuong"];
            nhanVien.tamTruTinhThanh = collection["tamTruTinhThanh"];
            nhanVien.ten = collection["ten"];
            nhanVien.tenAnhDaiDien = collection["tenAnhDaiDien"];
            nhanVien.tenAnhDaiDienLuu = collection["tenAnhDaiDienLuu"];
            //nhanVien.tenGoiKhac = collection["tenGoiKhac"];
            nhanVien.thuongTruPhuongXa = collection["thuongTruPhuongXa"];
            nhanVien.thuongTruQuanHuyen = String.IsNullOrEmpty(collection["thuongTruQuanHuyen"]) ? (int?)null : Convert.ToInt32(collection["thuongTruQuanHuyen"]);
            nhanVien.thuongTruSoNha = collection["thuongTruSoNha"];
            nhanVien.thuongTruTenDuong = collection["thuongTruTenDuong"];
            nhanVien.thuongTruTinhThanh = collection["thuongTruTinhThanh"];
            nhanVien.tinhTrangHonNhan = Convert.ToInt32(collection["tinhTrangHonNhan"]);
            nhanVien.maNgheNghiep = collection["maNgheNghiep"];
            nhanVien.loaiKetNap = collection["loaiKetNap"];
            nhanVien.trangThai = Convert.ToInt32(collection["trangThai"]);

            nhanVienChiNhanh = new tbl_NS_NhanVienChiNhanh();
            nhanVienChiNhanh.maChiNhanh = collection["maChiNhanhVanPhong"];
            nhanVienChiNhanh.maNhanVien = nhanVien.maNhanVien;
            nhanVienChiNhanh.ngayLap = DateTime.Now;
            nhanVienChiNhanh.nguoiLap = GetUser().manv;

            nhanVienChucDanh = new tbl_NS_NhanVienChucDanh();
            nhanVienChucDanh.maChucDanh = collection["maChucDanh"];
            nhanVienChucDanh.maNhanVien = nhanVien.maNhanVien;
            nhanVienChucDanh.ngayLap = DateTime.Now;
            nhanVienChucDanh.nguoiLap = GetUser().manv;

            nhanVienPhongBan = new tbl_NS_NhanVienPhongBan();
            nhanVienPhongBan.maNhanVien = nhanVien.maNhanVien;
            nhanVienPhongBan.maPhongBan = collection["maPhongBan"];
            nhanVienPhongBan.ngayLap = DateTime.Now;
            nhanVienPhongBan.nguoiLap = GetUser().manv;


        }

        public void GetAllDropdownlist()
        {
            var loaiNhanViens = context.tbl_NS_LoaiNhanViens;
            ViewBag.LoaiNhanViens = new SelectList(loaiNhanViens, "id", "tenLoaiNhanVien", model.idLoaiNhanVien);
            var listPhanCas = linqDM.tbl_NS_PhanCas.ToList();
            listPhanCas.Insert(0, new BatDongSan.Models.DanhMuc.tbl_NS_PhanCa { maPhanCa = 0, tenPhanCa = "Chọn phân ca" });
            ViewBag.ListPhanCas = new SelectList(listPhanCas, "maPhanCa", "tenPhanCa", model.maPhanCa);

            //var chiNhanhVPs = context.Sys_ChiNhanhVanPhongs.ToList();
            //chiNhanhVPs.Insert(0, new Sys_ChiNhanhVanPhong { maChiNhanh = "", tenChiNhanh = "" });
            //ViewBag.ChiNhanhVPs = new SelectList(chiNhanhVPs, "maChiNhanh", "tenChiNhanh", model.maChiNhanhVanPhong);

            var loaiThues = context.Sys_LoaiThues;
            ViewBag.LoaiThues = new SelectList(loaiThues, "id", "tenLoaiThue");

            var chiNhanhNganHangs = context.tbl_DM_ChiNhanhNganHangs.ToList();
            chiNhanhNganHangs.Insert(0, new BatDongSan.Models.NhanSu.tbl_DM_ChiNhanhNganHang { maChiNhanh = "", tenChiNhanh = "" });
            ViewBag.ChiNhanhNHs = new SelectList(chiNhanhNganHangs, "maChiNhanh", "tenChiNhanh", model.maChiNhanhNganHang);

            var boPhanTinhLuongs = context.GetTable<BatDongSan.Models.DanhMuc.tbl_NS_KhoiTinhLuong>();
            ViewBag.KhoiTinhLuongs = new SelectList(boPhanTinhLuongs, "maKhoiTinhLuong", "tenKhoiTinhLuong", model.idKhoiTinhLuong);

            var capBacQLs = context.Sys_CapBacQuanLies;
            ViewBag.CapBacQuanLys = new SelectList(capBacQLs, "maCapBac", "tenCapBac", model.maCapBacQL);

            var chucDanhs = context.Sys_ChucDanhs.OrderBy(o => o.TenChucDanh);
            ViewBag.ChucDanhs = new SelectList(chucDanhs, "MaChucDanh", "TenChucDanh", model.maChucDanh);

            var trinhDoChuyenMons = linqDM.tbl_DM_BangCaps;
            ViewBag.HocHams = new SelectList(trinhDoChuyenMons, "id", "tenBangCap", model.idTrinhDoChuyenMon);

            model.ngoaiNgus = context.Sys_ChungChiNgoaiNgus.ToList();

            var quocTichs = context.Sys_QuocTiches;
            ViewBag.QuocTichs = new SelectList(quocTichs, "maQuocTich", "tenQuocTich", model.maQuocTich);

            var danTocs = context.Sys_DanTocs;
            ViewBag.DanTocs = new SelectList(danTocs, "id", "ten", model.idDanToc);

            var tonGiaos = context.Sys_TonGiaos;
            ViewBag.TonGiaos = new SelectList(tonGiaos, "id", "ten", model.idTonGiao);

            var ngheNghieps = linqDM.tbl_DM_NgheNghieps.ToList();
            ngheNghieps.Insert(0, new tbl_DM_NgheNghiep { maNgheNghiep = "", tenNgheNghiep = "[Chọn ngành nghề]" });
            ViewBag.NgheNghieps = new SelectList(ngheNghieps, "maNgheNghiep", "tenNgheNghiep", model.maNgheNghiep);

            var loaiKetNaps = context.tbl_DM_LoaiKetNaps.ToList();
            loaiKetNaps.Insert(0, new tbl_DM_LoaiKetNap { maLoaiKetNap = "", tenLoaiKetNap = "" });
            ViewBag.LoaiKetNaps = new SelectList(loaiKetNaps, "maLoaiKetNap", "tenLoaiKetNap", model.loaiKetNap);

            var tinhThanhs = context.Sys_TinhThanhs.OrderBy(o => o.tenTinhThanh).ToList();
            tinhThanhs.Insert(0, new Sys_TinhThanh { maTinhThanh = "", tenTinhThanh = "[Chọn tỉnh thành]" });
            ViewBag.TamTruTinhThanhs = new SelectList(tinhThanhs, "maTinhThanh", "tenTinhThanh", model.tamTruTinhThanh);
            ViewBag.ThuongTruTinhThanhs = new SelectList(tinhThanhs, "maTinhThanh", "tenTinhThanh", model.thuongTruTinhThanh);

            var tamTruQuanHuyens = context.Sys_QuanHuyens.OrderBy(o => o.tenQuanHuyen).Where(s => s.maTinhThanh == model.tamTruTinhThanh);
            var thuongTruQuanHuyens = context.Sys_QuanHuyens.OrderBy(o => o.tenQuanHuyen).Where(s => s.maTinhThanh == model.thuongTruTinhThanh);
            ViewBag.TamTruQuanHuyens = new SelectList(tamTruQuanHuyens, "id", "tenQuanHuyen", model.tamTruQuanHuyen);
            ViewBag.ThuongTruQuanHuyens = new SelectList(thuongTruQuanHuyens, "id", "tenQuanHuyen", model.thuongTruQuanHuyen);
        }

        public void UploadAvartar(HttpPostedFileBase file)
        {
            if (file != null)
            {
                //if (!String.IsNullOrEmpty(nhanVien.anhDaiDienURL) && nhanVien.anhDaiDienURL != "/Images/card.gif")
                //{
                //    System.IO.File.Delete(Server.MapPath(nhanVien.anhDaiDienURL));
                //}
                //var date = DateTime.Now.ToString("yyyyMMdd-HHMMss");
                //string filePath = "/UploadFiles/Avartar/";

                ////Tạo tên mới cho file upload
                //string Generatedname = date.ToString() + file.FileName;
                //Directory.CreateDirectory(filePath);
                //var filePathOriginal = Server.MapPath(filePath);
                //string savedFileName = Path.Combine(filePathOriginal, Generatedname);
                //nhanVien.tenAnhDaiDien = file.FileName;
                //nhanVien.tenAnhDaiDienLuu = Generatedname;
                //nhanVien.anhDaiDienURL = filePath + Generatedname;
                ////nhanVien.anhDaiDienBinary = ImageHelper.ToByteArray(file);
                //file.SaveAs(savedFileName);
                //Lưu ảnh đại diện



                if (!string.IsNullOrEmpty(file.FileName) && file.ContentType.Contains("image"))
                {

                    DeleteFileOnServer();


                    UploadHelper fileHelper = new UploadHelper();

                    //fileHelper.PathSave = HttpContext.Server.MapPath("~/FileUploads/QLNV/");
                    //fileHelper.UploadFile(file);
                    string PathSave = HttpContext.Server.MapPath("~/UploadFiles/QLNV/");
                    string FileName = Path.GetFileName(file.FileName);
                    string FileSave = nhanVien.maNhanVien + Path.GetExtension(file.FileName);

                    if (!Directory.Exists(PathSave))
                    {
                        Directory.CreateDirectory(PathSave);
                    }
                    file.SaveAs(Path.Combine(PathSave, FileSave));
                    nhanVien.tenAnhDaiDienLuu = FileSave;
                    //fileHelper.FileSave;
                }

            }
        }
        private void DeleteFileOnServer()
        {
            FileInfo file = new FileInfo(HttpContext.Server.MapPath("~/UploadFiles/QLNV/") + nhanVien.maNhanVien + ".jpg");
            file.Refresh();
            if (file.Exists)
                file.Delete();
        }

        public void UploadFileDinhKem(FormCollection collection, IEnumerable<HttpPostedFileBase> files)
        {
            try
            {
                if (files != null)
                {
                    files = files.Where(s => s != null).OrderBy(o => o.FileName);
                    int i = 0;
                    string[] nameacceptable = collection.GetValues("nameaccept");
                    string[] thumbnails = collection.GetValues("thumbnail");
                    //Do files post về được sắp xếp theo tên file => Cần sắp xếp lại theo thứ tự mảng thumbnail theo tên (thumbnail có chứa tên file)
                    Array.Sort(thumbnails);
                    IList<Sys_FileDinhKem> fileDinhKems = new List<Sys_FileDinhKem>();
                    foreach (var file in files)
                    {
                        if (file != null)
                        {
                            Sys_FileDinhKem fileDinhKem = new Sys_FileDinhKem();
                            var date = DateTime.Now.ToString("yyyyMMdd-HHMMss");
                            string filePath = "/UploadFiles/NhanVien/";

                            //Tạo tên mới cho file upload
                            string Generatedname = date.ToString() + file.FileName;
                            Directory.CreateDirectory(filePath);
                            var filePathOriginal = Server.MapPath(filePath);
                            if (nameacceptable.Contains(file.FileName))
                            {
                                string savedFileName = Path.Combine(filePathOriginal, Generatedname);
                                fileDinhKem.controller = HtmlRequestHelper.GetRouteData().Controller;// Get Tên controller
                                fileDinhKem.Action = HtmlRequestHelper.GetRouteData().Action; //Get Tên Action
                                fileDinhKem.originalFileName = file.FileName;
                                fileDinhKem.savedFileName = Generatedname;
                                fileDinhKem.taiLieuURL = filePath + Generatedname;
                                fileDinhKem.identification = nhanVien.maNhanVien;
                                fileDinhKem.contentType = file.ContentType;
                                fileDinhKem.maNguoiUpLoad = GetUser().manv;
                                fileDinhKem.tenNguoiUpLoad = (string)Session["TenNhanVien"];
                                fileDinhKem.ngayLap = DateTime.Now;
                                if (Array.Exists(thumbnails, s => s.Contains(file.FileName)) && file.ContentType.Contains("image") == false)
                                {
                                    string thumbnail = thumbnails.Where(s => s.Contains(file.FileName)).First();
                                    //Tách giá trị value thumbnail thành 2 phần 1-Tên file, 2-Tên đường dẫn thumbnail
                                    fileDinhKem.thumbnailURL = Regex.Split(thumbnail, "-SplitPoint-").Last();
                                }
                                else
                                {
                                    if (file.ContentType.Contains("image"))
                                    {
                                        fileDinhKem.thumbnailURL = "/UploadFiles/NhanVien/" + Generatedname;
                                    }
                                }
                                fileDinhKems.Add(fileDinhKem);
                                file.SaveAs(savedFileName);
                                int Index = Array.IndexOf(nameacceptable, file.FileName);
                                Array.Clear(nameacceptable, Index, 1);
                            }
                        }
                        i++;
                    }
                    context.Sys_FileDinhKems.InsertAllOnSubmit(fileDinhKems);
                }
            }
            catch
            {
            }
        }


        public void SetModelData()
        {
            model = new NhanVienModel();
            model.anhDaiDienBinary = nhanVien.anhDaiDienBinary;
            model.anhDaiDienURL = nhanVien.anhDaiDienURL;
            model.CMNDNgayCap = nhanVien.CMNDNgayCap;
            model.CMNDNoiCap = nhanVien.CMNDNoiCap;
            model.CMNDSo = nhanVien.CMNDSo;
            model.dangBaoHiemNoiKhac = nhanVien.dangBaoHiemNoiKhac ?? false;
            model.diaChiLienLacKhanCap = nhanVien.diaChiLienLacKhanCap;
            model.email = nhanVien.email;
            model.gioiTinh = nhanVien.gioiTinh;
            model.ho = nhanVien.ho;
            model.hoanTatHoSo = nhanVien.hoanTatHoSo ?? false;
            model.idDanToc = nhanVien.idDanToc;
            model.idLoaiNhanVien = nhanVien.idLoaiNhanVien;
            model.idTonGiao = nhanVien.idTonGiao;
            model.idTrinhDoChuyenMon = nhanVien.idTrinhDoChuyenMon;
            model.laNguoiNgoaiQuoc = nhanVien.laNguoiNgoaiQuoc;
            model.loaiThue = nhanVien.loaiThue;
            model.maCapBacQL = nhanVien.maCapBacQL;
            model.maChamCong = nhanVien.maChamCong;
            model.maChiNhanhNganHang = nhanVien.maChiNhanhNganHang;
            model.idKhoiTinhLuong = Convert.ToInt32(nhanVien.idKhoiTinhLuong);
            model.maNhanVien = nhanVien.maNhanVien;
            model.maQuocTich = nhanVien.maQuocTich;
            model.maSoThue = nhanVien.maSoThue;
            model.ngaySinh = nhanVien.ngaySinh;
            model.ngayVaoLam = nhanVien.ngayVaoLam;

            model.ngayBatDauTinhPhepNam = nhanVien.ngayBatDauTinhPhep;

            model.nguoiLienLacKhanCap = nhanVien.nguoiLienLacKhanCap;
            model.nguyenQuan = nhanVien.nguyenQuan;
            model.noiSinh = nhanVien.noiSinh;
            model.phoneLienLacKhanCap = nhanVien.phoneLienLacKhanCap;
            model.phoneNumber1 = nhanVien.phoneNumber1;
            model.phoneNumber2 = nhanVien.phoneNumber2;
            model.soBaoHiem = nhanVien.soBaoHiem;
            model.soSoLaoDong = nhanVien.soSoLaoDong;
            model.soTaiKhoan = nhanVien.soTaiKhoan;
            model.tamTruPhuongXa = nhanVien.tamTruPhuongXa;
            model.tamTruQuanHuyen = nhanVien.tamTruQuanHuyen;
            model.tamTruSoNha = nhanVien.tamTruSoNha;
            model.tamTruTenDuong = nhanVien.tamTruTenDuong;
            model.tamTruTinhThanh = nhanVien.tamTruTinhThanh;
            model.ten = nhanVien.ten;
            model.tenAnhDaiDien = nhanVien.tenAnhDaiDien;
            model.tenAnhDaiDienLuu = nhanVien.tenAnhDaiDienLuu;
            //model.tenGoiKhac = nhanVien.tenGoiKhac;
            model.thuongTruPhuongXa = nhanVien.thuongTruPhuongXa;
            model.thuongTruQuanHuyen = nhanVien.thuongTruQuanHuyen;
            model.thuongTruSoNha = nhanVien.thuongTruSoNha;
            model.thuongTruTenDuong = nhanVien.thuongTruTenDuong;
            model.thuongTruTinhThanh = nhanVien.thuongTruTinhThanh;
            model.tinhTrangHonNhan = nhanVien.tinhTrangHonNhan;
            model.loaiKetNap = nhanVien.loaiKetNap;
            model.maNgheNghiep = nhanVien.maNgheNghiep;
            model.trangThai = nhanVien.trangThai;
            model.nhanVienNgoaiNgus = nhanVien.tbl_NS_NhanVienChungChiNgoaiNgus.ToList();
        }

        public void InsertChungChiNgoaiNgu(FormCollection collection)
        {
            var list = context.tbl_NS_NhanVienChungChiNgoaiNgus.Where(s => s.maNhanVien == nhanVien.maNhanVien);
            context.tbl_NS_NhanVienChungChiNgoaiNgus.DeleteAllOnSubmit(list);
            string[] arr = collection.GetValues("ngoaiNgu");
            ngoaiNgus = new List<tbl_NS_NhanVienChungChiNgoaiNgu>();
            for (int i = 0; i < arr.Length; i++)
            {
                if (arr[i] != "false")
                {
                    tbl_NS_NhanVienChungChiNgoaiNgu ngoaiNgu = new tbl_NS_NhanVienChungChiNgoaiNgu();
                    ngoaiNgu.idChungChi = Convert.ToInt32(arr[i]);
                    ngoaiNgu.maNhanVien = nhanVien.maNhanVien;
                    ngoaiNgus.Add(ngoaiNgu);
                }
            }
            context.tbl_NS_NhanVienChungChiNgoaiNgus.InsertAllOnSubmit(ngoaiNgus);
        }

        public void GetThongTinQuanHeGD(string maNhanVien)
        {
            try
            {
                model.quanHeGiaDinhs = (from q in context.tbl_NS_QuanHeGiaDinhs
                                        join l in context.tbl_DM_LoaiQuanHeGiaDinhs on q.idLoaiQuanHe equals l.id
                                        where q.maNhanVien == maNhanVien
                                        select new QuanHeGiaDinhModel
                                        {
                                            id = q.id,
                                            CMNDSo = q.CMNDSo,
                                            diaChiHienTai = q.diaChiHienTai,
                                            giamTruPhuThuoc = q.giamTruPhuThuoc,
                                            idLoaiQuanHe = q.idLoaiQuanHe,
                                            maNhanVien = q.maNhanVien,
                                            namSinh = q.namSinh,
                                            ngayBatDauGiam = q.ngayBatDauGiam,
                                            ngayCap = q.ngayCap,
                                            ngayKetThucGiam = q.ngayKetThucGiam,
                                            ngayLap = q.ngayLap,
                                            ngaySinh = q.ngaySinh,
                                            ngheNghiep = q.ngheNghiep,
                                            nguoiLap = q.nguoiLap,
                                            noiCap = q.noiCap,
                                            soTienGiam = q.soTienGiam,
                                            tenLoaiQuanHe = l.tenLoaiQuanHe,
                                            tenNguoiQuanHe = q.tenNguoiQuanHe
                                        }).ToList();
            }
            catch
            {
            }
        }


        public ActionResult QuanHeGiaDinh()
        {
            context = new LinqNhanSuDataContext();
            var loaiQuanHeGDs = context.GetTable<BatDongSan.Models.DanhMuc.tbl_DM_LoaiQuanHeGiaDinh>().ToList();
            loaiQuanHeGDs.Insert(0, new Models.DanhMuc.tbl_DM_LoaiQuanHeGiaDinh { id = 0, tenLoaiQuanHe = "[Chọn]" });
            ViewBag.LoaiQuanHeGDs = new SelectList(loaiQuanHeGDs, "id", "tenLoaiQuanHe");
            tbl_NS_QuanHeGiaDinh quanHeGD = new tbl_NS_QuanHeGiaDinh();
            return PartialView("PartialQuanHeGiaDinh", quanHeGD);
        }

        [HttpPost]
        public ActionResult QuanHeGiaDinh(FormCollection collection)
        {
            try
            {
                context = new LinqNhanSuDataContext();
                tbl_NS_QuanHeGiaDinh quanHeGD = new tbl_NS_QuanHeGiaDinh();
                quanHeGD.CMNDSo = collection["CMNDSo"];
                quanHeGD.diaChiHienTai = collection["diaChiHienTai"];
                quanHeGD.giamTruPhuThuoc = collection["giamTruPhuThuoc"].Contains("true") ? true : false;
                quanHeGD.idLoaiQuanHe = Convert.ToInt16(collection["idLoaiQuanHe"]);
                quanHeGD.maNhanVien = collection["maNhanVien"];
                quanHeGD.ngayBatDauGiam = String.IsNullOrEmpty(collection["ngayBatDauGiam"]) ? (DateTime?)null : DateTime.ParseExact(collection["ngayBatDauGiam"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                quanHeGD.ngayCap = String.IsNullOrEmpty(collection["ngayCap"]) ? (DateTime?)null : DateTime.ParseExact(collection["ngayCap"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                quanHeGD.ngayKetThucGiam = String.IsNullOrEmpty(collection["ngayKetThucGiam"]) ? (DateTime?)null : DateTime.ParseExact(collection["ngayKetThucGiam"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                quanHeGD.ngayLap = DateTime.Now;
                quanHeGD.ngaySinh = String.IsNullOrEmpty(collection["ngaySinh"]) ? (DateTime?)null : DateTime.ParseExact(collection["ngaySinh"], "dd/MM/yyyy", CultureInfo.InvariantCulture); ;
                quanHeGD.ngheNghiep = collection["ngheNghiep"];
                quanHeGD.nguoiLap = GetUser().manv;
                quanHeGD.noiCap = collection["noiCap"];
                quanHeGD.soTienGiam = String.IsNullOrEmpty(collection["soTienGiam"]) ? (decimal?)null : Convert.ToDecimal(collection["soTienGiam"]);
                quanHeGD.tenNguoiQuanHe = collection["tenNguoiQuanHe"];
                context.tbl_NS_QuanHeGiaDinhs.InsertOnSubmit(quanHeGD);
                context.SubmitChanges();
                model = new NhanVienModel();
                GetThongTinQuanHeGD(quanHeGD.maNhanVien);
                SaveActiveHistory("Cập nhật quan hệ gia đình: " + quanHeGD.maNhanVien);
                return View();
                //return RedirectToAction("Edit", new { id = quanHeGD.maNhanVien });
                //return PartialView("PartialEditQuanHeGD", model);
            }
            catch
            {
                return View();
            }
        }

        [HttpPost]
        public ActionResult DeleteQuanHeGiaDinh(int id)
        {
            try
            {
                using (context = new LinqNhanSuDataContext())
                {
                    var quanHeGiaDinh = context.tbl_NS_QuanHeGiaDinhs.Where(s => s.id == id).FirstOrDefault();
                    context.tbl_NS_QuanHeGiaDinhs.DeleteOnSubmit(quanHeGiaDinh);
                    SaveActiveHistory("Cập nhật quan hệ gia đình: " + id);
                    context.SubmitChanges();
                }
                return Json("");
            }
            catch
            {
                return View();
            }
        }

        [HttpPost]
        public ActionResult DeleteImage(int id)
        {
            try
            {
                context = new LinqNhanSuDataContext();
                var fileDinhkem = context.Sys_FileDinhKems.Where(s => s.id == id).FirstOrDefault();
                context.Sys_FileDinhKems.DeleteOnSubmit(fileDinhkem);
                context.SubmitChanges();
                System.IO.File.Delete(Server.MapPath(fileDinhkem.taiLieuURL));
                return Json(String.Empty);
            }
            catch
            {
                return View();
            }
        }

        public ActionResult Download(int id)
        {
            using (context = new LinqNhanSuDataContext())
            {
                var taiLieu = context.Sys_FileDinhKems.Where(s => s.id == id).FirstOrDefault();
                string savedFileName = Path.Combine("/UploadFiles/NhanVien/", taiLieu.savedFileName);
                return new DownloadResult { VirtualPath = savedFileName, FileDownloadName = taiLieu.originalFileName };
            }
        }

        public FileResult DownloadImportFile()
        {
            string savedFileName = Path.Combine("/UploadFiles/Template/", "PhieuNhapThongTinNhanVien.xlsx");
            return File(savedFileName, "multipart/form-data", "PhieuNhapThongTinNhanVien.xlsx");
        }

        [HttpPost]
        public ActionResult ImportExcelData(string excelPath)
        {
            try
            {
                string[] supportedFiles = { ".xlsx", ".xls" };
                HttpPostedFileBase File;
                File = Request.Files[0];
                if (File.ContentLength > 0)
                {
                    string extension = Path.GetExtension(File.FileName);
                    bool exist = Array.Exists(supportedFiles, element => element == extension);
                    if (exist == false)
                    {
                        return Json(new { success = false });
                    }
                    else
                    {
                        var date = DateTime.Now.ToString("yyyyMMdd-HHMMss");
                        string savedLocation = "/UploadFiles/NhanVien/";
                        Directory.CreateDirectory(savedLocation);
                        var filePath = Server.MapPath(savedLocation);
                        string fileName = date.ToString() + File.FileName;
                        string savedFileName = Path.Combine(filePath, fileName);
                        File.SaveAs(savedFileName);

                        ExcelDataProcessing excelDataProcessor = new ExcelDataProcessing(savedFileName);
                        DataTable dt = excelDataProcessor.GetDataTableWorkSheet("NhanVien");
                        IList<tbl_NS_NhanVien> nhanVienInports = new List<tbl_NS_NhanVien>();
                        IList<tbl_NS_NhanVienChiNhanh> chiNhanhImports = new List<tbl_NS_NhanVienChiNhanh>();
                        IList<tbl_NS_NhanVienChucDanh> chucDanhImports = new List<tbl_NS_NhanVienChucDanh>();
                        IList<tbl_NS_NhanVienPhongBan> phongBanImports = new List<tbl_NS_NhanVienPhongBan>();
                        foreach (DataRow row in dt.Rows)
                        {
                            if (String.IsNullOrEmpty(row["Mã nhân viên"].ToString()))
                                break;
                            tbl_NS_NhanVien nhanVienImport = new tbl_NS_NhanVien();
                            nhanVienImport.maNhanVien = row["Mã nhân viên"].ToString();
                            nhanVienImport.ho = row["Họ"].ToString();
                            nhanVienImport.ten = row["Tên"].ToString();
                            nhanVienImport.ngaySinh = String.IsNullOrEmpty(row["Ngày sinh"].ToString()) ? (DateTime?)null : DateTime.Parse(row["Ngày sinh"].ToString());
                            nhanVienImport.noiSinh = row["Nơi sinh"].ToString();
                            nhanVienImport.CMNDSo = row["Số CMND"].ToString();
                            nhanVienImport.CMNDNgayCap = String.IsNullOrEmpty(row["Ngày cấp CMND"].ToString()) ? (DateTime?)null : DateTime.Parse(row["Ngày cấp CMND"].ToString());
                            nhanVienImport.CMNDNoiCap = row["Nơi cấp CMND"].ToString();
                            nhanVienImport.ngayVaoLam = String.IsNullOrEmpty(row["Ngày vào làm"].ToString()) ? (DateTime?)null : DateTime.Parse(row["Ngày vào làm"].ToString());
                            nhanVienImport.email = row["Email"].ToString();
                            nhanVienImport.nguoiLap = GetUser().manv;
                            nhanVienImport.ngayLap = DateTime.Now;

                            tbl_NS_NhanVienChiNhanh chiNhanh = new tbl_NS_NhanVienChiNhanh();
                            chiNhanh.maChiNhanh = row["Mã chi nhánh văn phòng"].ToString();
                            chiNhanh.maNhanVien = nhanVienImport.maNhanVien;
                            chiNhanh.ngayLap = DateTime.Now;
                            chiNhanh.nguoiLap = GetUser().manv;
                            chiNhanhImports.Add(chiNhanh);

                            tbl_NS_NhanVienChucDanh chucDanh = new tbl_NS_NhanVienChucDanh();
                            chucDanh.maChucDanh = row["Mã chức danh"].ToString();
                            chucDanh.maNhanVien = nhanVienImport.maNhanVien;
                            chucDanh.ngayLap = DateTime.Now;
                            chucDanh.nguoiLap = GetUser().manv;
                            chucDanhImports.Add(chucDanh);

                            tbl_NS_NhanVienPhongBan phongBan = new tbl_NS_NhanVienPhongBan();
                            phongBan.maNhanVien = nhanVienImport.maNhanVien;
                            phongBan.maPhongBan = row["Mã phòng ban"].ToString();
                            phongBan.ngayLap = DateTime.Now;
                            phongBan.nguoiLap = GetUser().manv;
                            phongBanImports.Add(phongBan);


                            nhanVienImport.phoneNumber1 = row["Di động"].ToString();
                            nhanVienImport.phoneNumber2 = row["Điện thoại cố định"].ToString();
                            nhanVienInports.Add(nhanVienImport);
                        }
                        using (context = new LinqNhanSuDataContext())
                        {
                            context.tbl_NS_NhanViens.InsertAllOnSubmit(nhanVienInports);
                            context.tbl_NS_NhanVienChiNhanhs.InsertAllOnSubmit(chiNhanhImports);
                            context.tbl_NS_NhanVienPhongBans.InsertAllOnSubmit(phongBanImports);
                            context.tbl_NS_NhanVienChucDanhs.InsertAllOnSubmit(chucDanhImports);
                            context.SubmitChanges();
                        }
                        System.IO.File.Delete(Server.MapPath("/UploadFiles/NhanVien/" + fileName));
                    }
                }
                SaveActiveHistory("Import danh sách nhân viên");
                return Json(new { success = true });
            }
            catch
            {
                return View();
            }
        }

        #region Xuất excel
        public void XuatFileNhanVien()
        {
            var filename = "";
            var virtualPath = HttpRuntime.AppDomainAppVirtualPath;
            var fileStream = new FileStream(System.Web.HttpContext.Current.Server.MapPath(virtualPath + @"\Content\Report\ReportTemplate.xls"), FileMode.Open, FileAccess.Read);

            var workbook = new HSSFWorkbook(fileStream, true);
            filename += "DanhSachNhanVien.xls";


            var sheet = workbook.GetSheet("danhsachnhanvien");

            /*style title start*/
            //tạo font cho các title
            //font tiêu đề 
            HSSFFont hFontTieuDe = (HSSFFont)workbook.CreateFont();
            hFontTieuDe.FontHeightInPoints = 18;
            hFontTieuDe.Boldweight = 100 * 10;
            hFontTieuDe.FontName = "Times New Roman";
            hFontTieuDe.Color = HSSFColor.BLUE.index;

            //font tiêu đề 
            HSSFFont hFontTongGiaTriHT = (HSSFFont)workbook.CreateFont();
            hFontTongGiaTriHT.FontHeightInPoints = 11;
            hFontTongGiaTriHT.Boldweight = (short)FontBoldWeight.BOLD;
            hFontTongGiaTriHT.FontName = "Times New Roman";
            hFontTongGiaTriHT.Color = HSSFColor.BLACK.index;

            //font thông tin bảng tính
            HSSFFont hFontTT = (HSSFFont)workbook.CreateFont();
            hFontTT.IsItalic = true;
            hFontTT.Boldweight = (short)FontBoldWeight.BOLD;
            hFontTT.Color = HSSFColor.BLACK.index;
            hFontTT.FontName = "Times New Roman";
            hFontTieuDe.FontHeightInPoints = 11;

            //font chứ hoa đậm
            HSSFFont hFontNommalUpper = (HSSFFont)workbook.CreateFont();
            hFontNommalUpper.Boldweight = (short)FontBoldWeight.BOLD;
            hFontNommalUpper.Color = HSSFColor.BLACK.index;
            hFontNommalUpper.FontName = "Times New Roman";

            //font chữ bình thường
            HSSFFont hFontNommal = (HSSFFont)workbook.CreateFont();
            hFontNommal.Color = HSSFColor.BLACK.index;
            hFontNommal.FontName = "Times New Roman";

            //font chữ bình thường đậm
            HSSFFont hFontNommalBold = (HSSFFont)workbook.CreateFont();
            hFontNommalBold.Color = HSSFColor.BLACK.index;
            hFontNommalBold.Boldweight = (short)FontBoldWeight.BOLD;
            hFontNommalBold.FontName = "Times New Roman";

            //tạo font cho các title end

            //Set style
            var styleTitle = workbook.CreateCellStyle();
            styleTitle.SetFont(hFontTieuDe);
            styleTitle.Alignment = HorizontalAlignment.LEFT;

            //style infomation
            var styleInfomation = workbook.CreateCellStyle();
            styleInfomation.SetFont(hFontTT);
            styleInfomation.Alignment = HorizontalAlignment.LEFT;

            //style header
            var styleheadedColumnTable = workbook.CreateCellStyle();
            styleheadedColumnTable.SetFont(hFontNommalUpper);
            styleheadedColumnTable.WrapText = true;
            styleheadedColumnTable.BorderBottom = CellBorderType.THIN;
            styleheadedColumnTable.BorderLeft = CellBorderType.THIN;
            styleheadedColumnTable.BorderRight = CellBorderType.THIN;
            styleheadedColumnTable.BorderTop = CellBorderType.THIN;
            styleheadedColumnTable.VerticalAlignment = VerticalAlignment.CENTER;
            styleheadedColumnTable.Alignment = HorizontalAlignment.CENTER;

            var styleHeading1 = workbook.CreateCellStyle();
            styleHeading1.SetFont(hFontNommalBold);
            styleHeading1.WrapText = true;
            styleHeading1.BorderBottom = CellBorderType.THIN;
            styleHeading1.BorderLeft = CellBorderType.THIN;
            styleHeading1.BorderRight = CellBorderType.THIN;
            styleHeading1.BorderTop = CellBorderType.THIN;
            styleHeading1.VerticalAlignment = VerticalAlignment.CENTER;
            styleHeading1.Alignment = HorizontalAlignment.LEFT;

            var hStyleConLeft = (HSSFCellStyle)workbook.CreateCellStyle();
            hStyleConLeft.SetFont(hFontNommal);
            hStyleConLeft.VerticalAlignment = VerticalAlignment.TOP;
            hStyleConLeft.Alignment = HorizontalAlignment.LEFT;
            hStyleConLeft.WrapText = true;
            hStyleConLeft.BorderBottom = CellBorderType.THIN;
            hStyleConLeft.BorderLeft = CellBorderType.THIN;
            hStyleConLeft.BorderRight = CellBorderType.THIN;
            hStyleConLeft.BorderTop = CellBorderType.THIN;

            var hStyleConRight = (HSSFCellStyle)workbook.CreateCellStyle();
            hStyleConRight.SetFont(hFontNommal);
            hStyleConRight.VerticalAlignment = VerticalAlignment.TOP;
            hStyleConRight.Alignment = HorizontalAlignment.RIGHT;
            hStyleConRight.BorderBottom = CellBorderType.THIN;
            hStyleConRight.BorderLeft = CellBorderType.THIN;
            hStyleConRight.BorderRight = CellBorderType.THIN;
            hStyleConRight.BorderTop = CellBorderType.THIN;


            var hStyleConCenter = (HSSFCellStyle)workbook.CreateCellStyle();
            hStyleConCenter.SetFont(hFontNommal);
            hStyleConCenter.VerticalAlignment = VerticalAlignment.TOP;
            hStyleConCenter.Alignment = HorizontalAlignment.CENTER;
            hStyleConCenter.BorderBottom = CellBorderType.THIN;
            hStyleConCenter.BorderLeft = CellBorderType.THIN;
            hStyleConCenter.BorderRight = CellBorderType.THIN;
            hStyleConCenter.BorderTop = CellBorderType.THIN;
            //set style end


            Row rowC = null;
            //Khai báo row đầu tiên
            int firstRowNumber = 1;

            string rowtitle = "DANH SÁCH NHÂN VIÊN";
            var titleCell = HSSFCellUtil.CreateCell(sheet.CreateRow(firstRowNumber), 6, rowtitle.ToUpper());
            titleCell.CellStyle = styleTitle;

            ++firstRowNumber;

            var list1 = new List<string>();
            list1.Add("STT");
            list1.Add("Mã nhân viên");
            list1.Add("Tên nhân viên");
            list1.Add("Mã số thuế");
            list1.Add("CMND");
            list1.Add("Số điện thoại");
            list1.Add("Email");
            list1.Add("Ngày vào làm");
            list1.Add("Ngân hàng");
            list1.Add("Số tài khoản");
            list1.Add("Phòng ban");
            list1.Add("Loại");

            //Start row 13
            var headerRow = sheet.CreateRow(2);
            ReportHelperExcel.CreateHeaderRow(headerRow, 0, styleheadedColumnTable, list1);

            //Create header end
            context = new LinqNhanSuDataContext();
            var idRowStart = 3;
            var datas = context.sp_NS_NhanVien_Index(null, null, null, null, null).OrderBy(d => d.tenPhongBan).OrderBy(d => d.hoVaTen).ToList();
            //#region
            if (datas != null && datas.Count > 0)
            {
                var stt = 0;
                int dem = 0;
                //Giai đoạn
                foreach (var item in datas)
                {
                    dem = 0;

                    idRowStart++;
                    rowC = sheet.CreateRow(idRowStart);
                    ReportHelperExcel.SetAlignment(rowC, dem++, (++stt).ToString(), hStyleConCenter);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item.maNhanVien, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, (item.ho + " " + item.ten), hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item.maSoThue, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item.CMNDSo, hStyleConCenter);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item.phoneNumber1, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item.email, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:dd/MM/yyyy}", item.ngayVaoLam), hStyleConCenter);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item.tenChiNhanhNganHang, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item.soTaiKhoan, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item.tenPhongBan, hStyleConLeft);
                    ReportHelperExcel.SetAlignment(rowC, dem++, item.tenLoaiNhanVien, hStyleConLeft);
                    //ReportHelperExcel.SetAlignment(rowC, dem++, item.tenDonViTinh, hStyleConCenter);
                    //ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.###}", item.soLuongTonDau_BT), hStyleConRight);
                    //ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.###}", item.soLuongTonDau_HH), hStyleConRight);
                    //ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.###}", item.soLuongNhap_BT), hStyleConRight);
                    //ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.###}", item.soLuongNhap_HH), hStyleConRight);
                    //ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.###}", item.soLuongXuat_BT), hStyleConRight);
                    //ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.###}", item.soLuongXuat_HH), hStyleConRight);

                    //ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.###}", item.soLuongTonCuoi_BT), hStyleConRight);
                    //ReportHelperExcel.SetAlignment(rowC, dem++, String.Format("{0:#,##0.###}", item.soLuongTonCuoi_HH), hStyleConRight);
                }

                sheet.SetColumnWidth(0, 5 * 250);
                sheet.SetColumnWidth(1, 15 * 270);
                sheet.SetColumnWidth(2, 20 * 250);
                sheet.SetColumnWidth(3, 25 * 210);
                sheet.SetColumnWidth(4, 15 * 210);
                sheet.SetColumnWidth(5, 15 * 210);
                sheet.SetColumnWidth(6, 30 * 210);
                sheet.SetColumnWidth(7, 30 * 210);
                sheet.SetColumnWidth(8, 30 * 210);
                sheet.SetColumnWidth(9, 30 * 210);
                sheet.SetColumnWidth(10, 20 * 210);
                sheet.SetColumnWidth(11, 15 * 210);
            }
            else
            {

                sheet.SetColumnWidth(0, 5 * 250);
                sheet.SetColumnWidth(1, 15 * 270);
                sheet.SetColumnWidth(2, 20 * 250);
                sheet.SetColumnWidth(3, 15 * 210);
                sheet.SetColumnWidth(4, 15 * 210);
                sheet.SetColumnWidth(5, 30 * 210);
                sheet.SetColumnWidth(6, 15 * 210);
                sheet.SetColumnWidth(7, 30 * 210);
                sheet.SetColumnWidth(8, 30 * 210);
                sheet.SetColumnWidth(9, 20 * 210);
                sheet.SetColumnWidth(10, 15 * 210);

            }

            var stream = new MemoryStream();
            workbook.Write(stream);

            Response.ContentType = "application/vnd.ms-excel";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));
            Response.Clear();

            Response.BinaryWrite(stream.GetBuffer());
            Response.End();

        }
        #endregion
    }
}
