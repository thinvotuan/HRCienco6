using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Helper.Utils;
using BatDongSan.Models.DanhMuc;
using BatDongSan.Models.HeThong;
using BatDongSan.Models.NhanSu;
using BatDongSan.Utils;
using BatDongSan.Utils.Paging;
namespace BatDongSan.Controllers.DanhMuc
{
    public class DMNhomDanhGiaController : ApplicationController
    {
        private StringBuilder buildTree = null;
        private IList<tbl_DM_PhongBan> phongBans;
        LinqDanhMucDataContext linqDM = new LinqDanhMucDataContext();
        tbl_DM_NhomDanhGia nhomDanhGia;
        readonly string MCV = "DMNhomDanhGia";
        bool? permission;
        IList<sp_DanhMucNhomDanhGia_IndexResult> index;
        public ActionResult Index(int? page, int? pageSize, string searchString)
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
                int currentPageIndex = page.HasValue ? page.Value : 1;
                pageSize = pageSize ?? 30;
                int? tongSoDong = 0;
                index = linqDM.sp_DanhMucNhomDanhGia_Index(searchString, currentPageIndex, pageSize).ToList();
                try
                {
                    ViewBag.Count = index[0].tongDong;
                    tongSoDong = index[0].tongDong;
                }
                catch
                {
                    ViewBag.Count = 0;
                }
                TempData["Params"] = searchString + ",";
                if (Request.IsAjaxRequest())
                {
                    return PartialView("ViewIndex", index.ToPagedList(currentPageIndex, 30, true, tongSoDong));
                }
                return View(index.ToPagedList(currentPageIndex, 30, true, tongSoDong));
            }
            catch (Exception e)
            {
                ViewData["Message"] = e.Message;
                return View("Error");
            }
        }

        //
        // GET: /LoaiQuanHeGiaDinh/Details/5

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
                GetData(id);
                return View(nhomDanhGia);
            }
            catch (Exception e)
            {
                ViewData["Message"] = e.Message;
                return View("Error");
            }
        }

        //
        // GET: /LoaiQuanHeGiaDinh/Create

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
                nhomDanhGia = new tbl_DM_NhomDanhGia();
                nhomDanhGia.maNhomDanhGia = GenerateUtil.CheckLetter("MNDG", GetMax());
                nhomDanhGia.ngayLap = DateTime.Now;
                nhomDanhGia.nguoiLap = GetUser().manv;
                nhomDanhGia.tenNhomDanhGia = string.Empty;
                ViewBag.TenNhanVien = (string)Session["TenNhanVien"];
                return View(nhomDanhGia);
            }
            catch (Exception e)
            {
                ViewData["Message"] = e.Message;
                return View("Error");
            }
        }

        //
        // POST: /LoaiQuanHeGiaDinh/Create

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
                BindDataToSave(collection, true);
                linqDM.tbl_DM_NhomDanhGias.InsertOnSubmit(nhomDanhGia);
                linqDM.SubmitChanges();
                return RedirectToAction("Edit/" + nhomDanhGia.maNhomDanhGia);
            }
            catch (Exception e)
            {
                ViewData["Message"] = e.Message;
                return View("Error");
            }
        }

        //
        // GET: /LoaiQuanHeGiaDinh/Edit/5

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
                GetData(id);
                return View(nhomDanhGia);
            }
            catch (Exception e)
            {
                ViewData["Message"] = e.Message;
                return View("Error");
            }
        }

        //
        // POST: /LoaiQuanHeGiaDinh/Edit/5

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
                BindDataToSave(collection, false);
                linqDM.SubmitChanges();
                return RedirectToAction("Edit/" + nhomDanhGia.maNhomDanhGia);
            }
            catch (Exception e)
            {
                ViewData["Message"] = e.Message;
                return View("Error");
            }
        }


        //
        // POST: /LoaiQuanHeGiaDinh/Delete/5

        [HttpPost]
        public ActionResult Delete(string id)
        {
            try
            {
                //Cập nhật trạng thái xóa phiếu
                var delete = linqDM.tbl_DM_NhomDanhGias.Where(d => d.maNhomDanhGia == id);
                linqDM.tbl_DM_NhomDanhGias.DeleteAllOnSubmit(delete);

                //Delete chi tiết
                var deleteCT = linqDM.tbl_DM_NhomDanhGia_PhongBans.Where(d => d.maNhomDanhGia == id);
                linqDM.tbl_DM_NhomDanhGia_PhongBans.DeleteAllOnSubmit(deleteCT);
                linqDM.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch (Exception e)
            {
                ViewData["Message"] = e.Message;
                return View("Error");
            }
        }

        /// <summary>
        /// Insert vào qui trình duyệt phiếu
        /// </summary>
        /// <param name="col"></param>
        /// <param name="isCreate"></param>
        public void BindDataToSave(FormCollection col, bool isCreate)
        {
            if (isCreate == true)
            {
                nhomDanhGia = new tbl_DM_NhomDanhGia();
                nhomDanhGia.maNhomDanhGia = GenerateUtil.CheckLetter("MNDG", GetMax());
                nhomDanhGia.nguoiLap = GetUser().manv;
                nhomDanhGia.ngayLap = DateTime.Now;
            }
            else
            {
                nhomDanhGia = linqDM.tbl_DM_NhomDanhGias.Where(d => d.maNhomDanhGia == col.Get("maNhomDanhGia")).FirstOrDefault();
            }
            nhomDanhGia.ghiChu = col.Get("ghiChu");
            nhomDanhGia.tenNhomDanhGia = col.Get("tenNhomDanhGia");
            //Inser qui trình duyệt chi tiết
            var delNhomPhongBanChiTiet = linqDM.tbl_DM_NhomDanhGia_PhongBans.Where(d => d.maNhomDanhGia == col.Get("maNhomDanhGia"));
            linqDM.tbl_DM_NhomDanhGia_PhongBans.DeleteAllOnSubmit(delNhomPhongBanChiTiet);
            List<tbl_DM_NhomDanhGia_PhongBan> ct = new List<tbl_DM_NhomDanhGia_PhongBan>();
            tbl_DM_NhomDanhGia_PhongBan phongBan;
            string[] maPhongBan = col.GetValues("maPhongBan");
            if (maPhongBan != null)
            {
                for (int i = 0; i < maPhongBan.Count(); i++)
                {

                    phongBan = new tbl_DM_NhomDanhGia_PhongBan();
                    phongBan.maPhongBan = col.GetValues("maPhongBan")[i];
                    phongBan.maNhomDanhGia = nhomDanhGia.maNhomDanhGia;
                    ct.Add(phongBan);
                }
                if (ct != null && ct.Count > 0)
                {
                    linqDM.tbl_DM_NhomDanhGia_PhongBans.InsertAllOnSubmit(ct);
                }
            }
        }

        /// <summary>
        /// Thông tin dữ liệu qui trình duyệt
        /// </summary>
        /// <param name="id"></param>
        public void GetData(string id)
        {

            nhomDanhGia = linqDM.tbl_DM_NhomDanhGias.Where(d => d.maNhomDanhGia == id).FirstOrDefault();
            ViewBag.TenNhanVien = HoVaTen(nhomDanhGia.nguoiLap);
            //Chi tiết qui trình duyệt
            ViewBag.DanhSachiPB = linqDM.tbl_DM_NhomDanhGia_PhongBans.Where(d => d.maNhomDanhGia == id).Select(d => new PhongBanModel
            {
                MaPhongBan = d.maPhongBan,
                maNhomDanhGia = d.maNhomDanhGia,
                Ten = linqDM.tbl_DM_PhongBans.Where(c => c.maPhongBan == d.maPhongBan).Select(c => c.tenPhongBan).FirstOrDefault() ?? string.Empty,
            }).ToList();
        }

        /// <summary>
        /// Danh sách phòng ban
        /// </summary>
        /// <returns></returns>
        public ActionResult DanhSachNVPB(int stt)
        {
            try
            {
                StringBuilder buildTree = new StringBuilder();
                var phongBans = linqDM.tbl_DM_PhongBans.ToList();
                buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
                ViewBag.NVPB = buildTree.ToString();
                ViewBag.STT = stt;
                return View();
            }
            catch (Exception ex)
            {

                ViewData["Message"] = ex.Message;
                return View("error");
            }
        }

        /// <summary>
        /// Nhân viên theo phòng ban
        /// </summary>
        /// <param name="page"></param>
        /// <param name="searchString"></param>
        /// <param name="maPhongBan"></param>
        /// <returns></returns>
        public ActionResult LoadNhanVien(int? page, string searchString, string maPhongBan, int stt)
        {
            IList<sp_PB_DanhSachNhanVienResult> phongBan1s;
            phongBan1s = linqDM.sp_PB_DanhSachNhanVien(searchString, maPhongBan).ToList();
            ViewBag.isGet = "True";
            int currentPageIndex = page.HasValue ? page.Value : 1;
            ViewBag.Count = phongBan1s.Count();
            ViewBag.Search = searchString;
            ViewBag.MaPhongBan = maPhongBan;
            ViewBag.STT = stt;
            return PartialView("_LoadNhanVien", phongBan1s.ToPagedList(currentPageIndex, 10));
        }

        /// <summary>
        /// Hàm get max tỉ lệ thành tích
        /// </summary>
        /// <returns></returns>
        public string GetMax()
        {
            return linqDM.tbl_DM_NhomDanhGias.OrderByDescending(d=>d.ngayLap).Select(d => d.maNhomDanhGia).FirstOrDefault()??String.Empty;
        }

        #region Cây sơ đồ tổ chức nhân viên và phòng ban

        //Load danh sach nhân viên khi click button add them:
        public ActionResult DanhSachPhongBan()
        {
            try
            {
                buildTree = new StringBuilder();
                phongBans = linqDM.tbl_DM_PhongBans.ToList();
                buildTree = TreePhongBans.BuildTreeDepartment(phongBans);
                ViewBag.departments = buildTree.ToString();
                ViewBag.shiftType = linqDM.tbl_NS_PhanCas.Where(t => t.tenPhanCa != "").ToList();
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
        #endregion
    }
}
