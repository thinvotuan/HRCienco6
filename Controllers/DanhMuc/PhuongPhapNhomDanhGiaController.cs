using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Common;
using BatDongSan.Helper.Utils;
using BatDongSan.Models.DanhMuc;
using BatDongSan.Models.PhieuDeNghi;
using BatDongSan.Utils.Paging;


namespace BatDongSan.Controllers.DanhMuc
{
    public class PhuongPhapNhomDanhGiaController : ApplicationController
    {
        LinqPhieuDeNghiDataContext lqPhieuDN = new LinqPhieuDeNghiDataContext();
        LinqDanhMucDataContext linqDM = new LinqDanhMucDataContext();

        IList<sp_PhuongPhapNhomDanhGia_IndexResult> index;
        public const string taskIDSystem = "PPNhomDanhGia";//REGWORKVOTE
        IList<sp_DanhMucNhomDanhGia_IndexResult> indexNhom;
        public bool? permission;

        tbl_DM_PhuongPhapNhomDanhGia phuonPhapDG;
        //
        // GET: /PhuongPhapNhomDanhGia/

        #region Index

        public ActionResult Index(int? page, int? pageSize, string searchString)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXem);
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
                index = linqDM.sp_PhuongPhapNhomDanhGia_Index(searchString, currentPageIndex, pageSize).ToList();
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
        #endregion

        #region Create, Edit, Details
        public ActionResult Create()
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
                phuonPhapDG = new tbl_DM_PhuongPhapNhomDanhGia();
                phuonPhapDG.maPhuongPhapDanhGia = GenerateUtil.CheckLetter("PPNDG", GetMax());
                phuonPhapDG.ngayLap = DateTime.Now;
                phuonPhapDG.nguoiLap = GetUser().manv;
                phuonPhapDG.tenPhuongPhapDanhGia = string.Empty;
                ViewBag.TenNhanVien = (string)Session["TenNhanVien"];
                return View(phuonPhapDG);
            }
            catch (Exception e)
            {
                ViewData["Message"] = e.Message;
                return View("Error");
            }
        }

        [HttpPost]
        public ActionResult Create(FormCollection collection)
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
                BindDataToSave(collection, true);
                linqDM.tbl_DM_PhuongPhapNhomDanhGias.InsertOnSubmit(phuonPhapDG);
                linqDM.SubmitChanges();
                return RedirectToAction("Edit/" + phuonPhapDG.maPhuongPhapDanhGia);
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
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            try
            {
                GetData(id);
                return View(phuonPhapDG);
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
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            try
            {
                BindDataToSave(collection, false);
                linqDM.SubmitChanges();
                return RedirectToAction("Edit/" + phuonPhapDG.maPhuongPhapDanhGia);
            }
            catch (Exception e)
            {
                ViewData["Message"] = e.Message;
                return View("Error");
            }
        }


        public ActionResult Details(string id)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXemChiTiet);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion
            try
            {
                GetData(id);
                return View(phuonPhapDG);
            }
            catch (Exception e)
            {
                ViewData["Message"] = e.Message;
                return View("Error");
            }
        }

        [HttpPost]
        public ActionResult Delete(string id)
        {
            try
            {
                //Cập nhật trạng thái xóa phiếu
                var delete = linqDM.tbl_DM_PhuongPhapNhomDanhGias.Where(d => d.maPhuongPhapDanhGia == id);
                linqDM.tbl_DM_PhuongPhapNhomDanhGias.DeleteAllOnSubmit(delete);

                //Delete chi tiết
                var deleteCT = linqDM.tbl_DM_PhuongPhapNhomDanhGia_ChiTiets.Where(d => d.maPhuongPhapDanhGia == id);
                linqDM.tbl_DM_PhuongPhapNhomDanhGia_ChiTiets.DeleteAllOnSubmit(deleteCT);
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
        /// Insert vào phương pháp đánh giá
        /// </summary>
        /// <param name="col"></param>
        /// <param name="isCreate"></param>
        public void BindDataToSave(FormCollection col, bool isCreate)
        {
            if (isCreate == true)
            {
                phuonPhapDG = new tbl_DM_PhuongPhapNhomDanhGia();
                phuonPhapDG.maPhuongPhapDanhGia = GenerateUtil.CheckLetter("PPNDG", GetMax());
                phuonPhapDG.nguoiLap = GetUser().manv;
                phuonPhapDG.ngayLap = DateTime.Now;
            }
            else
            {
                phuonPhapDG = linqDM.tbl_DM_PhuongPhapNhomDanhGias.Where(d => d.maPhuongPhapDanhGia == col.Get("maPhuongPhapDanhGia")).FirstOrDefault();
            }
            phuonPhapDG.ghiChu = col.Get("ghiChu");
            phuonPhapDG.tenPhuongPhapDanhGia = col.Get("tenPhuongPhapDanhGia");
            //Inser qui trình duyệt chi tiết
            var delNhomPhongBanChiTiet = linqDM.tbl_DM_PhuongPhapNhomDanhGia_ChiTiets.Where(d => d.maPhuongPhapDanhGia == col.Get("maPhuongPhapDanhGia"));
            linqDM.tbl_DM_PhuongPhapNhomDanhGia_ChiTiets.DeleteAllOnSubmit(delNhomPhongBanChiTiet);
            List<tbl_DM_PhuongPhapNhomDanhGia_ChiTiet> ct = new List<tbl_DM_PhuongPhapNhomDanhGia_ChiTiet>();
            tbl_DM_PhuongPhapNhomDanhGia_ChiTiet danhGia;
            string[] maNhomDanhGia = col.GetValues("maNhomDanhGia");
            if (maNhomDanhGia != null)
            {
                for (int i = 0; i < maNhomDanhGia.Count(); i++)
                {

                    danhGia = new tbl_DM_PhuongPhapNhomDanhGia_ChiTiet();
                    danhGia.maPhuongPhapDanhGia = phuonPhapDG.maPhuongPhapDanhGia;
                    danhGia.maNhomDanhGia = col.GetValues("maNhomDanhGia")[i];
                    danhGia.tenNhomDanhGia = col.GetValues("tenNhomDanhGia")[i];
                    danhGia.soDiem = (string.IsNullOrEmpty(col.GetValues("soDiem")[i]) ? 0 : Convert.ToDecimal(col.GetValues("soDiem")[i]));
                    danhGia.thuHang = (string.IsNullOrEmpty(col.GetValues("thuHang")[i]) ? 0 : Convert.ToDecimal(col.GetValues("thuHang")[i]));
                    ct.Add(danhGia);
                }
                if (ct != null && ct.Count > 0)
                {
                    linqDM.tbl_DM_PhuongPhapNhomDanhGia_ChiTiets.InsertAllOnSubmit(ct);
                }
            }
        }


        /// <summary>
        /// Thông tin dữ liệu qui trình duyệt
        /// </summary>
        /// <param name="id"></param>
        public void GetData(string id)
        {
            phuonPhapDG = linqDM.tbl_DM_PhuongPhapNhomDanhGias.Where(d => d.maPhuongPhapDanhGia == id).FirstOrDefault();
            ViewBag.TenNhanVien = HoVaTen(phuonPhapDG.nguoiLap);
            //Chi tiết qui trình duyệt
            ViewBag.ChiTiet = linqDM.tbl_DM_PhuongPhapNhomDanhGia_ChiTiets.Where(d => d.maPhuongPhapDanhGia == id).OrderBy(d=>d.thuHang).Select(d => new PhuongPhapDanhGiaChiTietModel
            {
                MaPhuongPhapDanhGia = d.maPhuongPhapDanhGia,
                TenNhomDanhGia = d.tenNhomDanhGia,
                MaNhomDanhGia = d.maNhomDanhGia,
                SoDiem =Convert.ToDouble(d.soDiem ?? 0),
                ThuHang =Convert.ToInt32(d.thuHang ?? 0),
            }).ToList();
        }

        #endregion

        /// <summary>
        /// Hàm get max nhóm đánh giá
        /// </summary>
        /// <returns></returns>
        public string GetMax()
        {
            return linqDM.tbl_DM_PhuongPhapNhomDanhGias.OrderByDescending(d=>d.ngayLap).Select(d => d.maPhuongPhapDanhGia).FirstOrDefault()??String.Empty;
        }

        #region Danh sách nhóm đánh giá

        /// <summary>
        /// Danh sách thông tin nhóm đánh giá
        /// </summary>
        /// <returns></returns>
        /// DanhSachNhomDanhGia

        public ActionResult DanhSachNhomDanhGia(int? page, int? pageSize, string searchString)
        {
            #region Role user
            permission = GetPermission(taskIDSystem, BangPhanQuyen.QuyenXem);
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
                indexNhom = linqDM.sp_DanhMucNhomDanhGia_Index(searchString, currentPageIndex, pageSize).ToList();
                try
                {
                    ViewBag.Count = indexNhom[0].tongDong;
                    tongSoDong = indexNhom[0].tongDong;
                }
                catch
                {
                    ViewBag.Count = 0;
                }
                TempData["Params"] = searchString + ",";
                if (Request.IsAjaxRequest())
                {
                    return PartialView("LoadDanhGia", indexNhom.ToPagedList(currentPageIndex, 30, true, tongSoDong));
                }
                return View(indexNhom.ToPagedList(currentPageIndex, 30, true, tongSoDong));
            }
            catch (Exception e)
            {
                ViewData["Message"] = e.Message;
                return View("Error");
            }
        }

        #endregion
    }
}
