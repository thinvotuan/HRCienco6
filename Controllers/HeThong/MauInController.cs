using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.Xml;
using BatDongSan.Helper;
using BatDongSan.Helper.Common;
using BatDongSan.Helper.Utils;
using BatDongSan.Models.HeThong;
using BatDongSan.Utils.Paging;

namespace BatDongSan.Controllers.HeThong
{
    public class MauInController : ApplicationController
    {
        private LinqHeThongDataContext kd = new LinqHeThongDataContext();
        //
        // GET: /MauIn/
        private readonly string MCV = "MAUIN";
        private bool? permission;
        //private static int defaultPageSize = 10;
        
        private StringBuilder sb = new StringBuilder();
        private DMMauInHopDong record;
        private Sys_MauIn mauIn;
        private IList<sp_KD_MauIn_IndexResult> mauIns;
        public ActionResult Index(int? page, int? pageSize, string tuNgay, string denNgay, string qSearch, int? trangThaiDuyet)
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
                ViewBag.MaNguoiLap = GetUser().manv;
                //Trạng thái duyệt
                //TrangThaiDuyet(MCV);
                //Danh sách hồ sơ thanh lý hợp đồng         
                DateTime? fromDate = String.IsNullOrEmpty(tuNgay) ? (DateTime?)null : DateTime.ParseExact(tuNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                DateTime? toDate = String.IsNullOrEmpty(denNgay) ? (DateTime?)null : DateTime.ParseExact(denNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);

                //int currentPageIndex = page.HasValue ? page.Value : 1;
                //defaultPageSize = (pageSize ?? defaultPageSize);
                mauIns = kd.sp_KD_MauIn_Index(qSearch, fromDate, toDate, trangThaiDuyet).Skip(0).Take(50).ToList();
                //TempData["Params"] = qSearch + "," + fromDate + "," + toDate;
                //ViewBag.Count = mauIns.Count;
                //ViewBag.qSearch = qSearch ?? string.Empty;
                //ViewBag.MaNV = GetUser().manv;
                if (Request.IsAjaxRequest())
                {
                    return PartialView("ViewIndex", mauIns);
                }
                return View(mauIns);
            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.ToString();
                return View("Failed");
            }
        }

        public ActionResult Create(string id, string msg)
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
                mauIn = new Sys_MauIn();
                mauIn.MaMauIn = IdGenerator();
                //GetDuAnDropdownList(mauIn.MaDuAn);
                getLoaiMauInDropDownlist();
                if (msg == "true")
                {
                    ViewBag.File = "File không đúng định dạnh xml";
                }
                else
                {
                    ViewBag.File = string.Empty;
                }
                return View(mauIn);

            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.ToString();
                return View("Failed");
            }
        }

        [HttpPost]
        [ValidateInput(false)]
        public ActionResult Create(FormCollection collection, HttpPostedFileBase files)
        {
            try
            {
                // TODO: Add insert logic here             
                record = new DMMauInHopDong();
                mauIn = new Sys_MauIn();
                BindDataToSave(collection, true, files);
                if (String.IsNullOrEmpty(mauIn.TenFileDinhKem))
                {
                   // GetDuAnDropdownList(mauIn.MaDuAn);
                    getLoaiMauInDropDownlist();
                    return View(mauIn);
                }
                kd.SubmitChanges();
                return RedirectToAction("Edit", "MauIn", new { id = mauIn.MaMauIn });
            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.ToString();
                return View("Failed");
            }
        }

        public ActionResult Edit(string id, string msg)
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
                mauIn = kd.Sys_MauIns.Where(d => d.MaMauIn == id).FirstOrDefault();
                ThongTinLoaiVanBan();
               
                getLoaiMauInDropDownlist();                
                if (msg == "true")
                {
                    ViewBag.File = "File không đúng định dạnh xml";
                }
                else
                {
                    ViewBag.File = string.Empty;
                }
                return View(mauIn);

            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.ToString();
                return View("Failed");
            }
        }

        [HttpPost]
        [ValidateInput(false)]
        public ActionResult Edit(FormCollection collection, HttpPostedFileBase files)
        {
            try
            {
                // TODO: Add insert logic here             
                record = new DMMauInHopDong();
                mauIn = new Sys_MauIn();
                BindDataToSave(collection, false, files);
                if (String.IsNullOrEmpty(mauIn.TenFileDinhKem))
                {
                    return RedirectToAction("Edit", new { id = mauIn.MaMauIn });
                }
                kd.SubmitChanges();
                return RedirectToAction("Edit", "MauIn", new { id = mauIn.MaMauIn });
            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.ToString();
                return View("Failed");
            }
        }

        [HttpPost]
        public ActionResult Delete(string[] checkeds)
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
                if (checkeds != null)
                {
                    var hdVay = kd.Sys_MauIns.Where(d => checkeds.Contains(d.MaMauIn)).ToList();
                    kd.Sys_MauIns.DeleteAllOnSubmit(hdVay);

                    //Xóa người duyệt
                    var nguoiDuyet = kd.GetTable<BatDongSan.Models.HeThong.tbl_HT_DMNguoiDuyet>().Where(d => checkeds.Contains(d.maPhieu));
                    kd.GetTable<BatDongSan.Models.HeThong.tbl_HT_DMNguoiDuyet>().DeleteAllOnSubmit(nguoiDuyet);
                    kd.SubmitChanges();
                }
                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.ToString();
                return View("Failed");
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
            try
            {
                mauIn = kd.Sys_MauIns.Where(d => d.MaMauIn == id).FirstOrDefault();
                ThongTinLoaiVanBan();
               
                return View(mauIn);

            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.ToString();
                return View("Failed");
            }
        }

        public void BindDataToSave(FormCollection col, bool isCreate, HttpPostedFileBase file)
        {
            if (isCreate == true)
            {
                mauIn.MaMauIn = IdGenerator();
                mauIn.TenMau = col.Get("TenMau");
                mauIn.MaLoaiVanBan = Convert.ToInt32(col.Get("MaLoaiVanBan"));
                mauIn.NgayTao = DateTime.Now;
                mauIn.NguoiLap = GetUser().manv;
                mauIn.MaDuAn = col.Get("MaDuAn");
                mauIn.NgayApDung = DateTime.ParseExact(col["NgayApDung"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                UploadFile(file);
                kd.Sys_MauIns.InsertOnSubmit(mauIn);
                //if (UploadFile(col) == "true")
                //{                    
                //}
                //else
                //{
                //    mauIn.TenFileDinhKem = "false";
                //}
            }
            else
            {
                mauIn = kd.Sys_MauIns.Where(d => d.MaMauIn == col.Get("MaMauIn")).FirstOrDefault();
                mauIn.MaLoaiVanBan = Convert.ToInt32(col.Get("MaLoaiVanBan"));
                mauIn.TenMau = col.Get("TenMau");
                mauIn.MaDuAn = col.Get("MaDuAn");
                mauIn.NgayApDung = DateTime.ParseExact(col["NgayApDung"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                UploadFile(file);
                //if (UploadFile(col) == "true")
                //{
                //}
                //else
                //{
                //    mauIn.TenFileDinhKem = "false";
                //}
            }
        }

        #region Upload file
        /// <summary>
        /// Thin.VoTuan - 20/04/2016
        /// Lấy dữ liệu uplod file
        /// </summary>

        private bool UploadFile(HttpPostedFileBase file)
        {
            try
            {
                if (file != null)
                {
                    string mimeType = file.ContentType;
                    if (string.Equals(mimeType, "text/xml"))
                    {
                        UploadHelper fileHelper = new UploadHelper();
                        fileHelper.PathSave = HttpContext.Server.MapPath("~/FileUploads/MKinhDoanh/MauInHopDong/");
                        fileHelper.UploadFile(file);
                        mauIn.TenFileDinhKem = fileHelper.FileName;
                        mauIn.TenFileDinhKemLuu = fileHelper.FileSave;
                        XmlDocument doc = new XmlDocument();
                        doc.Load(fileHelper.PathSave + fileHelper.FileSave);
                        string xmlcontents1 = doc.OuterXml;
                        mauIn.NoiDung = xmlcontents1;
                    }
                    else
                    {
                        TempData["mssg"] = "Định đạng file đính kèm không đúng";
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }

        }

        public ActionResult Download(string id)
        {
            try
            {
                var ds = kd.Sys_MauIns.Where(d => d.MaMauIn == id).FirstOrDefault();
                if (ds != null)
                {
                    FileInfo file = new FileInfo(HttpContext.Server.MapPath("~/FileUploads/MKinhDoanh/MauInHopDong/") + ds.TenFileDinhKemLuu);
                    if (file.Exists)
                        return new DownloadResult { VirtualPath = "~/FileUploads/MKinhDoanh/MauInHopDong/" + ds.TenFileDinhKemLuu, FileDownloadName = ds.TenFileDinhKemLuu };
                    else
                    {
                        ViewData["Message"] = "File not found !";
                        return View("Error");
                    }
                }
                else
                {
                    ViewData["Message"] = "File not found !";
                    return View("Error");
                }
            }
            catch (Exception e)
            {
                ViewData["Message"] = e.Message;
                return View("Error");
            }

        }

        [HttpPost]
        public ActionResult DeleteFile(string id)
        {
            try
            {
                var getTT = kd.Sys_MauIns.Where(d => d.MaMauIn.Equals(id)).FirstOrDefault();
                FileInfo file = new FileInfo(HttpContext.Server.MapPath("/FileUploads/MKinhDoanh/MauInHopDong/") + getTT.TenFileDinhKemLuu);
                getTT.TenFileDinhKem = null;
                getTT.TenFileDinhKemLuu = null;
                getTT.NoiDung = null;
                if (file.Exists)
                {
                    file.Delete();
                }
                kd.SubmitChanges();
                return RedirectToAction("Edit", new { id = id });
            }
            catch
            {
                return View();
            }
        }

        #endregion

        //GetMax số hợp đồng bảo vệ
        private string GetMaxMauIn()
        {
            try
            {
                return kd.Sys_MauIns.Max(d => d.MaMauIn);
            }
            catch
            {
                return string.Empty;
            }
        }


        //Mã tăng tự động
        public string IdGenerator()
        {

            sb = new StringBuilder();
            var date = DateTime.Now;
            string lastID = kd.Sys_MauIns.OrderByDescending(d => d.MaMauIn).Select(d => d.MaMauIn).FirstOrDefault();
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
                return "MINVB" + nam + thang + "-001";
            }
            else
            {
                int? idSuffix = Convert.ToInt32(lastID.Substring(lastID.Length - 3)) + 1;
                if (idSuffix <= 0 || idSuffix == null)
                    return "MINVB" + nam + thang + "-001";
                else
                {
                    sb.Append(idSuffix.ToString());
                    while (sb.Length < 3)
                    {
                        sb.Insert(0, "0");
                    }
                    return "MINVB" + nam + thang + "-" + sb.ToString();
                }
            }
        }

        public void ThongTinLoaiVanBan()
        {
            mauIn.TenLoaiVanBan = kd.tbl_DM_LoaiMauIns.Where(s => s.id == mauIn.MaLoaiVanBan).Select(s => s.tenLoaiMauIn).FirstOrDefault();
        }

        //public void BindDataDuAn(string maDuAn)
        //{
        //    List<tbl_DA_DuAn> nganHang = kd.tbl_DA_DuAns.ToList();
        //    nganHang.Insert(0, new tbl_DA_DuAn { maDuAn = string.Empty, tenDuAn = "-- Chọn dự án --" });
        //    ViewBag.DuAn = new SelectList(nganHang, "maDuAn", "tenDuAn", maDuAn);
        //}

        [HttpPost]
        public ActionResult LoadMore(int offset, int fetchNext, string tuNgay, string denNgay, string qSearch, int? trangThaiDuyet)
        {
            try
            {
                DateTime? fromDate = String.IsNullOrEmpty(tuNgay) ? (DateTime?)null : DateTime.ParseExact(tuNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                DateTime? toDate = String.IsNullOrEmpty(denNgay) ? (DateTime?)null : DateTime.ParseExact(denNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                mauIns = kd.sp_KD_MauIn_Index(qSearch, fromDate, toDate, trangThaiDuyet).Skip(offset).Take(fetchNext).ToList();
                ViewBag.OffSet = offset;
                ViewBag.MaNguoiLap = GetUser().manv;
                return PartialView("ViewIndex", mauIns);
            }
            catch
            {
                return View();
            }
        }

        public void getLoaiMauInDropDownlist()
        {
            var loaiMauIns = kd.tbl_DM_LoaiMauIns;
            ViewBag.LoaiMauIns = new SelectList(loaiMauIns, "id", "tenLoaiMauIn", mauIn.MaLoaiVanBan);
        }
    }
}
