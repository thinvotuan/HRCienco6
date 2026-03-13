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
    public class QuiTrinhDuyetController : ApplicationController
    {
        LinqDanhMucDataContext linqDM = new LinqDanhMucDataContext();
        LinqHeThongDataContext db = new LinqHeThongDataContext();
        tbl_HT_QuiTrinhDuyet quiTrinhDuyet;
        readonly string MCV = "QuiTrinhDuyet";
        bool? permission;
        IList<sp_QuiTrinhDuyet_IndexResult> quiTrinhDuyets;
        public ActionResult Index(int? page, int? pageSize, string searchString)
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
            quiTrinhDuyets = db.sp_QuiTrinhDuyet_Index(searchString, currentPageIndex, pageSize).ToList();
            try
            {
                ViewBag.Count = quiTrinhDuyets[0].tongDong;
                tongSoDong = quiTrinhDuyets[0].tongDong;
            }
            catch
            {
                ViewBag.Count = 0;
            }
            TempData["Params"] = searchString + ",";
            if (Request.IsAjaxRequest())
            {
                return PartialView("ViewIndex", quiTrinhDuyets.ToPagedList(currentPageIndex, 30, true, tongSoDong));
            }
            return View(quiTrinhDuyets.ToPagedList(currentPageIndex, 30, true, tongSoDong));
        }

        //
        // GET: /LoaiQuanHeGiaDinh/Details/5

        public ActionResult Details(int id)
        {
            return View();
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

            tbl_HT_QuiTrinhDuyet quiTrinhDuyet = new tbl_HT_QuiTrinhDuyet();
            quiTrinhDuyet.NgayLap = DateTime.Now;
            quiTrinhDuyet.MaNV = GetUser().manv;
            ViewBag.TenNhanVien = (string)Session["TenNhanVien"];
            BindDataCongViecCanDuyet(string.Empty);
            BindDataTrangThaiDongMo(string.Empty);
            return View(quiTrinhDuyet);
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
                return RedirectToAction("Edit/" + db.tbl_HT_QuiTrinhDuyets.Max(d => d.Id).ToString());
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
                GetData(Convert.ToInt32(id));
                return View(quiTrinhDuyet);
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
        public ActionResult Edit(int id, FormCollection collection)
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
                return RedirectToAction("Edit/" + collection.Get("Id"));
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
                var delete = db.tbl_HT_QuiTrinhDuyets.Where(d => d.Id == Convert.ToInt32(id)).FirstOrDefault();
                delete.TrangThaiDong = 0;
                db.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        #region Bindata
        /// <summary>
        /// Công việc cần duyệt
        /// </summary>
        /// <param name="maCongViec"></param>
        public void BindDataCongViecCanDuyet(string maCongViec)
        {
            List<tbl_HT_CongViecCanDuyet> phDauThau = db.tbl_HT_CongViecCanDuyets.ToList();
            phDauThau.Insert(0, new tbl_HT_CongViecCanDuyet { maCongViec = string.Empty, tenCongViec = "-- Chọn công việc--" });
            ViewBag.CongViecCanDuyet = new SelectList(phDauThau, "maCongViec", "tenCongViec", maCongViec);
        }

        public void BindDataTrangThaiDongMo(string trangThai)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("1", "Mở");
            dic.Add("0", "Đóng");
            ViewBag.TrangThaiDM = dic;
        }
        #endregion

        /// <summary>
        /// Insert vào qui trình duyệt phiếu
        /// </summary>
        /// <param name="col"></param>
        /// <param name="isCreate"></param>
        public void BindDataToSave(FormCollection col, bool isCreate)
        {
            //Insert qui trình duyệt         
            int idQuiTrinh = 0;
            if (isCreate == true)
            {
                quiTrinhDuyet = new tbl_HT_QuiTrinhDuyet();
                quiTrinhDuyet.MaCongViec = col.Get("MaCongViec");
                quiTrinhDuyet.TenQuiTrinh = col.Get("TenQuiTrinh");
                quiTrinhDuyet.GhiChu = col.Get("GhiChu");
                quiTrinhDuyet.NgayLap = DateTime.ParseExact(col.Get("NgayLap"), "dd/MM/yyyy", null);
                quiTrinhDuyet.TrangThaiDong = Convert.ToInt32(col.Get("TrangThaiDong"));
                quiTrinhDuyet.MaNV = GetUser().manv;
                db.tbl_HT_QuiTrinhDuyets.InsertOnSubmit(quiTrinhDuyet);
                db.SubmitChanges();
                idQuiTrinh = Convert.ToInt32(db.tbl_HT_QuiTrinhDuyets.Max(d => d.Id));
            }
            else
            {
                quiTrinhDuyet = db.tbl_HT_QuiTrinhDuyets.Where(d => d.Id == Convert.ToInt32(col.Get("Id"))).FirstOrDefault();
                quiTrinhDuyet.GhiChu = col.Get("GhiChu");
                quiTrinhDuyet.TrangThaiDong = Convert.ToInt32(col.Get("TrangThaiDong"));
                idQuiTrinh = quiTrinhDuyet.Id;
            }
            //Inser qui trình duyệt chi tiết
            var delQuiTrinhDuyetChiTiet = db.tbl_HT_QuiTrinhDuyetChiTiets.Where(d => d.IdQuiTrinh == idQuiTrinh);
            db.tbl_HT_QuiTrinhDuyetChiTiets.DeleteAllOnSubmit(delQuiTrinhDuyetChiTiet);
            List<tbl_HT_QuiTrinhDuyetChiTiet> ct = new List<tbl_HT_QuiTrinhDuyetChiTiet>();
            tbl_HT_QuiTrinhDuyetChiTiet quiTrinhDuyetChiTiet;
            string[] maQuiTrinh = col.GetValues("maQuiTrinh");
            if (maQuiTrinh != null)
            {
                for (int i = 0; i < maQuiTrinh.Count(); i++)
                {
                    quiTrinhDuyetChiTiet = new tbl_HT_QuiTrinhDuyetChiTiet();
                    quiTrinhDuyetChiTiet.MaNhanVien = col.GetValues("maNhanVienDuyet")[i];
                    quiTrinhDuyetChiTiet.MaCapDuyet = col.GetValues("maQuiTrinh")[i];
                    quiTrinhDuyetChiTiet.ThuTuDuyet = Convert.ToInt32(col.GetValues("sttQuiTrinhDuyet")[i]);
                    quiTrinhDuyetChiTiet.IdQuiTrinh = idQuiTrinh;
                    ct.Add(quiTrinhDuyetChiTiet);
                }
                if (ct != null && ct.Count > 0)
                {
                    db.tbl_HT_QuiTrinhDuyetChiTiets.InsertAllOnSubmit(ct);
                    db.SubmitChanges();
                }
            }
        }

        /// <summary>
        /// Thông tin dữ liệu qui trình duyệt
        /// </summary>
        /// <param name="id"></param>
        public void GetData(int id)
        {
            quiTrinhDuyet = db.tbl_HT_QuiTrinhDuyets.Where(d => d.Id == id).FirstOrDefault();
            BindDataCongViecCanDuyet(quiTrinhDuyet.MaCongViec);
            BindDataTrangThaiDongMo(string.Empty);
            ViewBag.TrangThai = quiTrinhDuyet.TrangThaiDong;
            ViewBag.TenNhanVien = (db.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == quiTrinhDuyet.MaNV).Select(d => d.ho).FirstOrDefault() ?? string.Empty) + " " + (db.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == quiTrinhDuyet.MaNV).Select(d => d.ten).FirstOrDefault() ?? string.Empty);

            //Chi tiết qui trình duyệt
            ViewBag.ChiTietQuiTrinhDuyet = db.tbl_HT_QuiTrinhDuyetChiTiets.Where(d => d.IdQuiTrinh == id).Select(d => new ChiTietQuiTrinhDuyet
            {
                TenBuocDuyet = db.tbl_HT_QuiTrinhDuyet_BuocDuyets.Where(c => c.maBuocDuyet == d.MaCapDuyet).Select(c => c.tenBuocDuyet).FirstOrDefault() ?? string.Empty,
                MaBuocDuyet = d.MaCapDuyet,
                ThuTuDuyet = d.ThuTuDuyet ?? 0,
                MaNhanVien = d.MaNhanVien,
                TenNhanVien = (db.GetTable<tbl_NS_NhanVien>().Where(c => c.maNhanVien == d.MaNhanVien).Select(c => c.ho).FirstOrDefault() ?? string.Empty) + " " + (db.GetTable<tbl_NS_NhanVien>().Where(c => c.maNhanVien == d.MaNhanVien).Select(c => c.ten).FirstOrDefault() ?? string.Empty),
            }).ToList();
        }

        /// <summary>
        /// Chọn các bước duyệt
        /// </summary>
        /// <param name="stt"></param>
        /// <returns></returns>
        public ActionResult ChonQuiTrinh(int stt)
        {
            ViewData["stt"] = stt;
            ViewBag.BuocDuyet = db.tbl_HT_QuiTrinhDuyet_BuocDuyets.ToList();
            return View();
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
    }
}
