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
using BatDongSan.Models.NhanSu;
using BatDongSan.Utils.Paging;
namespace BatDongSan.Controllers.DanhMuc
{
    public class DanhSachCongNhanController : ApplicationController
    {
        private LinqDanhMucDataContext context = new LinqDanhMucDataContext();
        private LinqNhanSuDataContext contextNS = new LinqNhanSuDataContext();
        private IList<tbl_NS_DanhSachCongNhan> _DanhSachCongNhans;
        private IList<sp_NS_DanhSachCongNhan_IndexResult> danhSachCNs;
        private tbl_NS_DanhSachCongNhan _DanhSachCongNhan;
        
        
        private readonly string MCV = "DanhSachCongNhan";
        private bool? permission;
        //
        // GET: /DanhSachCongNhan/

        
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
            


            danhSachCNs = context.sp_NS_DanhSachCongNhan_Index(null,searchString, currentPageIndex, 30).OrderByDescending(d => d.id).ToList();

            int? tongSoDong = 0;
            try
            {
                ViewBag.Count = danhSachCNs[0].tongSoDong;
                tongSoDong = danhSachCNs[0].tongSoDong;
            }
            catch
            {
                ViewBag.Count = 0;
            }
            if (Request.IsAjaxRequest())
            {
                return PartialView("PartialIndex", danhSachCNs.ToPagedList(currentPageIndex, 30, true, tongSoDong));
            }
          
            ViewBag.searchString = searchString;
           
            return View(danhSachCNs.ToPagedList(currentPageIndex, 30, true, tongSoDong));
        }

        public void GetDataFromView(FormCollection collection)
        {
            _DanhSachCongNhan.CMNDNgayCap = String.IsNullOrEmpty(collection["CMNDNgayCap"]) ? (DateTime?)null : DateTime.ParseExact(collection["CMNDNgayCap"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
            _DanhSachCongNhan.CMNDNoiCap = collection["CMNDNoiCap"];
            _DanhSachCongNhan.CMNDSo = collection["CMNDSo"];
            _DanhSachCongNhan.email = collection["email"];
            _DanhSachCongNhan.gioiTinh = collection["gioiTinh"] == "True" ? true : false;
            _DanhSachCongNhan.ho = collection["ho"];
           
            _DanhSachCongNhan.ngaySinh = String.IsNullOrEmpty(collection["ngaySinh"]) ? (DateTime?)null : DateTime.ParseExact(collection["ngaySinh"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
            _DanhSachCongNhan.ngayVaoLam = String.IsNullOrEmpty(collection["ngayVaoLam"]) ? (DateTime?)null : DateTime.ParseExact(collection["ngayVaoLam"], "dd/MM/yyyy", CultureInfo.InvariantCulture); ;
             _DanhSachCongNhan.noiSinh = collection["noiSinh"];
             _DanhSachCongNhan.phoneNumber = collection["phoneNumber"];
             _DanhSachCongNhan.ten = collection["ten"];
            
             _DanhSachCongNhan.boPhan = collection["boPhan"];
             _DanhSachCongNhan.thongTinKhac = collection["thongTinKhac"];
             
             //nhanVien.tenAnhDaiDien = collection["tenAnhDaiDien"];
            //nhanVien.tenAnhDaiDienLuu = collection["tenAnhDaiDienLuu"];
            ////nhanVien.tenGoiKhac = collection["tenGoiKhac"];
         



        }
        public string IdGenerator()
        {

            StringBuilder sb = new StringBuilder();
            var date = DateTime.Now;
            string lastID = context.tbl_NS_DanhSachCongNhans.OrderByDescending(d => d.maCongNhan).Select(d => d.maCongNhan).FirstOrDefault();
           
            if (String.IsNullOrEmpty(lastID))
            {
                return "CN-" +  "0001";
            }
            else
            {
                int? idSuffix = Convert.ToInt32(lastID.Substring(lastID.Length - 4)) + 1;
                if (idSuffix <= 0 || idSuffix == null)
                    return "CN-" + "0001";
                else
                {
                    sb.Append(idSuffix.ToString());
                    while (sb.Length < 4)
                    {
                        sb.Insert(0, "0");
                    }
                    return "CN-"  + sb.ToString();
                }
            }
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
            _DanhSachCongNhan = new tbl_NS_DanhSachCongNhan();
            ViewBag.maCongNhan = IdGenerator();
            return PartialView("Create", _DanhSachCongNhan);
        }
        public int CheckMaCongNhan(string maCongNhan){
            var checkList = context.tbl_NS_DanhSachCongNhans.Where(d => d.maCongNhan == maCongNhan).FirstOrDefault();
            if (checkList != null) {
                return 1;
            }
            return 0;
        }
        //Check mã chấm công
        public int CheckMaChamCong(string maChamCong)
        {
            //BatDongSan.Models.ChamCong.LinqChamCongServerDataContext contextCC = new BatDongSan.Models.ChamCong.LinqChamCongServerDataContext();
            var checkList = context.tbl_NS_DanhSachCongNhans.Where(d => d.maChamCong == maChamCong).FirstOrDefault();
            if (checkList != null)
            {
                return 1;
            }
            var checkListNS = contextNS.tbl_NS_NhanViens.Where(d => d.maChamCong == maChamCong).FirstOrDefault();
            if (checkListNS != null)
            {
                return 1;
            }
            //var checkListNV = contextCC.WS_tblUserinfos.Where(d => d.badgenumber == maChamCong).FirstOrDefault();
            //if(checkListNV != null){
            // return 1;
            //}


            return 0;
        }
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
             // check ton tai cong nhan
                //BatDongSan.Models.ChamCong.LinqChamCongServerDataContext contextCC = new BatDongSan.Models.ChamCong.LinqChamCongServerDataContext();
                if (collection["maChamCong"] != null && collection["maChamCong"] != "")
                {
                    var checkList = context.tbl_NS_DanhSachCongNhans.Where(d => d.maCongNhan == collection["maCongNhan"] || d.maChamCong == collection["maChamCong"]).FirstOrDefault();
                    if (checkList != null)
                    {

                        return View("error");
                    }
                    var checkListNS = contextNS.tbl_NS_NhanViens.Where(d => d.maChamCong == collection["maChamCong"]).FirstOrDefault();
                    if (checkListNS != null)
                    {

                        return View("error");
                    }
                    

                    //var checkListNV = contextCC.WS_tblUserinfos.Where(d => d.badgenumber == collection["maChamCong"]).FirstOrDefault();
                    //if (checkListNV != null)
                    //{
                    //    return View("error");
                    //}
                }
                _DanhSachCongNhan = new tbl_NS_DanhSachCongNhan();
                GetDataFromView(collection);
                _DanhSachCongNhan.ngayLap = DateTime.Now;
                _DanhSachCongNhan.nguoiLap = GetUser().manv;
                _DanhSachCongNhan.maCongNhan = IdGenerator();
                _DanhSachCongNhan.maChamCong = collection["maChamCong"];
             

              
                context.tbl_NS_DanhSachCongNhans.InsertOnSubmit(_DanhSachCongNhan);
                context.SubmitChanges();
               // Insert vao tbl nhan vien
                //if (collection["maChamCong"] != null && collection["maChamCong"] != "")
                //{
                //    BatDongSan.Models.ChamCong.WS_tblUserinfo tblNhanVien = new BatDongSan.Models.ChamCong.WS_tblUserinfo();
                //    tblNhanVien.badgenumber = collection["maChamCong"];
                //    tblNhanVien.fullName = collection["ho"] + ' ' + collection["ten"];
                //    tblNhanVien.employeeId = collection["maCongNhan"];
                //    contextCC.WS_tblUserinfos.InsertOnSubmit(tblNhanVien);
                //    contextCC.SubmitChanges();
                //}
                SaveActiveHistory("Thêm công nhân: " + _DanhSachCongNhan.maCongNhan);
                
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
        
        public ActionResult Edit(int id)
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenSua);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            _DanhSachCongNhan = context.tbl_NS_DanhSachCongNhans.Where(s => s.id == id).FirstOrDefault();
            return PartialView("Edit", _DanhSachCongNhan);
        }

        //
        // POST: /ChucDanh/Edit/5

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

                _DanhSachCongNhan = context.tbl_NS_DanhSachCongNhans.Where(s => s.id == id).FirstOrDefault();
                _DanhSachCongNhan.ngayLap = _DanhSachCongNhan.ngayLap;
                _DanhSachCongNhan.nguoiLap =_DanhSachCongNhan.nguoiLap;

                //Lưu dữ liệu nhân viên
                GetDataFromView(collection);
                _DanhSachCongNhan.maCongNhan = _DanhSachCongNhan.maCongNhan;
                _DanhSachCongNhan.maChamCong = collection["maChamCong"];
                context.SubmitChanges();
                // Xoa nhan vien hien tai trong ban cham cong va cap nhat moi
                // Delete trong tblNhanVien DBChamCong
               

                    //BatDongSan.Models.ChamCong.LinqChamCongServerDataContext contextCC = new BatDongSan.Models.ChamCong.LinqChamCongServerDataContext();
                    //if (collection["maCCBanDau"] != null && collection["maCCBanDau"] != "")
                    //{
                    //    BatDongSan.Models.ChamCong.WS_tblUserinfo tblNhanVien = new BatDongSan.Models.ChamCong.WS_tblUserinfo();
                    //    tblNhanVien = contextCC.WS_tblUserinfos.Where(d => d.badgenumber == collection["maCCBanDau"]).FirstOrDefault();
                    //    if (tblNhanVien != null)
                    //    {
                    //        contextCC.WS_tblUserinfos.DeleteOnSubmit(tblNhanVien);
                    //        contextCC.SubmitChanges();
                    //    }
                    //}
                //    if (_DanhSachCongNhan.maChamCong != null && _DanhSachCongNhan.maChamCong != "")
                //    {
                //    // Add Moi tblNhanvien cham cong
                //    BatDongSan.Models.ChamCong.WS_tblUserinfo tblNhanVienMoi = new BatDongSan.Models.ChamCong.WS_tblUserinfo();
                //    tblNhanVienMoi.badgenumber = _DanhSachCongNhan.maChamCong;
                //    tblNhanVienMoi.fullName = _DanhSachCongNhan.ho + ' ' + _DanhSachCongNhan.ten;
                //    tblNhanVienMoi.employeeId = _DanhSachCongNhan.maCongNhan;
                //    contextCC.WS_tblUserinfos.InsertOnSubmit(tblNhanVienMoi);
                //    contextCC.SubmitChanges();
                //}
                SaveActiveHistory("Sửa công nhân: " + _DanhSachCongNhan.maCongNhan);
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }


        //
        // POST: /ChucDanh/Delete/5

        [HttpPost]
        public ActionResult Delete(int id)
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
                _DanhSachCongNhan = context.tbl_NS_DanhSachCongNhans.Where(s => s.id == id).FirstOrDefault();
                if (_DanhSachCongNhan != null)
                {
                    context.tbl_NS_DanhSachCongNhans.DeleteOnSubmit(_DanhSachCongNhan);
                    context.SubmitChanges();
                    // Delete trong tblNhanVien DBChamCong
                    //BatDongSan.Models.ChamCong.LinqChamCongServerDataContext contextCC = new BatDongSan.Models.ChamCong.LinqChamCongServerDataContext();
                    //if (_DanhSachCongNhan.maChamCong != null && _DanhSachCongNhan.maChamCong != "")
                    //{
                    //    BatDongSan.Models.ChamCong.WS_tblUserinfo tblNhanVien = new BatDongSan.Models.ChamCong.WS_tblUserinfo();
                    //    tblNhanVien = contextCC.WS_tblUserinfos.Where(d => d.badgenumber == _DanhSachCongNhan.maChamCong).FirstOrDefault();
                    //    if (tblNhanVien != null)
                    //    {
                    //        contextCC.WS_tblUserinfos.DeleteOnSubmit(tblNhanVien);
                    //        contextCC.SubmitChanges();
                    //    }
                    //}
                }
              
                SaveActiveHistory("Xóa công nhân: " + id);
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

    }
}
