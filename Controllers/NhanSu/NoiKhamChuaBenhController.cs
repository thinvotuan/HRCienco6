using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Models.NhanSu;

namespace BatDongSan.Controllers.NhanSu
{
    public class NoiKhamChuaBenhController : ApplicationController
    {
        private LinqNhanSuDataContext context;
        private IList<NoiKhamChuaBenhModel> noiKhams;
        private tbl_DM_NoiKhamChuaBenh noiKham;        

        public ActionResult Index()
        {
            context = new LinqNhanSuDataContext();
            noiKhams = (from n in context.tbl_DM_NoiKhamChuaBenhs
                       join t in context.Sys_TinhThanhs on n.maTinhThanh equals t.maTinhThanh
                       join l in context.tbl_DM_LoaiBenhViens on n.loai equals l.id into g
                       from a in g.DefaultIfEmpty()
                       orderby n.maNoiKham
                       select new NoiKhamChuaBenhModel{
                           diaChi = n.diaChi,
                           ghiChu = n.ghiChu,
                           loai = n.loai,
                           maNoiKham = n.maNoiKham,
                           maTinhThanh = n.maTinhThanh,
                           ngayCapNhat = n.ngayCapNhat,
                           nguoiLap = n.nguoiLap,
                           tenLoaiBenhVien = a.tenLoaiBenhVien,
                           tenTinhThanh = t.tenTinhThanh,
                           tenNoiKham = n.tenNoiKham
                       }).ToList();
            return View(noiKhams);
        }

        //
        // GET: /NoiKhamChuaBenh/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /NoiKhamChuaBenh/Create

        public ActionResult Create()
        {
            noiKham = new tbl_DM_NoiKhamChuaBenh();
            context = new LinqNhanSuDataContext();
            GetAllDropdownList();
            return PartialView("Create", noiKham);
        }

        //
        // POST: /NoiKhamChuaBenh/Create

        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {
                context = new LinqNhanSuDataContext();
                noiKham = new tbl_DM_NoiKhamChuaBenh();
                noiKham.diaChi = collection["diaChi"];
                noiKham.ghiChu = collection["ghiChu"];
                noiKham.loai = Convert.ToInt32(collection["loai"]);
                noiKham.maNoiKham = collection["maNoiKham"];
                noiKham.maTinhThanh = collection["maTinhThanh"];
                noiKham.ngayCapNhat = DateTime.Now;
                noiKham.nguoiLap = GetUser().manv;
                noiKham.tenNoiKham = collection["tenNoiKham"];
                var obj = context.tbl_DM_NoiKhamChuaBenhs.Where(s => s.maNoiKham == noiKham.maNoiKham).FirstOrDefault();
                if (obj != null)
                {
                    TempData["NoiKham"]= "Mã bệnh viện đã tồn tại";
                    var loaiBenhViens = context.tbl_DM_LoaiBenhViens.ToList();
                    loaiBenhViens.Insert(0, new tbl_DM_LoaiBenhVien { id = 0, tenLoaiBenhVien = "[Chọn Loại bệnh viện]" });
                    ViewBag.LoaiBenhViens = new SelectList(loaiBenhViens, "id", "tenLoaiBenhVien", noiKham.loai);
                    return View(noiKham);
                }
                context.tbl_DM_NoiKhamChuaBenhs.InsertOnSubmit(noiKham);
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /NoiKhamChuaBenh/Edit/5

        public ActionResult Edit(string id)
        {
            context = new LinqNhanSuDataContext();
            noiKham = context.tbl_DM_NoiKhamChuaBenhs.Where(s => s.maNoiKham == id).FirstOrDefault();
            GetAllDropdownList();
            return PartialView("Edit", noiKham);
        }

        //
        // POST: /NoiKhamChuaBenh/Edit/5

        [HttpPost]
        public ActionResult Edit(string id, FormCollection collection)
        {
            try
            {
                context = new LinqNhanSuDataContext();
                noiKham = context.tbl_DM_NoiKhamChuaBenhs.Where(s => s.maNoiKham == id).FirstOrDefault();
                noiKham.diaChi = collection["diaChi"];
                noiKham.ghiChu = collection["ghiChu"];
                noiKham.loai = Convert.ToInt32(collection["loai"]);                
                noiKham.maTinhThanh = collection["maTinhThanh"];
                noiKham.ngayCapNhat = DateTime.Now;
                noiKham.nguoiLap = GetUser().manv;
                noiKham.tenNoiKham = collection["tenNoiKham"];                
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // POST: /NoiKhamChuaBenh/Delete/5

        [HttpPost]
        public ActionResult Delete(string id)
        {
            try
            {
                context = new LinqNhanSuDataContext();
                noiKham = context.tbl_DM_NoiKhamChuaBenhs.Where(s => s.maNoiKham == id).FirstOrDefault();
                context.tbl_DM_NoiKhamChuaBenhs.DeleteOnSubmit(noiKham);
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        public void GetAllDropdownList() {
            var tinhThanhs = context.Sys_TinhThanhs.ToList();
            tinhThanhs.Insert(0, new Sys_TinhThanh { maTinhThanh = "", tenTinhThanh = "[Chọn tỉnh thành]" });
            ViewBag.TinhThanhs = new SelectList(tinhThanhs, "maTinhThanh", "tenTinhThanh", noiKham.maTinhThanh);
            
            var loaiBenhViens = context.tbl_DM_LoaiBenhViens.ToList();
            loaiBenhViens.Insert(0, new tbl_DM_LoaiBenhVien { id = 0, tenLoaiBenhVien = "[Chọn Loại bệnh viện]" });
            ViewBag.LoaiBenhViens = new SelectList(loaiBenhViens, "id", "tenLoaiBenhVien", noiKham.loai);
        }
    }
}
