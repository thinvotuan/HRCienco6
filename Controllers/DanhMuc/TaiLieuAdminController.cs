using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper;
using BatDongSan.Helper.Common;
using BatDongSan.Models.DanhMuc;

namespace BatDongSan.Controllers.DanhMuc
{
    public class TaiLieuAdminController : ApplicationController
    {
        private LinqDanhMucDataContext context;
        private IList<tbl_DM_TaiLieu> taiLieus;
        private tbl_DM_TaiLieu taiLieu;
        private readonly string MCV = "TaiLieuAdmin";
        private bool? permission;

        public ActionResult Index()
        {
            #region Role user
            permission = GetPermission(MCV, BangPhanQuyen.QuyenXem);
            if (!permission.HasValue)
                return View("LogIn");
            if (!permission.Value)
                return View("AccessDenied");
            #endregion

            taiLieu = new tbl_DM_TaiLieu();
            taiLieu.tenNguoiUpload = (string)Session["TenNhanVien"];
            taiLieu.ngayLap = DateTime.Now;
            ViewBag.TaiLieu = taiLieu;
            using (context = new LinqDanhMucDataContext())
            {
                taiLieus = context.tbl_DM_TaiLieus.ToList();
            }
            return View(taiLieus);
        }

        //
        // GET: /TaiLieu/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /TaiLieu/Create

        public ActionResult Create()
        {
            return View();
        }

        //
        // POST: /TaiLieu/Create

        [HttpPost]
        public ActionResult Create(FormCollection collection, IEnumerable<HttpPostedFileBase> files)
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
                using (context = new LinqDanhMucDataContext())
                {
                    string[] nameacceptable = collection.GetValues("nameaccept");
                    string[] thumbnails = collection.GetValues("thumbnail");
                    taiLieus = new List<tbl_DM_TaiLieu>();
                    foreach (var file in files)
                    {
                        if (file != null)
                        {
                            tbl_DM_TaiLieu taiLieu = new tbl_DM_TaiLieu();
                            var date = DateTime.Now.ToString("yyyyMMdd-HHMMss");
                            string filePath = "/UploadFiles/TaiLieu/";

                            //Tạo tên mới cho file upload
                            string Generatedname = date.ToString() + file.FileName;
                            Directory.CreateDirectory(filePath);
                            var filePathOriginal = Server.MapPath(filePath);
                            if (nameacceptable.Contains(file.FileName))
                            {
                                string savedFileName = Path.Combine(filePathOriginal, Generatedname);
                                taiLieu.linkTaiLieu = filePath + Generatedname;
                                taiLieu.maNguoiUpload = GetUser().manv;
                                taiLieu.ngayLap = DateTime.Now;
                                taiLieu.tenNguoiUpload = (string)Session["TenNhanVien"];
                                taiLieu.tenTaiLieu = file.FileName;
                                taiLieu.tenTaiLieuMoi = Generatedname;
                                if (Array.Exists(thumbnails, s => s.Contains(file.FileName)) && file.ContentType.Contains("image") == false)
                                {
                                    string thumbnail = thumbnails.Where(s => s.Contains(file.FileName)).First();
                                    //Tách giá trị value thumbnail thành 2 phần 1-Tên file, 2-Tên đường dẫn thumbnail
                                    taiLieu.thumbnailURL = Regex.Split(thumbnail, "-SplitPoint-").Last();
                                }
                                else
                                {
                                    if (file.ContentType.Contains("image"))
                                    {
                                        taiLieu.thumbnailURL = "/UploadFiles/TaiLieu/" + Generatedname;
                                    }
                                }
                                file.SaveAs(savedFileName);
                                taiLieus.Add(taiLieu);
                                int Index = Array.IndexOf(nameacceptable, file.FileName);
                                Array.Clear(nameacceptable, Index, 1);
                            }
                        }
                    }
                    context.tbl_DM_TaiLieus.InsertAllOnSubmit(taiLieus);
                    context.SubmitChanges();
                }
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /TaiLieu/Edit/5

        public ActionResult Edit(int id)
        {
            return View();
        }

        public ActionResult Download(int id)
        {
            using (context = new LinqDanhMucDataContext())
            {
                taiLieu = context.tbl_DM_TaiLieus.Where(s => s.id == id).FirstOrDefault();
                string savedFileName = Path.Combine("/UploadFiles/TaiLieu/", taiLieu.tenTaiLieuMoi);
                return new DownloadResult { VirtualPath = savedFileName, FileDownloadName = taiLieu.tenTaiLieu };                
            }
        }

        //
        // POST: /TaiLieu/Edit/5

        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // POST: /TaiLieu/Delete/5

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
                using (context = new LinqDanhMucDataContext())
                {
                    taiLieu = context.tbl_DM_TaiLieus.Where(s => s.id == id).FirstOrDefault();
                    context.tbl_DM_TaiLieus.DeleteOnSubmit(taiLieu);
                    context.SubmitChanges();
                    System.IO.File.Delete(Server.MapPath(taiLieu.linkTaiLieu));
                }
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
    }
}
