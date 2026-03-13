using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Models.DanhMuc;

namespace BatDongSan.Controllers.DanhMuc
{
    public class TaiLieuController : ApplicationController
    {
        private LinqDanhMucDataContext context;
        private IList<tbl_DM_TaiLieu> taiLieus;
        private tbl_DM_TaiLieu taiLieu;

        public ActionResult Index()
        {            
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
        public ActionResult Create(FormCollection collection)
        {
            try
            {                
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
        // GET: /TaiLieu/Delete/5

        public ActionResult Delete(int id)
        {
            return View();
        }

        //
        // POST: /TaiLieu/Delete/5

        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
    }
}
