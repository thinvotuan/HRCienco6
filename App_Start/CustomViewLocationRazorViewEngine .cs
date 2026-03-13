using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace BatDongSan.App_Start
{
    public class CustomViewLocationRazorViewEngine : RazorViewEngine
    {
        public CustomViewLocationRazorViewEngine()
        {
            AreaViewLocationFormats = new[]
            {
                "~/Areas/{2}/Views/{1}/{0}.cshtml",
                "~/Areas/{2}/Views/Shared/{0}.cshtml"
            };
            AreaMasterLocationFormats = new[]
            {
                "~/Areas/{2}/Views/{1}/{0}.cshtml",
                "~/Areas/{2}/Views/Shared/{0}.cshtml"
            };
            AreaPartialViewLocationFormats = new[]
            {
                "~/Areas/{2}/Views/{1}/{0}.cshtml",
                "~/Areas/{2}/Views/Shared/{0}.cshtml"
            };
            ViewLocationFormats = new[]
            {
                "~/Views/{1}/{0}.cshtml",
                "~/Views/Shared/{0}.cshtml",
                "~/Views/HeThong/{1}/{0}.cshtml",
                "~/Views/BaoCao/{1}/{0}.cshtml",
                "~/Views/DanhMuc/{1}/{0}.cshtml",
                "~/Views/KhachHang/{1}/{0}.cshtml",  
                "~/Views/TinhCong/{1}/{0}.cshtml",
                "~/Views/NhanSu/{1}/{0}.cshtml",
                "~/Views/PhieuDeNghi/{1}/{0}.cshtml",
                 "~/Views/QLSuatAn/{1}/{0}.cshtml",
            };
            MasterLocationFormats = new[]
            {
                "~/Views/{1}/{0}.cshtml",  
                "~/Views/Shared/{0}.cshtml",
            };
            PartialViewLocationFormats = new[]
            {
                "~/Views/{1}/{0}.cshtml",
                "~/Views/Shared/{0}.cshtml",
                "~/Views/HeThong/{1}/{0}.cshtml",
                "~/Views/BaoCao/{1}/{0}.cshtml",
                "~/Views/TinhCong/{1}/{0}.cshtml",  
                "~/Views/KhachHang/{1}/{0}.cshtml",
                "~/Views/DanhMuc/{1}/{0}.cshtml",
                "~/Views/NhanSu/{1}/{0}.cshtml",
                "~/Views/PhieuDeNghi/{1}/{0}.cshtml",
                 "~/Views/QLSuatAn/{1}/{0}.cshtml",
            };
        }
    }
}