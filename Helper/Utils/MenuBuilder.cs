using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using BatDongSan.Models.HeThong;

namespace BatDongSan.Utils
{

    public static class MenuHelper
    {
        private static StringBuilder build;
        private static IList<sp_Sys_CongViecCuaUsersResult> congViecs = null;

        public static StringBuilder BuildMenu(IList<sp_Sys_CongViecCuaUsersResult> congViecOfUsers, string parentID)
        {
            congViecs = congViecOfUsers;
            build = new StringBuilder();
            build.Append("<ul>");
            // lay con cap 1
            var lst = congViecOfUsers.Where(d => d.maCha != null && d.maCha.Equals(parentID));
            foreach (var item in lst.ToList())
            {
                var child = congViecs.Where(d => d.maCha != null && d.maCha.Equals(item.maCongViec));
                if (child.Count() > 0)
                {
                    build.Append("<li class='dropdown-submenu'><a href=\"#\"><i class='" + item.maIcon + "'></i>" + item.tenCongViec + "</a>");
                    GetChildRecursive(child.ToList());
                }
                else
                {
                    build.Append("<li><a href=/" + item.tenController + "/" + item.tenAction + "><i class='" + item.maIcon + "'></i>" + item.tenCongViec + "</a>");
                }
                build.Append("</li>");
            }
            build.Append("</ul>");
            return build;
        }

        // lay con cap 2...n
        private static void GetChildRecursive(List<sp_Sys_CongViecCuaUsersResult> listChildren)
        {
            build.Append("<ul class='dropdown-menu'>");
            foreach (var item in listChildren)
            {
                var child = congViecs.Where(d => d.maCha != null && d.maCha == item.maCongViec);
                if (child.Count() > 0)
                {
                    build.Append("<li><a href=\"#\">" + item.tenCongViec + "</a>");
                    GetChildRecursive(child.ToList());
                }
                else
                {
                    build.Append("<li><a href=/" + item.tenController + "/" + item.tenAction + "><i class='" + item.maIcon + "'></i>" + item.tenCongViec + "</a>");
                }
                build.Append("</li>");

            }
            build.Append("</ul>");
        }
        public static StringBuilder BuildMenuMobile(IList<sp_Sys_CongViecCuaUsersResult> congViecOfUsers, string parentID)
        {
            congViecs = congViecOfUsers;
            build = new StringBuilder();
            build.Append("<ul>");
            // lay con cap 1
            var lst = congViecOfUsers.Where(d => d.maCha != null && d.maCha.Equals(parentID));
            foreach (var item in lst.ToList())
            {
                var child = congViecs.Where(d => d.maCha != null && d.maCha.Equals(item.maCongViec));
                if (child.Count() > 0)
                {
                    build.Append("<li><a href=\"#\"><i class='" + item.maIcon + "'></i>" + item.tenCongViec + "</a>");
                    GetChildRecursiveMobile(child.ToList());
                }
                else
                {
                    build.Append("<li class='" + item.maCongViec + "'><a href=/" + item.tenController + "/" + item.tenAction + " class=\"SaveItemActive\"><i class='" + item.maIcon + "'></i>" + item.tenCongViec + "</a>");
                }
                build.Append("</li>");
            }
            build.Append("</ul>");
            return build;
        }

        // lay con cap 2...n
        private static void GetChildRecursiveMobile(List<sp_Sys_CongViecCuaUsersResult> listChildren)
        {
            build.Append("<ul>");
            foreach (var item in listChildren)
            {

                var child = congViecs.Where(d => d.maCha != null && d.maCha == item.maCongViec);
                if (child.Count() > 0)
                {
                    build.Append("<li  class='" + item.maCongViec + "'><a href=\"#\"><i class='" + item.maIcon + "'></i>" + item.tenCongViec + "</a>");
                    GetChildRecursive(child.ToList());
                }
                else
                {
                    build.Append("<li class='" + item.maCongViec + "'><a href=/" + item.tenController + "/" + item.tenAction + " class=\"SaveItemActive\"><i class='" + item.maIcon + "'></i>" + item.tenCongViec + "</a>");
                }
                build.Append("</li>");

            }
            build.Append("</ul>");
        }
    }
}