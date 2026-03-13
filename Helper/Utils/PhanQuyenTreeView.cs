using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using BatDongSan.Models.HeThong;

namespace BatDongSan.Helper.Utils
{
    public static class PhanQuyenTreeView
    {

        private static StringBuilder build;
        private static IList<sp_Sys_CongViecCuaUsersResult> congViecs = null;
        public static StringBuilder BuildTreeView(IList<sp_Sys_CongViecCuaUsersResult> congViecOfUsers, string parentID)
        {
            congViecs = congViecOfUsers;
            build = new StringBuilder();
            build.Append("<ul>");
            // lay con cap 1
            var lst = congViecOfUsers.Where(d => d.maCha != null && d.maCha.Equals(parentID));
            int i = 1;
            foreach (var item in lst.ToList())
            {
                var child = congViecs.Where(d => d.maCha != null && d.maCha.Equals(item.maCongViec));
                if (child.Count() > 0)
                {
                    build.Append("<li><input type='checkbox' id=" + item.maCongViec + "_" + i.ToString() + " /><label for=" + item.maCongViec + "_" + i.ToString() + "><a class='viewPrivillege' href='#' data-value =" + item.maCongViec + ">" + item.tenCongViec + "</a></label>");
                    GetChildRecursive(child.ToList());
                }
                else
                {
                    build.Append("<li><label><a class='viewPrivillege' href='#' data-value =" + item.maCongViec + ">" + item.tenCongViec + "</a></label>");
                }
                build.Append("</li>");
                i++;
            }
            build.Append("</ul>");
            return build;
        }

        // lay con cap 2...n
        private static void GetChildRecursive(List<sp_Sys_CongViecCuaUsersResult> listChildren)
        {
            int j = 1;
            build.Append("<ul>");
            foreach (var item in listChildren)
            {
                var child = congViecs.Where(d => d.maCha != null && d.maCha == item.maCongViec);
                if (child.Count() > 0)
                {
                    build.Append("<li><input type='checkbox' id=" + item.maCongViec + "_" + j.ToString() + " /><label for=" + item.maCongViec + "_" + j.ToString() + "><a class='viewPrivillege' href='#' data-value =" + item.maCongViec + ">" + item.tenCongViec + "</a></label>");
                    GetChildRecursive(child.ToList());
                }
                else
                {
                    build.Append("<li><label><a class='viewPrivillege' href='#' data-value =" + item.maCongViec + ">" + item.tenCongViec + "</a></label>");
                }
                build.Append("</li>");
                j++;

            }
            build.Append("</ul>");
        }
    }
}