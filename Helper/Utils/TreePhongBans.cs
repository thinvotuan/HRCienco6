using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using BatDongSan.Models.DanhMuc;

namespace BatDongSan.Helper.Utils
{
    public static class TreePhongBans
    {
        private static StringBuilder buildTree = null;
        private static IList<tbl_DM_PhongBan> _departments = null;

        #region Build tree department
        public static StringBuilder BuildTreeDepartment(IList<tbl_DM_PhongBan> departments)
        {
            _departments = departments;
            var parent = _departments.Where(d => string.IsNullOrEmpty(d.maCha)).ToList();
            buildTree = new StringBuilder();
            buildTree.Append("<ul class=\"navigation1\"  id='navigation1'>");
            // lay con cap 1                
            foreach (var item in parent)
            {
                //lay all con cua cha dau tien
                var child = _departments.Where(k => k.maCha == item.maPhongBan);

                if (child.Count() > 0)
                {
                    //buildTree.Append("<li><a style=\"font-weight:bold\" alt=\"" + item.departmentId + "\" href=\"javascript:void(0);\">" + item.departmentName + "</a>");
                    buildTree.Append("<li><a class=\"hvr-underline-from-center\" style=\"font-weight:bold\" flag=\"false\" alt=\"" + item.maPhongBan + "\" href=\"javascript:void(0);\">" + item.tenPhongBan + "</a>");

                    GetChildRecursive(child.ToList(), item.maPhongBan);
                }
                else
                {
                    buildTree.Append("<li><a class=\"hvr-underline-from-center\" style=\"font-weight:bold\" flag=\"true\" alt=\"" + item.maPhongBan + "\" href=\"javascript:void(0);\">" + item.tenPhongBan + "</a>");
                }

                buildTree.Append("</li>");
            }
            buildTree.Append("</ul>");
            return buildTree;
        }

        private static void GetChildRecursive(IList<tbl_DM_PhongBan> child, string p)
        {
            buildTree.Append("<ul>");
            foreach (var item in child)
            {
                var child1 = _departments.Where(d => d.maCha == item.maPhongBan);
                if (child1.Count() > 0)
                {
                    buildTree.Append("<li><a class=\"hvr-underline-from-center\" style=\"font-weight:bold\" href=\"javascript:void(0);\" flag=\"false\" alt=\"" + item.maPhongBan + "\">" + item.tenPhongBan + "</a>");
                    GetChildRecursive(child1.ToList(), item.maPhongBan);
                }
                else
                {
                    buildTree.Append("<li><a class=\"hvr-underline-from-center\" href=\"javascript:void(0);\" flag=\"true\" alt=\"" + item.maPhongBan + "\">" + item.tenPhongBan + "</a>");
                }

                buildTree.Append("</li>");

            }
            buildTree.Append("</ul>");
        }
        #endregion
    }


    public static class TreePhongBanAjax
    {
        private static StringBuilder buildTree = null;
        private static IList<tbl_DM_PhongBan> _departments = null;

        #region Build tree department
        public static StringBuilder BuildTreeDepartment(IList<tbl_DM_PhongBan> departments)
        {
            _departments = departments;
            var parent = _departments.Where(d => string.IsNullOrEmpty(d.maCha)).ToList();
            buildTree = new StringBuilder();
            buildTree.Append("<ul class='navigation1' id='navigation1' style='width: 100%; min-width: 500px; display: block'>");
            // lay con cap 1                
            foreach (var item in parent)
            {
                //lay all con cua cha dau tien
                var child = _departments.Where(k => k.maCha == item.maPhongBan);

                if (child.Count() > 0)
                {
                    //buildTree.Append("<li><a style=\"font-weight:bold\" alt=\"" + item.departmentId + "\" href=\"javascript:void(0);\">" + item.departmentName + "</a>");
                    buildTree.Append("<li role='presentation'><a  class=\"hvr-underline-from-center\"  data-ajax='true' data-ajax-begin='beginPaging' data-ajax-failure='failurePaging' data-ajax-mode='replace' data-ajax-success='successPaging' data-ajax-update='#content' style='font-weight:bold' id='" + item.maPhongBan + "' href='/NhanVien?maPhongBan=" + item.maPhongBan + "'>" + item.tenPhongBan + "</a>");

                    GetChildRecursive(child.ToList(), item.maPhongBan);
                }
                else
                {
                    buildTree.Append("<li role='presentation'><a  class=\"hvr-underline-from-center\" data-ajax='true' data-ajax-begin='beginPaging' data-ajax-failure='failurePaging' data-ajax-mode='replace' data-ajax-success='successPaging' data-ajax-update='#content' style='font-weight:bold' id='" + item.maPhongBan + "' href='/NhanVien?maPhongBan=" + item.maPhongBan + "'>" + item.tenPhongBan + "</a>");
                }

                buildTree.Append("</li>");
            }
            buildTree.Append("</ul>");
            return buildTree;
        }

        private static void GetChildRecursive(IList<tbl_DM_PhongBan> child, string p)
        {
            buildTree.Append("<ul>");
            foreach (var item in child)
            {
                var child1 = _departments.Where(d => d.maCha == item.maPhongBan);
                if (child1.Count() > 0)
                {
                    buildTree.Append("<li role='presentation'><a  class=\"hvr-underline-from-center\" data-ajax='true' data-ajax-begin='beginPaging' data-ajax-failure='failurePaging' data-ajax-mode='replace' data-ajax-success='successPaging' data-ajax-update='#content' style='font-weight:bold' id='" + item.maPhongBan + "' href='/NhanVien?maPhongBan=" + item.maPhongBan + "'>" + item.tenPhongBan + "</a>");
                    GetChildRecursive(child1.ToList(), item.maPhongBan);
                }
                else
                {
                    buildTree.Append("<li role='presentation'><a  class=\"hvr-underline-from-center\" data-ajax='true' data-ajax-begin='beginPaging' data-ajax-failure='failurePaging' data-ajax-mode='replace' data-ajax-success='successPaging' data-ajax-update='#content' id='" + item.maPhongBan + "' href='/NhanVien?maPhongBan=" + item.maPhongBan + "'>" + item.tenPhongBan + "</a>");
                }

                buildTree.Append("</li>");

            }
            buildTree.Append("</ul>");
        }
        #endregion
    }

    public static class TreePhongBanStyle
    {
        private static StringBuilder buildTree = null;
        private static IList<tbl_DM_PhongBan> _departments = null;

        #region Build tree department
        public static StringBuilder BuildTreeDepartment(IList<tbl_DM_PhongBan> departments)
        {
            _departments = departments;
            var parent = _departments.Where(d => string.IsNullOrEmpty(d.maCha)).ToList();
            buildTree = new StringBuilder();
            buildTree.Append("<ul class='navigation1' id='navigation1' style='width: 100%; min-width: 500px; display: block'>");
            // lay con cap 1                
            foreach (var item in parent)
            {
                //lay all con cua cha dau tien
                var child = _departments.Where(k => k.maCha == item.maPhongBan);

                if (child.Count() > 0)
                {
                    //buildTree.Append("<li><a style=\"font-weight:bold\" alt=\"" + item.departmentId + "\" href=\"javascript:void(0);\">" + item.departmentName + "</a>");
                    buildTree.Append("<li role='presentation'><a  class=\"hvr-underline-from-center\" style='font-weight:bold' id='" + item.maPhongBan + "' href='javascript:void(0)'>" + item.tenPhongBan + "</a>");

                    GetChildRecursive(child.ToList(), item.maPhongBan);
                }
                else
                {
                    buildTree.Append("<li role='presentation'><a  class=\"hvr-underline-from-center\" style='font-weight:bold' id='" + item.maPhongBan + "' href='javascript:void(0)'>" + item.tenPhongBan + "</a>");
                }

                buildTree.Append("</li>");
            }
            buildTree.Append("</ul>");
            return buildTree;
        }

        private static void GetChildRecursive(IList<tbl_DM_PhongBan> child, string p)
        {
            buildTree.Append("<ul>");
            foreach (var item in child)
            {
                var child1 = _departments.Where(d => d.maCha == item.maPhongBan);
                if (child1.Count() > 0)
                {
                    buildTree.Append("<li role='presentation'><a  class=\"hvr-underline-from-center\" style='font-weight:bold' id='" + item.maPhongBan + "' href='javascript:void(0)'>" + item.tenPhongBan + "</a>");
                    GetChildRecursive(child1.ToList(), item.maPhongBan);
                }
                else
                {
                    buildTree.Append("<li role='presentation'><a  class=\"hvr-underline-from-center\" id='" + item.maPhongBan + "' href='javascript:void(0)'>" + item.tenPhongBan + "</a>");
                }

                buildTree.Append("</li>");

            }
            buildTree.Append("</ul>");
        }
        #endregion
    }
}