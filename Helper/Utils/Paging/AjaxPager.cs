using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.Web.Mvc.Ajax;
using System.Web.Routing;

namespace BatDongSan.Utils.Paging
{
    public class AjaxPager
    {
        private AjaxHelper ajaxHelper;
        private ViewContext viewContext;
        private readonly int pageSize;
        private readonly int currentPage;
        private readonly int totalItemCount;        
        private readonly RouteValueDictionary linkWithoutPageValuesDictionary;
        private readonly AjaxOptions ajaxOptions;

        public AjaxPager(AjaxHelper helper, ViewContext viewContext, int pageSize, int currentPage, int totalItemCount, AjaxOptions options, RouteValueDictionary valueDictionary)
        {
            this.ajaxHelper = helper;
            this.viewContext = viewContext;
            this.pageSize = pageSize;
            this.currentPage = currentPage;
            this.totalItemCount = totalItemCount;
            this.ajaxOptions = options;
            this.linkWithoutPageValuesDictionary = valueDictionary;
        }

        public string RenderHtml()
        {
            int pageCount = (int)Math.Ceiling(this.totalItemCount / (double)this.pageSize);
            int nrOfPagesToDisplay = 10;

            var sb = new StringBuilder();

            // Previous
            if (this.currentPage > 1)
            {
                sb.Append(GeneratePageLink("Trước", this.currentPage - 1));
            }
            else
            {
                sb.Append("<span class=\"disabled\"></span>");
            }

            int start = 1;
            int end = pageCount;

            if (pageCount > nrOfPagesToDisplay)
            {
                int middle = (int)Math.Ceiling(nrOfPagesToDisplay / 2d) - 1;
                int below = (this.currentPage - middle);
                int above = (this.currentPage + middle);

                if (below < 4)
                {
                    above = nrOfPagesToDisplay;
                    below = 1;
                }
                else if (above > (pageCount - 4))
                {
                    above = pageCount;
                    below = (pageCount - nrOfPagesToDisplay);
                }

                start = below;
                end = above;
            }

            if (start > 3)
            {
                sb.Append(GeneratePageLink("1", 1));
                sb.Append(GeneratePageLink("2", 2));
                sb.Append("<li><span>...</span></li>");
            }
            for (int i = start; i <= end; i++)
            {
                if (i == this.currentPage)
                {
                    sb.AppendFormat("<li class='active'><a href='#'>{0}</a></li>", i);
                }
                else
                {
                    sb.Append(GeneratePageLink(i.ToString(), i));
                }
            }
            if (end < (pageCount - 3))
            {
                sb.Append("<li><span>...</span></li>");
                sb.Append(GeneratePageLink((pageCount - 1).ToString(), pageCount - 1));
                sb.Append(GeneratePageLink(pageCount.ToString(), pageCount));
            }

            // Next
            if (this.currentPage < pageCount)
            {
                sb.Append(GeneratePageLink("Sau", (this.currentPage + 1)));
            }
            else
            {
                sb.Append("<span class=\"disabled\"></span>");
            }
            return sb.ToString();
        }

        private string GeneratePageLink(string linkText, int pageNumber)
        {
            var pageLinkValueDictionary = new RouteValueDictionary(this.linkWithoutPageValuesDictionary);
            pageLinkValueDictionary.Add("page", pageNumber);

            return "<li>" + ajaxHelper.ActionLink(linkText, pageLinkValueDictionary["action"].ToString(), pageLinkValueDictionary, ajaxOptions).ToString() + "</li>";
        }
    }
}