using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace BatDongSan.Helper.Utils
{
    public static class HtmlRequestHelper
    {
        public static RouteData GetRouteData()
        {
            var routeValues = HttpContext.Current.Request.RequestContext.RouteData.Values;
            RouteData routData = new RouteData();
            if (routeValues.ContainsKey("controller"))
                routData.Controller = (string)routeValues["controller"];
            else
                routData.Controller = string.Empty;

            if (routeValues.ContainsKey("action"))
                routData.Action = (string)routeValues["action"];      
            else
                routData.Action = string.Empty;
           
            return routData;
        }       
    }

    public class RouteData
    {
        public string Controller;
        public string Action;
        public string Areas;
    }
}
