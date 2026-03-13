using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

namespace BatDongSan.Helper.Utils
{
    public static class GenerateUtil
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
        public class RouteData
        {
            public string Controller;
            public string Action;
            public string Areas;
        }
        public static string CheckLetterPA(string tienToPhieu, string lastID)
        {
            StringBuilder sb = new StringBuilder();
            var date = DateTime.Now;
            string nam = date.Year.ToString();
            nam = nam.Remove(0, 2);
            string thang = string.Empty,
                ngay = string.Empty;
            if (date.Month < 10)
            {
                thang = "0" + date.Month;
            }
            else
            {
                thang = date.Month.ToString();
            }

            if (date.Day < 10)
            {
                ngay = "0" + date.Day;
            }
            else
            {
                ngay = date.Day.ToString();
            }

            string dateOfID = nam + thang + ngay;

            bool inDay = (lastID ?? string.Empty).Contains(dateOfID);
            if (inDay == false)
            {
                return tienToPhieu + nam + thang + ngay + "-0001";
            }
            else
            {
                int? idSuffix = Convert.ToInt32(lastID.Substring(lastID.Length - 4)) + 1;
                if (idSuffix <= 0 || idSuffix == null)
                    return tienToPhieu + nam + thang + ngay + "-0001";
                else
                {
                    sb.Append(idSuffix.ToString());
                    while (sb.Length < 4)
                    {
                        sb.Insert(0, "0");
                    }
                    return tienToPhieu + nam + thang + ngay + "-" + sb.ToString();
                }
            }
        }
        public static string CheckLetter(string tienToPhieu, string lastID)
        {
            StringBuilder sb = new StringBuilder();
            var date = DateTime.Now;
            string nam = date.Year.ToString();
            nam = nam.Remove(0, 2);
            string thang = string.Empty,
                ngay = string.Empty;
            if (date.Month < 10)
            {
                thang = "0" + date.Month;
            }
            else
            {
                thang = date.Month.ToString();
            }

            if (date.Day < 10)
            {
                ngay = "0" + date.Day;
            }
            else
            {
                ngay = date.Day.ToString();
            }

            string dateOfID = nam + thang + ngay;

            bool inDay = (lastID ?? string.Empty).Contains(dateOfID);
            if (inDay == false)
            {
                return tienToPhieu + nam + thang + ngay + "-001";
            }
            else
            {
                int? idSuffix = Convert.ToInt32(lastID.Substring(lastID.Length - 3)) + 1;
                if (idSuffix <= 0 || idSuffix == null)
                    return tienToPhieu + nam + thang + ngay + "-001";
                else
                {
                    sb.Append(idSuffix.ToString());
                    while (sb.Length < 3)
                    {
                        sb.Insert(0, "0");
                    }
                    return tienToPhieu + nam + thang + ngay + "-" + sb.ToString();
                }
            }
        }
    }
}