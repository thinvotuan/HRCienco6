using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using PagedList;

namespace BatDongSan.Utils.Paging
{
    public static class IEnumerableExtensions
    {
        #region IEnumerable<T> extensions

        public static IPagedList<T> ToPagedList<T>(this IEnumerable<T> source, int pageIndex, int pageSize)
        {
            return new PagedList<T>(source, pageIndex, pageSize);
        }

        public static IPagedList<T> ToPagedList<T>(this IEnumerable<T> source, int pageIndex, int pageSize, int totalCount)
        {
            return new PagedList<T>(source, pageIndex, pageSize, totalCount);
        }

        public static IPagedList<T> ToPagedList<T>(this IList<T> source, int pageIndex, int pageSize, bool custom, int? totalCount)
        {
            return new PagedList<T>(source, pageIndex, pageSize, custom, totalCount);
        }

        #endregion
    }
}