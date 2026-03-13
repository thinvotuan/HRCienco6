using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace BatDongSan.Utils.Paging
{
    public interface IPagedList<T> : IList<T>
    {
        int PageCount { get; }
        int TotalItemCount { get; }
        int PageIndex { get; }
        int PageNumber { get; }
        int PageSize { get; }
        bool HasPreviousPage { get; }
        bool HasNextPage { get; }
        bool IsFirstPage { get; }
        bool IsLastPage { get; }
    }
}