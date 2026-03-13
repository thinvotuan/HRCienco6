using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using BatDongSan.Models.HeThong;

namespace BatDongSan.Models.HeThong
{
    public class VuViecRepository
    {
        private LinqHeThongDataContext context;
        public VuViecRepository()
        {
            context = new LinqHeThongDataContext();
        }
        public IQueryable<Sys_VuViec> GetVuViecByCongViecVaNhomUser(string maCongViec, string maNhomUser)
        {
            var vuViecs = from vuViec in context.Sys_VuViecs
                          join congViec_VuViec in context.Sys_CongViecVaVuViecs on vuViec.maVuViec equals congViec_VuViec.maVuViec
                          where congViec_VuViec.maCongViec == maCongViec && congViec_VuViec.maNhomUser == maNhomUser
                          select vuViec;
            return vuViecs;
        }

        public IQueryable<Sys_VuViec> GetVuViecByCongViecVaUserName(string userName, string maCongViec)
        {
            var vuViecs = from vuViec in context.Sys_VuViecs
                          join congViec_VuViec in context.Sys_CongViecVaVuViecs on vuViec.maVuViec equals congViec_VuViec.maVuViec
                          join nhomUser in context.Sys_NhomUsers on congViec_VuViec.maNhomUser equals nhomUser.maNhomUser
                          join userThuocNhom in context.Sys_UserThuocNhoms on nhomUser.maNhomUser equals userThuocNhom.maNhomUser
                          join user in context.Sys_Users on userThuocNhom.userId equals user.userId
                          where congViec_VuViec.maCongViec.Equals(maCongViec) && user.userName.Equals(userName)
                          select vuViec;
            return vuViecs;
        }
    }
}