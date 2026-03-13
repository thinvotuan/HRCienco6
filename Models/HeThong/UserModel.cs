using BatDongSan.Models.HeThong;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace BatDongSan.Models.HeThong
{
    public class UserModel
    {
        public List<Sys_UserThuocNhom> NhomUsers { get; set; }  
        public UserModel()
        {
            NhomUsers = new List<Sys_UserThuocNhom>();
        }
        public int UserId { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }        
        public string Email { get; set; }
        public string Telephone { get; set; } 
        public string Note { get; set; }
        public bool Status { get; set; }
        public string MaNV { get; set; }
        public string MaNhom { get; set; }
        public string TenNhom { get; set; }
        public string HoTenNV { get; set; }        
    }
}