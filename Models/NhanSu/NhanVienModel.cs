using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;

namespace BatDongSan.Models.NhanSu
{
    public class NhanVienModel
    {

        public HopDongLaoDongModel hopDong { get; set; }
        public string maNhanVien { get; set; }

        public string ten { get; set; }

        public string ho { get; set; }

        public string hoVaTen { get; set; }

        public string tenGoiKhac { get; set; }

        public System.Nullable<bool> gioiTinh { get; set; }

        public string tenGioiTinh { get; set; }

        public System.Nullable<System.DateTime> ngaySinh { get; set; }

        public string namSinh { get; set; }

        public string noiSinh { get; set; }

        public System.Nullable<int> tinhTrangHonNhan { get; set; }

        public string honNhan { get; set; }

        public string nguyenQuan { get; set; }

        public System.Nullable<int> idDanToc { get; set; }

        public System.Nullable<int> idTonGiao { get; set; }

        public System.Nullable<int> idBangCap { get; set; }

        public System.Nullable<byte> idLoaiNhanVien { get; set; }

        public System.Nullable<int> idTrinhDoChuyenMon { get; set; }

        public string CMNDSo { get; set; }

        public string CMNDNoiCap { get; set; }

        public System.Nullable<System.DateTime> CMNDNgayCap { get; set; }

        public string phoneNumber1 { get; set; }

        public string phoneNumber2 { get; set; }

        public string email { get; set; }

        public System.Nullable<System.DateTime> ngayVaoLam { get; set; }

        public string maChamCong { get; set; }

        public System.Nullable<int> salaryTypeId { get; set; }

        public string soTaiKhoan { get; set; }

        public string maChiNhanhNganHang { get; set; }

        public string maSoThue { get; set; }

        public string thongTinKhac { get; set; }

        public string tenAnhDaiDien { get; set; }

        public string tenAnhDaiDienLuu { get; set; }

        public System.Data.Linq.Binary anhDaiDienBinary { get; set; }

        public string anhDaiDienURL { get; set; }

        public string facultyId { get; set; }

        public System.Nullable<bool> laNguoiNgoaiQuoc { get; set; }

        public string nguoiNgoaiQuoc { get; set; }

        public System.Nullable<int> trangThai { get; set; }

        public bool hoanTatHoSo { get; set; }

        public string cardPartyMember { get; set; }

        public System.Nullable<System.DateTime> admissionDate { get; set; }

        public string admissionPlace { get; set; }

        public string nguoiLap { get; set; }

        public System.Nullable<System.DateTime> ngayLap { get; set; }

        public bool dangBaoHiemNoiKhac { get; set; }

        public System.Nullable<bool> isTeachingEnglishESL { get; set; }

        public System.Nullable<int> trainingModuleId { get; set; }

        public string employeeIdOld { get; set; }

        public string maQuocTich { get; set; }

        public string WorkPermit { get; set; }

        public string maCapBacQL { get; set; }

        public System.Nullable<bool> isCommitment { get; set; }

        public string thuongTruSoNha { get; set; }

        public string thuongTruTenDuong { get; set; }

        public string thuongTruPhuongXa { get; set; }

        public System.Nullable<int> thuongTruQuanHuyen { get; set; }

        public string tenQuanHuyenThuongTru { get; set; }

        public string thuongTruTinhThanh { get; set; }

        public string tenTinhThanhThuongTru { get; set; }

        public string tamTruSoNha { get; set; }

        public string tamTruTenDuong { get; set; }

        public string tamTruPhuongXa { get; set; }

        public System.Nullable<int> tamTruQuanHuyen { get; set; }

        public string tenQuanHuyenTamTru { get; set; }

        public string tamTruTinhThanh { get; set; }

        public string tenTinhThanhTamTru { get; set; }

        public string nguoiLienLacKhanCap { get; set; }

        public string diaChiLienLacKhanCap { get; set; }

        public string phoneLienLacKhanCap { get; set; }

        public string cardGroupMember { get; set; }

        public System.Nullable<System.DateTime> admissionGroupDate { get; set; }

        public string admissionGroupPlace { get; set; }

        public string soBaoHiem { get; set; }

        public System.Nullable<bool> chuaCoMSThue { get; set; }

        public string soSoLaoDong { get; set; }

        public string EmployeeIdOldNew { get; set; }

        public System.Nullable<bool> isTeachingEnglishVoc { get; set; }

        public string emailInd { get; set; }

        public System.Nullable<int> loaiThue { get; set; }

        public string maPhongBan { get; set; }

        public string tenPhongBan { get; set; }

        public string maChiNhanhVanPhong { get; set; }

        public string tenChiNhanh { get; set; }

        public string maChucDanh { get; set; }

        public string tenChucDanh { get; set; }

        public string tenChiNhanhVanPhong { get; set; }

        public string tenDanToc { get; set; }

        public string tenTonGiao { get; set; }

        public string tenTrinhDoChuyenMon { get; set; }

        public string tenKhoiTinhLuong { get; set; }

        public string tenLoaiThue { get; set; }

        public string idBoPhanTinhLuong { get; set; }

        public string tenCapBacQL { get; set; }

        public int idKhoiTinhLuong { get; set; }

        public string tenLoaiNhanVien { get; set; }

        public string tenQuocTich { get; set; }

        public string tenChiNhanhNganHang { get; set; }

        public Image hinhAnh { get; set; }

        public IList<Sys_ChungChiNgoaiNgu> ngoaiNgus { get; set; }

        public IList<tbl_NS_NhanVienChungChiNgoaiNgu> nhanVienNgoaiNgus { get; set; }

        public IList<QuanHeGiaDinhModel> quanHeGiaDinhs { get; set; }

        public IList<Sys_FileDinhKem> fileDinhKems { get; set; }

        public string tenDangNhap { get; set; }

        public string soDienThoai { get; set; }

        public int userID { get; set; }

        public NhanVienModel(string maNV, string fullName)
        {
            this.maNhanVien = maNV;
            this.hoVaTen = fullName;
        }

        public NhanVienModel()
        {
            hopDong = new HopDongLaoDongModel();
        }


        public short maPhanCa { get; set; }

        public string maNgheNghiep { get; set; }

        public string tenNgheNghiep { get; set; }

        public string loaiKetNap { get; set; }

        public string tenLoaiKetNap { get; set; }

        public string matKhau { get; set; }

        public int? soCapBac { get; set; }

        public string userName { get; set; }

        public bool choPhepVaoERPKhac { get; set; }

        public System.Nullable<System.DateTime> ngayBatDauTinhPhepNam { get; set; }
    }
}

public class QuanHeGiaDinhModel
{

    public int id { get; set; }

    public string maNhanVien { get; set; }

    public string tenNguoiQuanHe { get; set; }

    public short idLoaiQuanHe { get; set; }

    public System.Nullable<System.DateTime> ngaySinh { get; set; }

    public string namSinh { get; set; }

    public string ngheNghiep { get; set; }

    public string diaChiHienTai { get; set; }

    public string CMNDSo { get; set; }

    public string noiCap { get; set; }

    public System.Nullable<System.DateTime> ngayCap { get; set; }

    public bool giamTruPhuThuoc { get; set; }

    public System.Nullable<decimal> soTienGiam { get; set; }

    public System.Nullable<System.DateTime> ngayBatDauGiam { get; set; }

    public System.Nullable<System.DateTime> ngayKetThucGiam { get; set; }

    public string nguoiLap { get; set; }

    public System.Nullable<System.DateTime> ngayLap { get; set; }

    public string tenLoaiQuanHe { get; set; }
}