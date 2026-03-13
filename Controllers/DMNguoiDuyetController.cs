using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Web;
using System.Web.Mvc;
using BatDongSan.Helper.Utils;
using BatDongSan.Models.DanhMuc;
using BatDongSan.Models.HeThong;
using BatDongSan.Models.NhanSu;
using BatDongSan.Models.PhieuDeNghi;

namespace BatDongSan.Controllers
{
    public class DMNguoiDuyetController : ApplicationController
    {
        BatDongSan.Models.DanhMuc.LinqDanhMucDataContext lqDM = new BatDongSan.Models.DanhMuc.LinqDanhMucDataContext();
        LinqNhanSuDataContext ns = new LinqNhanSuDataContext();
        LinqHeThongDataContext heThong = new LinqHeThongDataContext();
        LinqPhieuDeNghiDataContext lqPhieuDN = new LinqPhieuDeNghiDataContext();
        private tbl_NS_NhanVienChucDanh nhanVienChucDanh;
        private tbl_NS_NhanVienPhongBan nhanVienPhongBan;
        LinqThuanVietDataContext lqThuanViet = new LinqThuanVietDataContext();

        private tbl_NS_PhanCaNhanVien nhanVienPhanCa;
        //
        // GET: /DMNguoiDuyet/

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult SelectCapDuyet1(string id, string maCongViec, byte trangThai, string title, string maNhanVien, string lyDo, int? idQuiTrinh, string url)
        {
            try
            {
                ViewBag.DSNhanVienDuyet = GetUserNguoiDuyet1CapCongTac(maNhanVien);
                ViewBag.IdQuiTrinh = idQuiTrinh;
                int soNgayNghiPhep = 0;
                if (id.StartsWith("DNNP"))
                {
                    soNgayNghiPhep = (int?)heThong.GetTable<BatDongSan.Models.PhieuDeNghi.tbl_NS_PhieuNghiPhep>().Select(d => d.soNgayNghiPhepThucTe).FirstOrDefault() ?? 0;
                }
                ViewData["maPhieu"] = id;
                ViewData["maCongViec"] = maCongViec;
                ViewData["soNgayNghiPhep"] = soNgayNghiPhep;
                ViewData["trangThai"] = 0;
                ViewData["lyDo"] = lyDo;
                ViewData["tenNhanVien"] = (heThong.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == GetUser().manv).Select(d => d.ho).FirstOrDefault() ?? string.Empty) + " " + (heThong.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == GetUser().manv).Select(d => d.ten).FirstOrDefault() ?? string.Empty);
                ViewBag.URL = url;
                if (idQuiTrinh == 31)
                {
                    ViewData["lisQuiTrinhDuyet"] = heThong.sp_QuiTrinhDuyetAS(maCongViec).Where(d => d.Id == 31).ToList();
                }
                else
                {
                    ViewData["lisQuiTrinhDuyet"] = heThong.sp_QuiTrinhDuyetAS(maCongViec).Where(d => d.Id != 31).ToList();
                }
                return View("SelectCapDuyet1");
            }
            catch
            {
                return View();
            }

        }

        public ActionResult SelectCapDuyet(string id, string maCongViec, byte trangThai, string title, string maNhanVien, string lyDo, int? idQuiTrinh, string url)
        {
            try
            {
                ViewBag.DSNhanVienDuyet = GetUserNguoiDuyet1CapCongTac(maNhanVien);
                ViewData["maPhieu"] = id;
                ViewData["maCongViec"] = maCongViec;
                ViewData["soNgayNghiPhep"] = 0;
                ViewData["trangThai"] = 0;
                ViewData["lyDo"] = lyDo;
                ViewData["tenNhanVien"] = (heThong.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == GetUser().manv).Select(d => d.ho).FirstOrDefault() ?? string.Empty) + " " + (heThong.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == GetUser().manv).Select(d => d.ten).FirstOrDefault() ?? string.Empty);
                ViewBag.URL = url;
                ViewBag.IdQuiTrinh = idQuiTrinh;
                return View("SelectCapDuyet");
            }
            catch
            {
                return View();
            }

        }

        /// <summary>
        /// Chọn qui trình duyệt
        /// </summary>
        /// <param name="id"></param>
        /// <param name="maCongViec"></param>
        /// <param name="maQuiTrinhDuyet"></param>
        /// <returns></returns>
        public ActionResult SelectCapDuyetDong(string id, string maCongViec, int? maQuiTrinhDuyet, string url)
        {
            try
            {
                ViewBag.MaPhieu = id;
                ViewBag.MaCongViec = maCongViec;
                ViewBag.IDQuiTrinh = maQuiTrinhDuyet;
                ViewData["lisQuiTrinhDuyet"] = heThong.sp_QuiTrinhDuyetDong(maCongViec).ToList();
                ViewBag.URL = url;
                return View("SelectCapDuyetDong");
            }
            catch
            {
                return View();
            }

        }

        public ActionResult SelectCapDuyetDongERP(string id, string maCongViec, int? maQuiTrinhDuyet, string url)
        {

            try
            {
                ViewBag.MaPhieu = id;
                ViewBag.MaCongViec = maCongViec;
                ViewBag.IDQuiTrinh = maQuiTrinhDuyet;
                ViewData["lisQuiTrinhDuyet"] = lqThuanViet.sp_QuiTrinhDuyetAS(maCongViec).Select(d => new QuiTrinhDuyet
            {
                Id = d.Id,
                TenQuiTrinh = d.TenQuiTrinh,
                ChuoiBuocDuyet = d.chuoiBuocDuyet,
            }).ToList();

                ViewBag.URL = url;
                return View("SelectCapDuyetDongERP");
            }
            catch
            {
                return View();
            }

        }

        /// <summary>
        /// Lịch sử duyệt phiếu
        /// </summary>
        /// <param name="id"></param>
        /// <param name="maQuiTrinh"></param>
        /// <returns></returns>
        public ActionResult LichSuDuyetPhieu(string id, int? maQuiTrinh)
        {
            try
            {
                ViewBag.LichSuDP = heThong.sp_NS_QuiTrinhDuyet_LichSuDuyet(id, maQuiTrinh).ToList();
                return View();
            }
            catch
            {
                return View();
            }

        }

        public ActionResult TraLai(string id, string maCongViec, byte trangThai, string title, string maNhanVien, string lyDo, string url, int idQuiTrinh)
        {
            try
            {
                ViewData["maPhieu"] = id;
                ViewData["maCongViec"] = maCongViec;
                ViewData["trangThai"] = 0;
                ViewData["lyDo"] = lyDo;
                ViewData["tenNhanVien"] = (heThong.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == GetUser().manv).Select(d => d.ho).FirstOrDefault() ?? string.Empty) + " " + (heThong.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == GetUser().manv).Select(d => d.ten).FirstOrDefault() ?? string.Empty);
                ViewBag.URL = url;
                ViewBag.IDQuiTrinh = idQuiTrinh;
                return View("TraLai");
            }
            catch
            {
                return View();
            }

        }

        /// <summary>
        /// Trả lại phiếu cho người duyệt trước
        /// </summary>
        /// <param name="id"></param>
        /// <param name="maCongViec"></param>
        /// <param name="maQuiTrinhDuyet"></param>
        /// <returns></returns>
        public ActionResult TraLaiDong(string id, string tenBuocDuyet, int? maQuiTrinhDuyet, string maCongViec, string url)
        {
            try
            {
                ViewBag.MaPhieu = id;
                ViewBag.MaCongViec = maCongViec;
                ViewBag.IDQuiTrinh = maQuiTrinhDuyet;
                ViewBag.DanhSachTraLai = heThong.sp_QuiTrinhDuyet_TraLai(id, maQuiTrinhDuyet, tenBuocDuyet).ToList();
                ViewBag.URL = url;
                return View("TraLaiDong");
            }
            catch
            {
                return View();
            }

        }

        private List<NhanVienModel> GetUserNguoiDuyet1CapCongTac(string maNhanVien)
        {
            string maPhongBan = heThong.GetTable<tbl_NS_NhanVienPhongBan>().Where(d => d.maNhanVien == maNhanVien).Select(d => d.maPhongBan).FirstOrDefault() ?? string.Empty;

            var objNguoiDuyet = heThong.sp_NS_QuiTrinhDuyet(maNhanVien).Select(d => new NhanVienModel
            {
                maNhanVien = d.maNhanVienDuyet,
                hoVaTen = d.hoTen,
                tenPhongBan = d.departmentName
            }).ToList();
            return objNguoiDuyet;
        }

        /// <summary>
        /// Duyệt phiếu này
        /// </summary>
        /// <param name="id"></param>
        /// <param name="maCongViec"></param>
        /// <param name="maQuiTrinhDuyet"></param>
        /// <returns></returns>
        public ActionResult DuyetPhieu(string id, string maCongViec, int? maQuiTrinhDuyet, string url)
        {
            try
            {
                ViewBag.MaPhieu = id;
                ViewBag.MaCongViec = maCongViec;
                ViewBag.IDQuiTrinh = maQuiTrinhDuyet;
                var chiTietDuyet = DanhSachQuiTrinhDuyet(Convert.ToInt16(maQuiTrinhDuyet));
                ViewData["chiTietQTD"] = chiTietDuyet;
                DMNguoiDuyetController nd = new DMNguoiDuyetController();
                var duyet = nd.GetDetailByMaPhieuTheoQuiTrinhDong(id, maQuiTrinhDuyet ?? 0);
                int thuTuDuyet = (int?)chiTietDuyet.OrderByDescending(d => d.ThuTuDuyet).Select(d => d.ThuTuDuyet).FirstOrDefault() ?? 1;
                int thuTuDuyetHT = (int?)chiTietDuyet.OrderByDescending(d => d.MaBuocDuyet == (duyet == null ? string.Empty : duyet.maBuocDuyet)).Select(d => d.ThuTuDuyet).FirstOrDefault() ?? 1;
                ViewBag.TrangThaiDuyet = (int?)heThong.tbl_HT_DMNguoiDuyets.Where(d => d.maPhieu == id).OrderByDescending(d => d.ID).Select(d => d.trangThai).FirstOrDefault() ?? 0;
                ViewBag.Duyet = duyet;
                string hoTen = string.Empty;
                string tenHienThiFooter = string.Empty;
                List<ChiTietQuiTrinhDuyet> chiTiet = new List<ChiTietQuiTrinhDuyet>();
                ChiTietQuiTrinhDuyet ct;
                foreach (var it in chiTietDuyet)
                {
                    ct = new ChiTietQuiTrinhDuyet();
                    var cl = "color:red";
                    if (it.ThuTuDuyet < thuTuDuyetHT && duyet.maBuocDuyet != it.MaBuocDuyet)
                    {
                        cl = "color:Blue";
                    }
                    else if (it.ThuTuDuyet > thuTuDuyetHT)
                    {
                        cl = "color:black";
                    }
                    hoTen = (heThong.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == it.MaNhanVien).Select(d => d.ho).FirstOrDefault() ?? string.Empty) + " " + (heThong.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == it.MaNhanVien).Select(d => d.ten).FirstOrDefault() ?? string.Empty);
                    if (thuTuDuyet == it.ThuTuDuyet)
                    {

                        tenHienThiFooter = "(" + it.TenHienThiFooter + ")";
                    }
                    else
                    {
                        tenHienThiFooter = "(" + it.TenHienThiFooter + ")" + "=>";
                    }
                    ct.MaMau = cl;
                    ct.TenHienThiFooter = tenHienThiFooter;
                    ct.TenNhanVien = hoTen;
                    chiTiet.Add(ct);
                }
                ViewBag.ChiTietDuyet = chiTiet;
                ViewBag.URL = url;
                return View("DuyetPhieu");
            }
            catch
            {
                return View();
            }

        }

        /// <summary>
        /// Duyệt các phiếu đăng ký công tác, nghỉ phép, tăng ca
        /// </summary>
        /// <param name="id"></param>
        /// <param name="maCongViec"></param>
        /// <param name="maQuiTrinhDuyet"></param>
        /// <param name="url"></param>
        /// <returns></returns>
        public ActionResult DuyetPhieuDK(string id, string maCongViec, int? maQuiTrinhDuyet, string url, string maNhanVien)
        {
            try
            {
                ViewBag.MaPhieu = id;
                ViewBag.MaCongViec = maCongViec;
                ViewBag.IDQuiTrinh = maQuiTrinhDuyet;
                var chiTietDuyet = DanhSachQuiTrinhDuyet(Convert.ToInt16(maQuiTrinhDuyet));
                ViewData["chiTietQTD"] = chiTietDuyet;
                DMNguoiDuyetController nd = new DMNguoiDuyetController();
                var duyet = nd.GetDetailByMaPhieuTheoQuiTrinhDong(id, maQuiTrinhDuyet ?? 0);
                int thuTuDuyet = (int?)chiTietDuyet.OrderByDescending(d => d.ThuTuDuyet).Select(d => d.ThuTuDuyet).FirstOrDefault() ?? 1;
                int thuTuDuyetHT = (int?)chiTietDuyet.OrderByDescending(d => d.MaBuocDuyet == (duyet == null ? string.Empty : duyet.maBuocDuyet)).Select(d => d.ThuTuDuyet).FirstOrDefault() ?? 1;
                ViewBag.TrangThaiDuyet = (int?)heThong.tbl_HT_DMNguoiDuyets.Where(d => d.maPhieu == id).OrderByDescending(d => d.ID).Select(d => d.trangThai).FirstOrDefault() ?? 0;
                ViewBag.Duyet = duyet;
                string hoTen = string.Empty;
                string tenHienThiFooter = string.Empty;
                List<ChiTietQuiTrinhDuyet> chiTiet = new List<ChiTietQuiTrinhDuyet>();
                ChiTietQuiTrinhDuyet ct;
                foreach (var it in chiTietDuyet)
                {
                    ct = new ChiTietQuiTrinhDuyet();
                    var cl = "color:red";
                    if (it.ThuTuDuyet < thuTuDuyetHT && duyet.maBuocDuyet != it.MaBuocDuyet)
                    {
                        cl = "color:Blue";
                    }
                    else if (it.ThuTuDuyet > thuTuDuyetHT)
                    {
                        cl = "color:black";
                    }
                    hoTen = (heThong.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == it.MaNhanVien).Select(d => d.ho).FirstOrDefault() ?? string.Empty) + " " + (heThong.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == it.MaNhanVien).Select(d => d.ten).FirstOrDefault() ?? string.Empty);
                    if (thuTuDuyet == it.ThuTuDuyet)
                    {

                        tenHienThiFooter = "(" + it.TenHienThiFooter + ")";
                    }
                    else
                    {
                        tenHienThiFooter = "(" + it.TenHienThiFooter + ")" + "=>";
                    }
                    ct.MaMau = cl;
                    ct.TenHienThiFooter = tenHienThiFooter;
                    ct.TenNhanVien = hoTen;
                    chiTiet.Add(ct);
                }
                ViewBag.ChiTietDuyet = chiTiet;
                ViewBag.URL = url;
                ViewBag.MaNhanVien = GetUser().manv;
                return View("DuyetPhieuDK");
            }
            catch
            {
                return View();
            }

        }

        /// <summary>
        /// Danh sách qui trình duyệt của bước duyệt
        /// </summary>
        /// <param name="maQuiTrinhDuyet"></param>
        /// <returns></returns>
        public List<ChiTietQuiTrinhDuyet> DanhSachQuiTrinhDuyet(int maQuiTrinhDuyet)
        {
            var chiTiet = (from t in heThong.tbl_HT_QuiTrinhDuyet_BuocDuyets
                           join ct in heThong.tbl_HT_QuiTrinhDuyetChiTiets on t.maBuocDuyet equals ct.MaCapDuyet
                           where ct.IdQuiTrinh == maQuiTrinhDuyet
                           select new ChiTietQuiTrinhDuyet
                           {
                               ThuTuDuyet = ct.ThuTuDuyet ?? 1,
                               TrangThai = t.id,
                               MaBuocDuyet = t.maBuocDuyet,
                               TenBuocDuyet = t.tenBuocDuyet,
                               TenHienThiFooter = t.tenBuocDuyet,
                               MaNhanVien = ct.MaNhanVien
                           }).OrderBy(d => d.ThuTuDuyet).ToList();
            return chiTiet;
        }


        #region Duyệt qui trình động của các phiếu đề nghị

        /// <summary>
        /// Thông tin mã phiếu theo qui trình của từng bước duyệt
        /// </summary>
        /// <param name="maPhieu"></param>
        /// <param name="maQuiTrinh"></param>
        /// <returns></returns>
        public DMNguoiDuyet GetDetailByMaPhieuTheoQuiTrinhDong(string maPhieu, int maQuiTrinh)
        {
            var data = (from a in heThong.tbl_HT_DMNguoiDuyets
                        join b in heThong.GetTable<tbl_NS_NhanVien>() on a.maNV equals b.maNhanVien into nvTmp
                        from nvTmps in nvTmp.DefaultIfEmpty()
                        where a.maPhieu == maPhieu //&& ct.IdQuiTrinh   == maQuiTrinh
                        orderby a.ID descending
                        select new DMNguoiDuyet
                        {
                            maPhieu = a.maPhieu,
                            maCongViec = a.maCongViec,
                            nguoiDuyet = new NhanVienModel(a.maNV, nvTmps.ho + " " + nvTmps.ten),
                            ngayDuyet = a.ngayDuyet,
                            tinhTrang = (int)a.trangThai,
                            lyDo = a.lyDo
                        }).Skip(0).Take(1).ToList();
            var ls = from a in data
                     join bd in heThong.tbl_HT_QuiTrinhDuyet_BuocDuyets on a.tinhTrang equals bd.id
                     join ct in heThong.tbl_HT_QuiTrinhDuyetChiTiets on bd.maBuocDuyet equals ct.MaCapDuyet
                     where ct.IdQuiTrinh == maQuiTrinh
                     select new DMNguoiDuyet
                     {
                         maPhieu = a.maPhieu,
                         maCongViec = a.maCongViec,
                         nguoiDuyet = a.nguoiDuyet,
                         ngayDuyet = a.ngayDuyet,
                         tenBuocDuyet = bd.tenBuocDuyet,
                         tinhTrang = a.tinhTrang,
                         maBuocDuyet = bd.maBuocDuyet,
                         Id = Convert.ToInt32(ct.Id),
                         lyDo = a.lyDo
                     };
            return ls.FirstOrDefault() ?? new DMNguoiDuyet(maPhieu);
        }

        /// <summary>
        /// Duyệt theo qui trình động
        /// </summary>
        /// <param name="maPhieu"></param>
        /// <param name="maCongViec"></param>
        /// <param name="sendMail"></param>
        /// <param name="sendSMS"></param>
        /// <param name="maNhanVien"></param>
        /// <returns></returns>
        public JsonResult DuyeTheoQuiTrinhDong(string maPhieu, string maCongViec, bool? sendMail, bool? sendSMS, string maNhanVien, string tenNhanVien, int? soNgayNghiPhep, string lyDo, int? idQuiTrinh, string url)
        {
            try
            {
                var getDMNguoiDuyet = heThong.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == maPhieu).FirstOrDefault();
                List<tbl_HT_DMNguoiDuyet> nd = new List<tbl_HT_DMNguoiDuyet>();
                tbl_HT_DMNguoiDuyet ht;
                if (getDMNguoiDuyet == null)
                {
                    ht = new tbl_HT_DMNguoiDuyet();
                    ht.maCongViec = maCongViec;
                    ht.maPhieu = maPhieu;
                    ht.maNV = GetUser().manv;
                    ht.trangThai = 0;
                    ht.ngayDuyet = DateTime.Now;
                    ht.lyDo = lyDo;
                    nd.Add(ht);
                }

                ht = new tbl_HT_DMNguoiDuyet();
                var buocDuyet = heThong.sp_QuiTrinhDuyet_TrangTiepTheo(idQuiTrinh, maPhieu, 1).FirstOrDefault();
                ht.maCongViec = maCongViec;
                ht.maPhieu = maPhieu;
                ht.maNV = maNhanVien;
                ht.trangThai = Convert.ToByte(buocDuyet == null ? 0 : buocDuyet.trangThai);
                ht.ngayDuyet = DateTime.Now;
                ht.lyDo = lyDo;
                nd.Add(ht);
                heThong.tbl_HT_DMNguoiDuyets.InsertAllOnSubmit(nd);
                heThong.SubmitChanges();
                if (maPhieu.StartsWith("DKTC"))
                {
                    var phieu = lqPhieuDN.tbl_NS_PhieuTangCas.Where(d => d.soPhieu == maPhieu).FirstOrDefault();
                    // Bi Check if maNhanVienDuyet thuoc Tbl_DM_PhongBan lay ma quy trinh duyet khac
                    var CheckQT = lqPhieuDN.GetTable<tbl_DM_PhongBan>().Where(d => d.maNhanVienDuyet == phieu.nguoiLap).FirstOrDefault();
                    if (CheckQT != null)
                    {
                        idQuiTrinh = 31;
                    }
                    phieu.maQuiTrinhDuyet = idQuiTrinh;
                    lqPhieuDN.SubmitChanges();


                    if (ht.trangThai == 3)
                    {
                        var ds = heThong.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == maPhieu && d.maCongViec == maCongViec).FirstOrDefault();
                        ds.maNV = null;
                        heThong.SubmitChanges();
                    }
                    if (ht.trangThai == 4)
                    {
                        var ds = heThong.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == maPhieu && d.maCongViec == maCongViec && d.trangThai == 3).FirstOrDefault();
                        ds.maNV = GetUser().manv;
                        heThong.SubmitChanges();
                    }
                }
                else if (maPhieu.StartsWith("DKCT"))
                {
                    var phieu = lqPhieuDN.tbl_NS_PhieuCongTacs.Where(d => d.maPhieu == maPhieu).FirstOrDefault();
                    phieu.maQuiTrinhDuyet = idQuiTrinh;
                    lqPhieuDN.SubmitChanges();
                    if (ht.trangThai == 3)
                    {
                        var ds = heThong.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == maPhieu && d.maCongViec == maCongViec).FirstOrDefault();
                        ds.maNV = null;
                        heThong.SubmitChanges();
                    }
                    if (ht.trangThai == 4)
                    {
                        var ds = heThong.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == maPhieu && d.maCongViec == maCongViec && d.trangThai == 3).FirstOrDefault();
                        ds.maNV = GetUser().manv;
                        heThong.SubmitChanges();
                    }
                }
                else if (maPhieu.StartsWith("DKNP"))
                {
                    var phieu = lqPhieuDN.tbl_NS_PhieuNghiPheps.Where(d => d.maPhieu == maPhieu).FirstOrDefault();
                    phieu.maQuiTrinhDuyet = idQuiTrinh;
                    lqPhieuDN.SubmitChanges();
                    if (ht.trangThai == 3)
                    {
                        var ds = heThong.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == maPhieu && d.maCongViec == maCongViec).FirstOrDefault();
                        ds.maNV = null;
                        heThong.SubmitChanges();
                    }
                    if (ht.trangThai == 4)
                    {
                        var ds = heThong.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == maPhieu && d.maCongViec == maCongViec && d.trangThai == 3).FirstOrDefault();
                        ds.maNV = GetUser().manv;
                        heThong.SubmitChanges();
                    }
                }
                else if (maPhieu.StartsWith("DNDC"))
                {
                    var phieu = lqPhieuDN.tbl_NS_PhieuDieuChuyens.Where(d => d.soPhieu == maPhieu).FirstOrDefault();
                    phieu.maQuiTrinhDuyet = idQuiTrinh;
                    lqPhieuDN.SubmitChanges();
                    if (ht.trangThai == 3)
                    {
                        var ds = heThong.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == maPhieu && d.maCongViec == maCongViec).FirstOrDefault();
                        ds.maNV = null;
                        heThong.SubmitChanges();
                    }
                    if (ht.trangThai == 4)
                    {
                        var ds = heThong.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == maPhieu && d.maCongViec == maCongViec && d.trangThai == 3).FirstOrDefault();
                        ds.maNV = GetUser().manv;
                        heThong.SubmitChanges();

                        //Insert vào tbl_NS_NhanVienPhongBan khi điều chuyển nhân viên đó qua phòng ban mới
                        var nhanVienDieuChuyen = lqPhieuDN.tbl_NS_PhieuDieuChuyens.Where(d => d.soPhieu == maPhieu).FirstOrDefault();
                        if (nhanVienDieuChuyen != null)
                        {
                            tbl_NS_NhanVienPhongBan nvPhongBan = new tbl_NS_NhanVienPhongBan();
                            nvPhongBan.maNhanVien = nhanVienDieuChuyen.maNhanVien;
                            nvPhongBan.maPhongBan = nhanVienDieuChuyen.maPhongBanMoi;
                            nvPhongBan.nguoiLap = nhanVienDieuChuyen.nguoiLap;
                            nvPhongBan.ngayLap = (DateTime)phieu.ngayChuyen;
                            ns.tbl_NS_NhanVienPhongBans.InsertOnSubmit(nvPhongBan);

                            //Cập nhật bộ phận tính lương mới cho nhân viên
                            //var boPhanTinhLuong = ns.tbl_NS_BoPhanTinhLuongs.Where(d => d.maNhanVien == nhanVienDieuChuyen.maNhanVien).FirstOrDefault();
                            //if (boPhanTinhLuong != null)//Cập nhật lại bộ phận tính lương
                            //{
                            //    boPhanTinhLuong.maBoPhan = nhanVienDieuChuyen.boPhanTinhLuongMoi;
                            //    boPhanTinhLuong.tenBoPhan = ns.tbl_NS_BoPhanTinhLuongs.Where(d => d.maBoPhan == nhanVienDieuChuyen.boPhanTinhLuongMoi).Select(d => d.tenBoPhan).FirstOrDefault() ?? string.Empty;
                            //}
                            //else//Insert vô bộ phận tính lương
                            //{
                            //    tbl_NS_BoPhanTinhLuong tl = new tbl_NS_BoPhanTinhLuong();
                            //    tl.maBoPhan = nhanVienDieuChuyen.boPhanTinhLuongMoi;
                            //    tl.tenBoPhan = ns.tbl_NS_BoPhanTinhLuongs.Where(d => d.maBoPhan == nhanVienDieuChuyen.boPhanTinhLuongMoi).Select(d => d.tenBoPhan).FirstOrDefault() ?? string.Empty;
                            //    tl.maNhanVien = nhanVienDieuChuyen.maNhanVien;
                            //    ns.tbl_NS_BoPhanTinhLuongs.InsertOnSubmit(tl);
                            //}
                            ns.SubmitChanges();
                        }
                    }
                }
                else if (maPhieu.StartsWith("DNTV"))
                {
                    var phieu = lqPhieuDN.tbl_NS_PhieuThoiViecs.Where(d => d.soPhieu == maPhieu).FirstOrDefault();
                    phieu.maQuiTrinhDuyet = idQuiTrinh;
                    lqPhieuDN.SubmitChanges();
                    if (ht.trangThai == 3)
                    {
                        var ds = heThong.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == maPhieu && d.maCongViec == maCongViec).FirstOrDefault();
                        ds.maNV = null;
                        heThong.SubmitChanges();
                    }
                    if (ht.trangThai == 4)
                    {
                        var ds = heThong.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == maPhieu && d.maCongViec == maCongViec && d.trangThai == 3).FirstOrDefault();
                        ds.maNV = GetUser().manv;
                        heThong.SubmitChanges();

                        //Cập nhật lại trạng thái của mã nhân viên đó là 1 
                        var nv = ns.tbl_NS_NhanViens.Where(d => d.maNhanVien == phieu.maNhanVien).FirstOrDefault();
                        nv.trangThai = 1;
                        //Cập nhật lại status của sys_user là 0
                        var us = heThong.Sys_Users.Where(d => d.manv == phieu.maNhanVien).FirstOrDefault();
                        us.status = false;
                        heThong.SubmitChanges();
                        ns.SubmitChanges();
                    }
                }
                else if (maPhieu.StartsWith("DNDLL"))
                {
                    var phieu = lqPhieuDN.tbl_NS_PhieuDiLamLais.Where(d => d.soPhieu == maPhieu).FirstOrDefault();
                    phieu.maQuiTrinhDuyet = idQuiTrinh;
                    lqPhieuDN.SubmitChanges();
                    if (ht.trangThai == 3)
                    {
                        var ds = heThong.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == maPhieu && d.maCongViec == maCongViec).FirstOrDefault();
                        ds.maNV = null;
                        heThong.SubmitChanges();
                    }
                    if (ht.trangThai == 4)
                    {
                        var ds = heThong.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == maPhieu && d.maCongViec == maCongViec && d.trangThai == 3).FirstOrDefault();
                        ds.maNV = GetUser().manv;
                        heThong.SubmitChanges();

                        //Cập nhật lại trạng thái của mã nhân viên đó là 0 
                        var nv = ns.tbl_NS_NhanViens.Where(d => d.maNhanVien == phieu.maNhanVien).FirstOrDefault();
                        nv.trangThai = 0;
                        //Cập nhật lại status của sys_user là 1
                        var us = heThong.Sys_Users.Where(d => d.manv == phieu.maNhanVien).FirstOrDefault();
                        us.status = true;
                        heThong.SubmitChanges();
                        ns.SubmitChanges();
                    }
                }
                if (sendMail == true)
                {
                    SendEmail(maNhanVien, maPhieu, maCongViec, url);
                }
                if (sendSMS == true)
                {
                    string soDienThoai = heThong.GetTable<tbl_NS_NhanVien>().Where(t => t.maNhanVien == maNhanVien).Select(d => d.phoneNumber1).FirstOrDefault() ?? string.Empty;
                    string linkURL = System.Web.Configuration.WebConfigurationManager.AppSettings["URL"];
                    SendSMSTV(soDienThoai, "Ban co so phieu : " + maPhieu + " can duoc xu ly gap. Vui long truy cap link : " + url + " de xu ly. Xin cam on.");
                }
                return Json(string.Empty);
            }
            catch (Exception ex)
            {
                return Json(ex.ToString());
            }
        }


        /// <summary>
        /// Không duyệt phiếu
        /// </summary>
        /// <param name="maPhieu"></param>
        /// <param name="maCongViec"></param>
        /// <param name="sendMail"></param>
        /// <param name="sendSMS"></param>
        /// <param name="maNhanVien"></param>
        /// <param name="tenNhanVien"></param>
        /// <param name="soNgayNghiPhep"></param>
        /// <param name="lyDo"></param>
        /// <param name="idQuiTrinh"></param>
        /// <returns></returns>
        public JsonResult TraLaiNew(string maPhieu, string maCongViec, bool? sendMail, bool? sendSMS, string maNhanVien, string tenNhanVien, int? soNgayNghiPhep, string lyDo, int? idQuiTrinh)
        {
            try
            {
                List<tbl_HT_DMNguoiDuyet> nd = new List<tbl_HT_DMNguoiDuyet>();
                tbl_HT_DMNguoiDuyet ht;
                ht = new tbl_HT_DMNguoiDuyet();
                ht.maCongViec = maCongViec;
                ht.maPhieu = maPhieu;
                ht.maNV = GetUser().manv;
                int trangThai = (int?)heThong.tbl_HT_QuiTrinhDuyet_BuocDuyets.Where(d => d.maBuocDuyet == "KDD").Select(d => d.id).FirstOrDefault() ?? 0;
                ht.trangThai = (byte)trangThai;
                ht.ngayDuyet = DateTime.Now;
                ht.lyDo = lyDo;
                nd.Add(ht);
                heThong.tbl_HT_DMNguoiDuyets.InsertAllOnSubmit(nd);
                heThong.SubmitChanges();
                return Json(string.Empty);
            }
            catch (Exception ex)
            {
                return Json(ex.ToString());
            }
        }

        /// <summary>
        /// Trả lại phiếu duyệt trước đó
        /// </summary>
        /// <param name="maCongViec"></param>
        /// <param name="maPhieu"></param>
        /// <param name="maQuiTrinh"></param>
        /// <param name="lyDo"></param>
        /// <param name="hoTen"></param>
        /// <returns></returns>
        public JsonResult TraLaiPhieuDuyet(string maCongViec, string maPhieu, int maQuiTrinh, string lyDo, string hoTen, bool? sendMail, bool? sendSMS, string url)
        {
            try
            {
                string maNV = string.Empty;
                int idHT = heThong.tbl_HT_DMNguoiDuyets.Where(d => d.maPhieu == maPhieu && d.maCongViec == maCongViec).OrderByDescending(d => d.ID).Select(d => d.ID).FirstOrDefault();
                maNV = heThong.tbl_HT_DMNguoiDuyets.Where(d => d.maCongViec == maCongViec && d.maPhieu == maPhieu && d.ID != idHT).OrderByDescending(d => d.ID).Select(d => d.maNV).FirstOrDefault() ?? string.Empty;
                var buocDuyet = heThong.sp_QuiTrinhDuyet_TrangTiepTheo(maQuiTrinh, maPhieu, 0).ToList();
                tbl_HT_DMNguoiDuyet item = new tbl_HT_DMNguoiDuyet();
                item.maCongViec = maCongViec;
                item.maPhieu = maPhieu;
                item.ngayDuyet = DateTime.Now;
                item.lyDo = hoTen + ": " + lyDo;
                item.trangThai = (byte)buocDuyet.FirstOrDefault().trangThai;
                item.maNV = maNV;
                //buocDuyet.FirstOrDefault().MaNhanVien;              
                item.trangThaiTraLai = 1;
                heThong.tbl_HT_DMNguoiDuyets.InsertOnSubmit(item);
                heThong.SubmitChanges();
                //var ds = heThong.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == maPhieu && d.maCongViec == maCongViec && d.trangThaiTraLai == 1).FirstOrDefault();
                //if (ds != null)
                //{
                //    if (ds.maNV == null || ds.maNV == string.Empty)
                //    {
                //        string maNhanVienTL = heThong.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == maPhieu && d.maCongViec == maCongViec && d.trangThai == ds.trangThai && (d.maNV != null || d.maNV != string.Empty)).Select(d => d.maNV).FirstOrDefault() ?? string.Empty;
                //        ds.maNV = maNhanVienTL;
                //        maNV = maNhanVienTL;
                //        heThong.SubmitChanges();
                //    }
                //}
                if (sendMail == true)
                {
                    SendEmail(maNV, maPhieu, maCongViec, url);
                }
                if (sendSMS == true)
                {
                    string soDienThoai = heThong.GetTable<tbl_NS_NhanVien>().Where(t => t.maNhanVien == (buocDuyet == null ? string.Empty : maNV)).Select(d => d.phoneNumber1).FirstOrDefault() ?? string.Empty;
                    string linkURL = System.Web.Configuration.WebConfigurationManager.AppSettings["URL"];
                    SendSMSTV(soDienThoai, "Ban co so phieu : " + maPhieu + " can duoc xu ly gap. Vui long truy cap link : " + linkURL + " de xu ly. Xin cam on.");
                }
                return Json(String.Empty);
            }
            catch
            {
                return Json("Error");
            }
        }

        /// <summary>
        /// Trả lại phiếu trước đó
        /// </summary>
        /// <param name="maPhieu"></param>
        /// <param name="nhaVienDuyet"></param>
        /// <param name="maCongViec"></param>
        /// <param name="sendMail"></param>
        /// <param name="sendSMS"></param>
        /// <param name="maNhanVien"></param>
        /// <param name="tenNhanVien"></param>
        /// <param name="lyDo"></param>
        /// <param name="idQuiTrinh"></param>
        /// <returns></returns>
        public JsonResult TraLaiNewDong(string maPhieu, string nhaVienDuyet, string maCongViec, bool? sendMail, bool? sendSMS, string maNhanVien, string tenNhanVien, string lyDo, int? idQuiTrinh, string url)
        {
            try
            {
                var noiDung = string.Empty;
                var tenGoiThau = string.Empty;
                var tenDuAn = string.Empty;
                string[] chuoi = nhaVienDuyet.Split('/');
                List<tbl_HT_DMNguoiDuyet> nd = new List<tbl_HT_DMNguoiDuyet>();
                tbl_HT_DMNguoiDuyet ht;
                ht = new tbl_HT_DMNguoiDuyet();
                ht.maCongViec = maCongViec;
                ht.maPhieu = maPhieu;
                ht.maNV = chuoi[0];
                int trangThai = Convert.ToInt32(chuoi[1]);
                ht.trangThai = (byte)trangThai;
                ht.ngayDuyet = DateTime.Now;
                ht.trangThaiTraLai = 1;
                ht.lyDo = lyDo;
                nd.Add(ht);
                heThong.tbl_HT_DMNguoiDuyets.InsertAllOnSubmit(nd);
                heThong.SubmitChanges();
                if (sendMail == true)
                {
                    SendEmail(ht.maNV, maPhieu, maCongViec, url);
                }
                if (sendSMS == true)
                {
                    string soDienThoai = heThong.GetTable<Sys_User>().Where(t => t.manv == ht.maNV).Select(d => d.telephone).FirstOrDefault() ?? string.Empty;
                    string linkURL = System.Web.Configuration.WebConfigurationManager.AppSettings["URL"];
                    SendSMSTV(soDienThoai, "Ban co so phieu : " + maPhieu + " can duoc xu ly gap. Vui long truy cap link : " + linkURL + " de xu ly. Xin cam on.");
                }
                return Json(string.Empty);
            }
            catch (Exception ex)
            {
                return Json(ex.ToString());
            }
        }

        #endregion

        #region Duyệt qui trình động có mã nhân viên mặc định theo quy trình duyệt
        public JsonResult DuyeTheoQuiTrinhDongMNV(string maPhieu, string maCongViec, bool? sendMail, bool? sendSMS, string maNhanVien, string tenNhanVien, string lyDo, int? idQuiTrinh, string url)
        {
            try
            {
                var getDMNguoiDuyet = heThong.tbl_HT_DMNguoiDuyets.OrderByDescending(d => d.ID).Where(d => d.maPhieu == maPhieu && d.maCongViec == maCongViec).FirstOrDefault();
                List<tbl_HT_DMNguoiDuyet> nd = new List<tbl_HT_DMNguoiDuyet>();
                tbl_HT_DMNguoiDuyet ht;
                if (getDMNguoiDuyet == null)
                {
                    ht = new tbl_HT_DMNguoiDuyet();
                    ht.maCongViec = maCongViec;
                    ht.maPhieu = maPhieu;
                    ht.maNV = GetUser().manv;
                    ht.trangThai = 0;
                    ht.ngayDuyet = DateTime.Now;
                    ht.lyDo = lyDo;
                    nd.Add(ht);
                    if (maCongViec.Equals("PhieuTamUng"))
                    {
                        var phieuTamUng = ns.tbl_NS_PhieuTamUngs.Where(d => d.soPhieu == maPhieu).FirstOrDefault();
                        phieuTamUng.maQuiTrinhDuyet = idQuiTrinh;
                        ns.SubmitChanges();
                    }
                    else if (maCongViec.Equals("ThuNhapKhac"))
                    {
                        var phieuTNK = ns.tbl_NS_ThuNhapKhacs.Where(d => d.id == Convert.ToInt32(maPhieu)).FirstOrDefault();
                        phieuTNK.maQuiTrinhDuyet = idQuiTrinh;
                        ns.SubmitChanges();
                    }
                    else if (maCongViec.Equals("HopDongLaoDong"))
                    {
                        var hopDong = ns.tbl_NS_HopDongLaoDongs.Where(d => d.soHopDong == maPhieu).FirstOrDefault();
                        hopDong.maQuiTrinhDuyet = idQuiTrinh;
                        ns.SubmitChanges();
                    }
                    else if (maCongViec.Equals("PhuLucHopDongLaoDong"))
                    {
                        var phuLuc = ns.tbl_NS_PhuLucHopDongs.Where(d => d.soPhuLuc == maPhieu).FirstOrDefault();
                        phuLuc.maQuiTrinhDuyet = idQuiTrinh;
                        ns.SubmitChanges();
                    }
                    else if (maCongViec.Equals("PhieuDieuChuyen"))
                    {
                        var dieuChuyen = lqPhieuDN.tbl_NS_PhieuDieuChuyens.Where(d => d.soPhieu == maPhieu).FirstOrDefault();
                        dieuChuyen.maQuiTrinhDuyet = idQuiTrinh;
                        lqPhieuDN.SubmitChanges();
                    }
                    else if (maCongViec.Equals("PhieuDiLamLai"))
                    {
                        var phieuDLL = lqPhieuDN.tbl_NS_PhieuDiLamLais.Where(d => d.soPhieu == maPhieu).FirstOrDefault();
                        phieuDLL.maQuiTrinhDuyet = idQuiTrinh;
                        lqPhieuDN.SubmitChanges();
                    }
                    else if (maCongViec.Equals("PhieuThoiViec"))
                    {
                        var phieuThoiViec = lqPhieuDN.tbl_NS_PhieuThoiViecs.Where(d => d.soPhieu == maPhieu).FirstOrDefault();
                        phieuThoiViec.maQuiTrinhDuyet = idQuiTrinh;
                        lqPhieuDN.SubmitChanges();
                    }
                    else if (maCongViec.Equals("PhuCapTheoCongTrinh"))
                    {
                        var phuCapTheoCT = lqDM.tbl_NS_PhuCapTheoCongTrinhs.Where(d => d.maPhuCap == maPhieu).FirstOrDefault();
                        phuCapTheoCT.maQuiTrinhDuyet = idQuiTrinh;
                        lqDM.SubmitChanges();
                    }
                    else if (maCongViec.Equals("PhieuBoNhiem"))
                    {
                        var phieuBoNhiem = lqPhieuDN.tbl_NS_PhieuDieuChuyens.Where(d => d.soPhieu == maPhieu).FirstOrDefault();
                        phieuBoNhiem.maQuiTrinhDuyet = idQuiTrinh;
                        lqPhieuDN.SubmitChanges();
                    }
                    else if (maCongViec.Equals("PhieuDieuChuyenBN"))
                    {
                        var phieuBoNhiem = lqPhieuDN.tbl_NS_PhieuDieuChuyens.Where(d => d.soPhieu == maPhieu).FirstOrDefault();
                        phieuBoNhiem.maQuiTrinhDuyet = idQuiTrinh;
                        lqPhieuDN.SubmitChanges();
                    }
                    else if (maCongViec.Equals("PhieuCapNhatNgayCong"))
                    {
                        var phieuCapNhat = lqPhieuDN.tbl_NS_PhieuCapNhatNgayCongs.Where(d => d.maPhieuCapNhatNgayCong == maPhieu).FirstOrDefault();
                        phieuCapNhat.maQuiTrinhDuyet = idQuiTrinh;
                        lqPhieuDN.SubmitChanges();
                    }
                }
                else if (getDMNguoiDuyet != null && getDMNguoiDuyet.trangThai == 0)
                {
                    if (maCongViec.Equals("PhieuTamUng"))
                    {
                        var phieuTamUng = ns.tbl_NS_PhieuTamUngs.Where(d => d.soPhieu == maPhieu).FirstOrDefault();
                        phieuTamUng.maQuiTrinhDuyet = idQuiTrinh;
                        ns.SubmitChanges();
                    }
                    else if (maCongViec.Equals("ThuNhapKhac"))
                    {
                        var phieuTNK = ns.tbl_NS_ThuNhapKhacs.Where(d => d.id == Convert.ToInt32(maPhieu)).FirstOrDefault();
                        phieuTNK.maQuiTrinhDuyet = idQuiTrinh;
                        ns.SubmitChanges();
                    }
                    else if (maCongViec.Equals("HopDongLaoDong"))
                    {
                        var hopDong = ns.tbl_NS_HopDongLaoDongs.Where(d => d.soHopDong == maPhieu).FirstOrDefault();
                        hopDong.maQuiTrinhDuyet = idQuiTrinh;
                        ns.SubmitChanges();
                    }
                    else if (maCongViec.Equals("PhuLucHopDongLaoDong"))
                    {
                        var phuLuc = ns.tbl_NS_PhuLucHopDongs.Where(d => d.soPhuLuc == maPhieu).FirstOrDefault();
                        phuLuc.maQuiTrinhDuyet = idQuiTrinh;
                        ns.SubmitChanges();
                    }
                    else if (maCongViec.Equals("PhieuDieuChuyen"))
                    {
                        var dieuChuyen = lqPhieuDN.tbl_NS_PhieuDieuChuyens.Where(d => d.soPhieu == maPhieu).FirstOrDefault();
                        dieuChuyen.maQuiTrinhDuyet = idQuiTrinh;
                        lqPhieuDN.SubmitChanges();
                    }
                    else if (maCongViec.Equals("PhieuDiLamLai"))
                    {
                        var phieuDLL = lqPhieuDN.tbl_NS_PhieuDiLamLais.Where(d => d.soPhieu == maPhieu).FirstOrDefault();
                        phieuDLL.maQuiTrinhDuyet = idQuiTrinh;
                        lqPhieuDN.SubmitChanges();
                    }
                    else if (maCongViec.Equals("PhieuThoiViec"))
                    {
                        var phieuThoiViec = lqPhieuDN.tbl_NS_PhieuThoiViecs.Where(d => d.soPhieu == maPhieu).FirstOrDefault();
                        phieuThoiViec.maQuiTrinhDuyet = idQuiTrinh;
                        lqPhieuDN.SubmitChanges();
                    }
                    else if (maCongViec.Equals("PhuCapTheoCongTrinh"))
                    {
                        var phuCapTheoCT = lqDM.tbl_NS_PhuCapTheoCongTrinhs.Where(d => d.maPhuCap == maPhieu).FirstOrDefault();
                        phuCapTheoCT.maQuiTrinhDuyet = idQuiTrinh;
                        lqDM.SubmitChanges();
                    }
                    else if (maCongViec.Equals("PhieuBoNhiem"))
                    {
                        var phieuBoNhiem = lqPhieuDN.tbl_NS_PhieuDieuChuyens.Where(d => d.soPhieu == maPhieu).FirstOrDefault();
                        phieuBoNhiem.maQuiTrinhDuyet = idQuiTrinh;
                        lqPhieuDN.SubmitChanges();
                    }
                    else if (maCongViec.Equals("PhieuDieuChuyenBN"))
                    {
                        var phieuBoNhiem = lqPhieuDN.tbl_NS_PhieuDieuChuyens.Where(d => d.soPhieu == maPhieu).FirstOrDefault();
                        phieuBoNhiem.maQuiTrinhDuyet = idQuiTrinh;
                        lqPhieuDN.SubmitChanges();
                    }
                    else if (maCongViec.Equals("PhieuCapNhatNgayCong"))
                    {
                        var phieuCapNhat = lqPhieuDN.tbl_NS_PhieuCapNhatNgayCongs.Where(d => d.maPhieuCapNhatNgayCong == maPhieu).FirstOrDefault();
                        phieuCapNhat.maQuiTrinhDuyet = idQuiTrinh;
                        lqPhieuDN.SubmitChanges();
                    }
                }

                ht = new tbl_HT_DMNguoiDuyet();
                var buocDuyet = heThong.sp_QuiTrinhDuyet_TrangTiepTheo(idQuiTrinh, maPhieu, 1).FirstOrDefault();
                ht.maCongViec = maCongViec;
                ht.maPhieu = maPhieu;
                ht.maNV = buocDuyet.MaNhanVien;
                ht.trangThai = Convert.ToByte(buocDuyet == null ? 0 : buocDuyet.trangThai);
                ht.ngayDuyet = DateTime.Now;
                ht.lyDo = lyDo;
                nd.Add(ht);
                heThong.tbl_HT_DMNguoiDuyets.InsertAllOnSubmit(nd);
                heThong.SubmitChanges();
                if (maCongViec.Equals("PhieuThoiViec"))
                {
                    int trangThai = (int?)heThong.tbl_HT_DMNguoiDuyets.Where(d => d.maCongViec == "PhieuThoiViec" && d.maPhieu == maPhieu).OrderByDescending(d => d.ID).Select(d => d.trangThai).FirstOrDefault() ?? 0;
                    if (trangThai == 2)
                    {
                        string maNVThoiViec = lqPhieuDN.tbl_NS_PhieuThoiViecs.Where(d => d.soPhieu == maPhieu).Select(d => d.maNhanVien).FirstOrDefault() ?? string.Empty;
                        var ds = ns.tbl_NS_NhanViens.Where(d => d.maNhanVien == maNVThoiViec).FirstOrDefault();
                        ds.trangThai = 1;
                        ns.SubmitChanges();
                    }
                }
                if (maCongViec.Equals("PhieuDiLamLai"))
                {
                    int trangThai = (int?)heThong.tbl_HT_DMNguoiDuyets.Where(d => d.maCongViec == "PhieuDiLamLai" && d.maPhieu == maPhieu).OrderByDescending(d => d.ID).Select(d => d.trangThai).FirstOrDefault() ?? 0;
                    if (trangThai == 2)
                    {
                        string maNVThoiViec = lqPhieuDN.tbl_NS_PhieuDiLamLais.Where(d => d.soPhieu == maPhieu).Select(d => d.maNhanVien).FirstOrDefault() ?? string.Empty;
                        var ds = ns.tbl_NS_NhanViens.Where(d => d.maNhanVien == maNVThoiViec).FirstOrDefault();
                        ds.trangThai = 0;
                        ns.SubmitChanges();
                    }
                }
                if (maCongViec.Equals("PhieuDieuChuyen"))
                {
                    int trangThai = (int?)heThong.tbl_HT_DMNguoiDuyets.Where(d => d.maCongViec == "PhieuDieuChuyen" && d.maPhieu == maPhieu).OrderByDescending(d => d.ID).Select(d => d.trangThai).FirstOrDefault() ?? 0;
                    if (trangThai == 2)
                    {
                        var rowNVDC = lqPhieuDN.tbl_NS_PhieuDieuChuyens.Where(d => d.soPhieu == maPhieu).FirstOrDefault();
                        if (rowNVDC != null)
                        {
                            var maNhanVienDC = rowNVDC.maNhanVien;
                            var maPhongBanMoiDC = rowNVDC.maPhongBanMoi;
                            var maPhanCaMoiDC = rowNVDC.maPhanCaMoi;
                            var maChucDanhMoiDC = rowNVDC.maChucDanhMoi;
                            var BoPhanTinhLuongMoiDC = rowNVDC.boPhanTinhLuongMoi;
                            DateTime? ngayChuyenMoiDC = rowNVDC.ngayChuyen;//DateTime.ParseExact(rowNVDC.ngayChuyen.Value.ToString() ?? String.Empty, "dd/MM/yyyy", CultureInfo.InvariantCulture);

                            // Insert tbl_NS_NhanVienChucDanh, tbl_NS_NhanVienPhongBan, tbl_NS_PhanCaNhanVien 

                            nhanVienChucDanh = new tbl_NS_NhanVienChucDanh();
                            nhanVienChucDanh.maChucDanh = maChucDanhMoiDC;
                            nhanVienChucDanh.maNhanVien = maNhanVienDC;
                            nhanVienChucDanh.ngayLap = rowNVDC.ngayChuyen;
                            nhanVienChucDanh.nguoiLap = GetUser().manv;

                            nhanVienPhongBan = new tbl_NS_NhanVienPhongBan();
                            nhanVienPhongBan.maNhanVien = maNhanVienDC;
                            nhanVienPhongBan.maPhongBan = maPhongBanMoiDC;
                            nhanVienPhongBan.ngayLap = rowNVDC.ngayChuyen ?? DateTime.Now;
                            nhanVienPhongBan.nguoiLap = GetUser().manv;

                            nhanVienPhanCa = new tbl_NS_PhanCaNhanVien();
                            nhanVienPhanCa.maNhanVien = maNhanVienDC;

                            nhanVienPhanCa.maPhanCa = short.Parse(!string.IsNullOrEmpty(maPhanCaMoiDC) ? maPhanCaMoiDC : "0");
                            nhanVienPhanCa.ngayLap = ngayChuyenMoiDC;
                            nhanVienPhanCa.nguoiLap = GetUser().manv;
                            nhanVienPhanCa.ngayApDung = ngayChuyenMoiDC ?? DateTime.Now;
                            // Update vao table tbl_NS_NhanVien
                            var ds = ns.tbl_NS_NhanViens.Where(d => d.maNhanVien == maNhanVienDC).FirstOrDefault();
                            ds.idKhoiTinhLuong = Convert.ToInt32(BoPhanTinhLuongMoiDC);

                            lqDM.tbl_NS_PhanCaNhanViens.InsertOnSubmit(nhanVienPhanCa);
                            ns.tbl_NS_NhanVienPhongBans.InsertOnSubmit(nhanVienPhongBan);
                            ns.tbl_NS_NhanVienChucDanhs.InsertOnSubmit(nhanVienChucDanh);
                            lqDM.SubmitChanges();
                            ns.SubmitChanges();

                        }
                    }
                }
                //Phieu dieu chuyen BoNhiem
                if (maCongViec.Equals("PhieuDieuChuyenBN"))
                {
                    int trangThai = (int?)heThong.tbl_HT_DMNguoiDuyets.Where(d => d.maCongViec == "PhieuDieuChuyenBN" && d.maPhieu == maPhieu).OrderByDescending(d => d.ID).Select(d => d.trangThai).FirstOrDefault() ?? 0;
                    if (trangThai == 2)
                    {
                        var rowNVDC = lqPhieuDN.tbl_NS_PhieuDieuChuyens.Where(d => d.soPhieu == maPhieu).FirstOrDefault();
                        if (rowNVDC != null)
                        {
                            var maNhanVienDC = rowNVDC.maNhanVien;
                            var maPhongBanMoiDC = rowNVDC.maPhongBanMoi;
                            var maPhanCaMoiDC = rowNVDC.maPhanCaMoi;
                            var maChucDanhMoiDC = rowNVDC.maChucDanhMoi;
                            var BoPhanTinhLuongMoiDC = rowNVDC.boPhanTinhLuongMoi;
                            DateTime? ngayChuyenMoiDC = rowNVDC.ngayChuyen;

                            // Insert tbl_NS_NhanVienChucDanh, tbl_NS_NhanVienPhongBan, tbl_NS_PhanCaNhanVien 

                            nhanVienChucDanh = new tbl_NS_NhanVienChucDanh();
                            nhanVienChucDanh.maChucDanh = maChucDanhMoiDC;
                            nhanVienChucDanh.maNhanVien = maNhanVienDC;
                            nhanVienChucDanh.ngayLap = rowNVDC.ngayChuyen;
                            nhanVienChucDanh.nguoiLap = GetUser().manv;

                            nhanVienPhongBan = new tbl_NS_NhanVienPhongBan();
                            nhanVienPhongBan.maNhanVien = maNhanVienDC;
                            nhanVienPhongBan.maPhongBan = maPhongBanMoiDC;
                            nhanVienPhongBan.ngayLap = (DateTime)rowNVDC.ngayChuyen;
                            nhanVienPhongBan.nguoiLap = GetUser().manv;

                            nhanVienPhanCa = new tbl_NS_PhanCaNhanVien();
                            nhanVienPhanCa.maNhanVien = maNhanVienDC;

                            nhanVienPhanCa.maPhanCa = short.Parse(!string.IsNullOrEmpty(maPhanCaMoiDC) ? maPhanCaMoiDC : "0");
                            nhanVienPhanCa.ngayLap = DateTime.Now;
                            nhanVienPhanCa.nguoiLap = GetUser().manv;
                            nhanVienPhanCa.ngayApDung = ngayChuyenMoiDC ?? DateTime.Now;
                            // Update vao table tbl_NS_NhanVien
                            var ds = ns.tbl_NS_NhanViens.Where(d => d.maNhanVien == maNhanVienDC).FirstOrDefault();
                            ds.idKhoiTinhLuong = Convert.ToInt32(BoPhanTinhLuongMoiDC);

                            lqDM.tbl_NS_PhanCaNhanViens.InsertOnSubmit(nhanVienPhanCa);
                            ns.tbl_NS_NhanVienPhongBans.InsertOnSubmit(nhanVienPhongBan);
                            ns.tbl_NS_NhanVienChucDanhs.InsertOnSubmit(nhanVienChucDanh);
                            lqDM.SubmitChanges();
                            ns.SubmitChanges();

                        }
                    }
                }
                //Phieu bo nhiem
                if (maCongViec.Equals("PhieuBoNhiem"))
                {
                    int trangThai = (int?)heThong.tbl_HT_DMNguoiDuyets.Where(d => d.maCongViec == "PhieuBoNhiem" && d.maPhieu == maPhieu).OrderByDescending(d => d.ID).Select(d => d.trangThai).FirstOrDefault() ?? 0;
                    if (trangThai == 2)
                    {
                        var rowNVDC = lqPhieuDN.tbl_NS_PhieuDieuChuyens.Where(d => d.soPhieu == maPhieu).FirstOrDefault();
                        if (rowNVDC != null)
                        {
                            var maNhanVienDC = rowNVDC.maNhanVien;
                            var maPhongBanMoiDC = rowNVDC.maPhongBanMoi;
                            var maPhanCaMoiDC = rowNVDC.maPhanCaMoi;
                            var maChucDanhMoiDC = rowNVDC.maChucDanhMoi;
                            var BoPhanTinhLuongMoiDC = rowNVDC.boPhanTinhLuongMoi;
                            DateTime? ngayChuyenMoiDC = rowNVDC.ngayChuyen;

                            // Insert tbl_NS_NhanVienChucDanh, tbl_NS_NhanVienPhongBan, tbl_NS_PhanCaNhanVien 

                            nhanVienChucDanh = new tbl_NS_NhanVienChucDanh();
                            nhanVienChucDanh.maChucDanh = maChucDanhMoiDC;
                            nhanVienChucDanh.maNhanVien = maNhanVienDC;
                            nhanVienChucDanh.ngayLap = rowNVDC.ngayChuyen;
                            nhanVienChucDanh.nguoiLap = GetUser().manv;

                            nhanVienPhongBan = new tbl_NS_NhanVienPhongBan();
                            nhanVienPhongBan.maNhanVien = maNhanVienDC;
                            nhanVienPhongBan.maPhongBan = maPhongBanMoiDC;
                            nhanVienPhongBan.ngayLap = (DateTime)rowNVDC.ngayChuyen;
                            nhanVienPhongBan.nguoiLap = GetUser().manv;

                            nhanVienPhanCa = new tbl_NS_PhanCaNhanVien();
                            nhanVienPhanCa.maNhanVien = maNhanVienDC;

                            nhanVienPhanCa.maPhanCa = short.Parse(!string.IsNullOrEmpty(maPhanCaMoiDC) ? maPhanCaMoiDC : "0");
                            nhanVienPhanCa.ngayLap = (DateTime)rowNVDC.ngayChuyen;
                            nhanVienPhanCa.nguoiLap = GetUser().manv;
                            nhanVienPhanCa.ngayApDung = ngayChuyenMoiDC ?? DateTime.Now;
                            // Update vao table tbl_NS_NhanVien
                            var ds = ns.tbl_NS_NhanViens.Where(d => d.maNhanVien == maNhanVienDC).FirstOrDefault();
                            ds.idKhoiTinhLuong = Convert.ToInt32(BoPhanTinhLuongMoiDC);

                            lqDM.tbl_NS_PhanCaNhanViens.InsertOnSubmit(nhanVienPhanCa);
                            ns.tbl_NS_NhanVienPhongBans.InsertOnSubmit(nhanVienPhongBan);
                            ns.tbl_NS_NhanVienChucDanhs.InsertOnSubmit(nhanVienChucDanh);
                            lqDM.SubmitChanges();
                            ns.SubmitChanges();

                        }
                    }
                }
                try
                {
                    if (maCongViec.Equals("HopDongLaoDong") || (maCongViec.Equals("PhuLucHopDongLaoDong")))
                    {
                        string maLoaiHD = string.Empty;
                        int trangThai = (int?)heThong.tbl_HT_DMNguoiDuyets.Where(d =>
                            (d.maCongViec == "PhuLucHopDongLaoDong" || d.maCongViec == "HopDongLaoDong")
                            && d.maPhieu == maPhieu).OrderByDescending(d => d.ID).Select(d => d.trangThai).FirstOrDefault() ?? 0;
                        if (trangThai == 2 || trangThai == 4)
                        {
                            string nhanVienHD = string.Empty;
                            if (maCongViec == "HopDongLaoDong")
                            {
                                var hdld = ns.tbl_NS_HopDongLaoDongs.Where(d => d.soHopDong == maPhieu).FirstOrDefault();
                                maLoaiHD = hdld.maLoaiHopDong;
                                nhanVienHD = hdld == null ? string.Empty : hdld.maNhanVien;
                            }
                            else if (maCongViec == "PhuLucHopDongLaoDong")
                            {
                                var phuLucHDLD = ns.tbl_NS_PhuLucHopDongs.Where(d => d.soPhuLuc == maPhieu).FirstOrDefault();
                                maLoaiHD = phuLucHDLD.maLoaiHopDong;
                                var hdld = ns.tbl_NS_HopDongLaoDongs.Where(d => d.soHopDong == phuLucHDLD.soHopDong).FirstOrDefault();

                                nhanVienHD = hdld == null ? string.Empty : hdld.maNhanVien;
                            }
                            var updateNhanVien = ns.tbl_NS_NhanViens.Where(d => d.maNhanVien == nhanVienHD).FirstOrDefault();
                            if (updateNhanVien != null)
                            {
                                updateNhanVien.maHDPL = maPhieu;
                                ns.SubmitChanges();
                            }
                            if (maLoaiHD == "HĐCT" || maLoaiHD == "HD2Noi" || maLoaiHD == "THOIVU")
                            {
                                ns.sp_DN_CapNhatPhepNamNhanVienMoi(DateTime.Now.Year);
                            }
                        }
                    }
                }
                catch { }
                if (sendMail == true)
                {
                    if (buocDuyet != null && buocDuyet.trangThai != 2)
                    {
                        SendEmail(buocDuyet.MaNhanVien, maPhieu, maCongViec, url);
                    }
                }
                if (sendSMS == true)
                {
                    if (buocDuyet != null && buocDuyet.trangThai != 2)
                    {
                        string soDienThoai = heThong.GetTable<Sys_User>().Where(t => t.manv == (buocDuyet == null ? string.Empty : buocDuyet.MaNhanVien)).Select(d => d.telephone).FirstOrDefault() ?? string.Empty;
                        string linkURL = System.Web.Configuration.WebConfigurationManager.AppSettings["URL"];
                        SendSMSTV(soDienThoai, "Ban co so phieu : " + maPhieu + " can duoc xu ly gap. Vui long truy cap link : " + url + " de xu ly. Xin cam on.");
                    }
                }

                return Json(string.Empty);
            }
            catch (Exception ex)
            {
                return Json(ex.ToString());
            }
        }
        #endregion


        #region Gửi mail
        public bool SendEmail(string maNhanVien, string maPhieu, string maCongViec, string url)
        {
            try
            {
                string hoTen = (heThong.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == maNhanVien).Select(d => d.ho).FirstOrDefault() ?? string.Empty) + " " + (heThong.GetTable<tbl_NS_NhanVien>().Where(d => d.maNhanVien == maNhanVien).Select(d => d.ten).FirstOrDefault() ?? string.Empty);
                string email = (heThong.GetTable<Sys_User>().Where(d => d.manv == maNhanVien).Select(d => d.email).FirstOrDefault() ?? string.Empty);

                var congViec = heThong.Sys_CongViecs.Where(d => d.maCongViec == maCongViec).FirstOrDefault();
                MailHelper mailInit = new MailHelper();// lay cac tham so trong webconfig                
                StringBuilder content = new StringBuilder();
                content.Append("<h1>Email từ hệ thống nhân sự</h1><br />");
                content.Append("<p>Xin chào: " + hoTen + " !</p>");
                content.Append("<p>Bạn có 1 tài liệu đang chờ xử lý từ công việc <strong>" + (congViec == null ? string.Empty : (congViec.tenCongViec)) + "</strong> với số phiếu là: <strong>" + maPhieu + "</strong> </p>");
                //content.Append("<p>Link truy cập: " + Request.UrlReferrer.ToString() + " !</p>");
                content.Append("<p>Link truy cập: " + url + " !</p>");
                content.Append("<p style='font-style:italic'>Thanks and Regards!</p>");
                MailAddress toMail = new MailAddress(email, hoTen);
                mailInit.Body = content.ToString();
                mailInit.ToMail = toMail;
                mailInit.SendMail();

                return true;
            }
            catch
            {
                return false;
            }
        }
        #endregion

        #region Gửi sms
        public bool SendSMSTV(string mobile, string TextSMS)
        {
            SMSThuanViet SMSTV = new SMSThuanViet();
            if (mobile.Length > 9)
            {
                string sql = "INSERT INTO MessageOut(receiver, msg, status) VALUES('" + mobile + "', '" + TextSMS + "', 'send')";
                bool StatusSend = SMSTV.ExeCuteNonQuery(sql);
                return StatusSend;
            }
            else
            {
                return false;
            }
        }
        #endregion

        #region Duyệt qui trình động ERP
        public JsonResult DuyetTheoQuiTrinhDongERP(string maPhieu, int? maQuiTrinh, string maCongViec, bool? sendMail, bool? sendSMS, string url)
        {
            DMNguoiDuyetController _nguoiDuyet = new DMNguoiDuyetController();
            try
            {

                var duyetLuong = lqThuanViet.tbl_DeNghiChiLuongs.Where(d => d.maPhieu == maPhieu).FirstOrDefault();
                List<HT_DMNguoiDuyet> listND = new List<HT_DMNguoiDuyet>();
                HT_DMNguoiDuyet nd;
                var getDMNguoiDuyet = lqThuanViet.HT_DMNguoiDuyets.Where(t => t.maPhieu.Equals(maPhieu)).OrderByDescending(t => t.ID).FirstOrDefault();

                DMNguoiDuyet item = new DMNguoiDuyet();
                string error = "";
                if (getDMNguoiDuyet == null)
                {
                    nd = new HT_DMNguoiDuyet();
                    nd.maCongViec = maCongViec;
                    nd.maPhieu = maPhieu;
                    nd.ngayDuyet = DateTime.Now;
                    nd.trangThai = 0;
                    nd.maNV = GetUser().manv;
                    nd.noiDung = duyetLuong == null ? string.Empty : duyetLuong.noiDung;
                    nd.noiDungDuyet = HoVaTen(GetUser().manv);
                    listND.Add(nd);
                }
                var buocDuyet = lqThuanViet.sp_QuiTrinhDuyet_TrangTiepTheo(maQuiTrinh, maPhieu, 1).ToList();
                nd = new HT_DMNguoiDuyet();
                nd.maCongViec = maCongViec;
                nd.maPhieu = maPhieu;
                nd.ngayDuyet = DateTime.Now;
                nd.trangThai = (byte?)buocDuyet.Select(d => d.trangThai).FirstOrDefault() ?? 0;
                nd.maNV = buocDuyet.Select(d => d.MaNhanVien).FirstOrDefault() ?? string.Empty;
                nd.noiDung = duyetLuong == null ? string.Empty : duyetLuong.noiDung;
                listND.Add(nd);
                lqThuanViet.HT_DMNguoiDuyets.InsertAllOnSubmit(listND);
                //Cập nhật mã qui trình duyệt                 
                duyetLuong.maQuiTrinhDuyet = maQuiTrinh;
                lqThuanViet.SubmitChanges();

                if (sendMail == true)
                {
                    SendEmail(nd.maNV, maPhieu, maCongViec, url);
                }
                if (sendSMS == true)
                {
                    string soDienThoai = heThong.GetTable<Sys_User>().Where(t => t.manv == (buocDuyet == null ? string.Empty : nd.maNV)).Select(d => d.telephone).FirstOrDefault() ?? string.Empty;
                    string linkURL = System.Web.Configuration.WebConfigurationManager.AppSettings["URL"];
                    SendSMSTV(soDienThoai, "Ban co so phieu : " + maPhieu + " can duoc xu ly gap. Vui long truy cap link : " + linkURL + " de xu ly. Xin cam on.");
                }
                return Json(String.Empty);
            }
            catch (Exception e)
            {
                return Json("Lỗi: " + e.Message);
            }
        }
        #endregion
    }
}
