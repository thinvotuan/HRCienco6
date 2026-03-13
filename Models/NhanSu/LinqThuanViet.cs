namespace BatDongSan.Models.NhanSu
{
    partial class LinqThuanVietDataContext
    {
        public LinqThuanVietDataContext()
            : base(global::System.Configuration.ConfigurationManager.ConnectionStrings["dbThuanVietConnectionString"].ToString(), mappingSource)
        {
            OnCreated();
            this.CommandTimeout = 3600;
        }
    }
}
