namespace BatDongSan.Models.NhanSu
{
    partial class LinqNhanSuDataContext
    {
        public LinqNhanSuDataContext()
            : base(global::System.Configuration.ConfigurationManager.ConnectionStrings["BatDongSanConnectionString"].ToString(), mappingSource)
        {
            OnCreated();
            this.CommandTimeout = 3600;
        }
    }
}
