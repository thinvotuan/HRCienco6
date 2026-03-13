namespace BatDongSan.Models.DanhMuc
{
    partial class LinqDanhMucDataContext
    {
        public LinqDanhMucDataContext()
            : base(global::System.Configuration.ConfigurationManager.ConnectionStrings["BatDongSanConnectionString"].ToString(), mappingSource)
        {
            OnCreated();
            this.CommandTimeout = 3600;
        }
    }
}
