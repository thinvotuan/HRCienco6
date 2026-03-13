namespace BatDongSan.Models.HeThong
{
    partial class LinqHeThongDataContext
    {
        public LinqHeThongDataContext()
            : base(global::System.Configuration.ConfigurationManager.ConnectionStrings["BatDongSanConnectionString"].ToString(), mappingSource)
        {
            OnCreated();
            this.CommandTimeout = 3600;
        }
    }
}
