namespace BatDongSan.Models.QLSuatAn
{
    partial class QLSuatAnDataContext
    {
        public QLSuatAnDataContext()
            : base(global::System.Configuration.ConfigurationManager.ConnectionStrings["BatDongSanConnectionString"].ToString(), mappingSource)
        {
            OnCreated();
            this.CommandTimeout = 3600;
        }
    }
}
