namespace BatDongSan.Models.DBChamCong
{
    partial class DBChamCongDataContext
    {
        public DBChamCongDataContext()
            : base(global::System.Configuration.ConfigurationManager.ConnectionStrings["BDChamCongConnectionString"].ToString(), mappingSource)
        {
            OnCreated();
            this.CommandTimeout = 3600;
        }
    }
}
