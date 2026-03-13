namespace BatDongSan.Models.VIP
{
    partial class LinqVIPDataContext
    {
        public LinqVIPDataContext()
            : base(global::System.Configuration.ConfigurationManager.ConnectionStrings["VIPConnectionString"].ToString(), mappingSource)
        {
            OnCreated();
            this.CommandTimeout = 3600;
        }
    }
}
