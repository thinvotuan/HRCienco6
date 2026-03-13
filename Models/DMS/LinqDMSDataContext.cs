namespace BatDongSan.Models.DMS
{
    partial class LinqDMSDataContextDataContext
    {
        public LinqDMSDataContextDataContext()
            : base(global::System.Configuration.ConfigurationManager.ConnectionStrings["DMSConnectionString"].ToString(), mappingSource)
        {
            OnCreated();
            this.CommandTimeout = 3600;
        }
    }
}
