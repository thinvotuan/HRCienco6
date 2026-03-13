namespace BatDongSan.Models.ERP
{
    partial class linqERPCienco6DataContext
    {
        public linqERPCienco6DataContext()
            : base(global::System.Configuration.ConfigurationManager.ConnectionStrings["ERPConnectionString"].ToString(), mappingSource)
        {
            OnCreated();
            this.CommandTimeout = 3600;
        }
    }
}
