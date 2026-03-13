namespace BatDongSan.Models.PhieuDeNghi
{
    partial class LinqPhieuDeNghiDataContext
    {
        public LinqPhieuDeNghiDataContext()
            : base(global::System.Configuration.ConfigurationManager.ConnectionStrings["BatDongSanConnectionString"].ToString(), mappingSource)
        {
            OnCreated();
            this.CommandTimeout = 3600;
        }
    }
}
