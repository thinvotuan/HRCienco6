namespace BatDongSan.Models.QLSuatAn.LuuVet
{
    partial class QLSALuuVetDataContext
    {
        public QLSALuuVetDataContext()
            : base(global::System.Configuration.ConfigurationManager.ConnectionStrings["QuanLySuatAnNhanSuConnectionString"].ToString(), mappingSource)
        {
            OnCreated();
            this.CommandTimeout = 3600;
        }
    }
}
