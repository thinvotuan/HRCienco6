using System;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;


namespace BatDongSan.Helper.Utils
{
    public class SMSThuanViet
    {
        SqlConnection con;
        SqlDataAdapter da;
        DataSet ds;
        SqlCommand cmd;
        string StringConnectDB = System.Configuration.ConfigurationManager.ConnectionStrings["dbSMSThuanVietConnectionString"].ConnectionString;
        public bool CheckDbConnection(string StringConnection)
        {
            try
            {
                if (con == null) con = new SqlConnection(StringConnection);
                if (con.State == ConnectionState.Closed) con.Open();
                return true;
            }
            catch (Exception)
            {
                return false;
            }

        }

        public void disconnect()
        {
            if ((con != null) && (con.State == ConnectionState.Open)) con.Close();
        }
        public DataSet get(string sql)
        {
            if (CheckDbConnection(StringConnectDB) == true)
            {
                da = new SqlDataAdapter(sql, con);
                ds = new DataSet();
                da.Fill(ds); disconnect();
                return ds;
            }
            else
            {
                return null;
            }

        }
        public bool ExeCuteNonQuery(string sql)
        {
            if (CheckDbConnection(StringConnectDB) == true)
            {
                cmd = new SqlCommand(sql, con);
                cmd.ExecuteNonQuery();
                disconnect();
                return true;
            }
            else
            {
                return false;
            }
        }

    }
}
