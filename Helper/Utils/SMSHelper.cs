using System;
using System.Data;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Net;
using System.IO;

namespace BatDongSan.Helper.Utils
{
    public class SMSHelper
    {
        public string User { get; set; }
        public string Password { get; set; }
        public string ApiId { get; set; }
        public string To { get; set; } // exp: 0840974900766 -> 84974900766
        public string Text { get; set; }

        public SMSHelper()
        {
            User = ConfigurationManager.AppSettings["UserSMS"];
            Password = ConfigurationManager.AppSettings["PasswordSMS"];
            ApiId = ConfigurationManager.AppSettings["api_id"];
        }

        public string SendSMS()
        {
            string s = "";
            try
            {
                //http://api.clickatell.com/http/sendmsg?user=worldsoft&password=PASSWORD&api_id=3383381&to=84974900766&text="AAA"

                WebClient client = new WebClient();
                // Add a user agent header in case the requested URI contains a query.
                client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                client.QueryString.Add("user", this.User);
                client.QueryString.Add("password", this.Password);
                client.QueryString.Add("api_id", this.ApiId);
                client.QueryString.Add("to", this.To);
                client.QueryString.Add("text", this.Text);
                string baseurl = "http://api.clickatell.com/http/sendmsg";
                Stream data = client.OpenRead(baseurl);
                StreamReader reader = new StreamReader(data);
                s = reader.ReadToEnd();
                data.Close();
                reader.Close();

                Log4Net.WriteLog(log4net.Core.Level.Info, "Message to phone number [" + this.To + "] sent.");
            }
            catch (Exception e)
            {
                s = e.Message;
                Log4Net.WriteLog(log4net.Core.Level.Error, "SMS: " + e.Message);
            }
            return s;
        }
    }
}
