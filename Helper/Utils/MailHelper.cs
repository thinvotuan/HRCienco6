using System;
using System.Data;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Mvc.Ajax;
using System.Text;
using System.Net.Mail;
using System.ComponentModel;
using System.IO;

namespace BatDongSan.Helper.Utils
{
    public class MailHelper
    {
        public string Url { get; set; }
        public bool IsGmail { get; set; }
        public string SmtpServer { get; set; }
        public int Port { get; set; }
        public string Subject { get; set; }
        public string CCMail { get; set; }
        public string From { get; set; }
        public string Password { get; set; }
        public string FromName { get; set; }
        public string AttachmentPath { get; set; }
        public bool IsHtml { get; set; } // Content is html        

        //feild required
        public string Body { get; set; }
        public MailAddress ToMail { get; set; }
        public string[] listToEmail { get; set; }

        public MailHelper()
        {
            Url = ConfigurationManager.AppSettings["URL"];
            IsGmail = Convert.ToBoolean(ConfigurationManager.AppSettings["IsGmail"]);
            SmtpServer = ConfigurationManager.AppSettings["SmtpServer"];
            Port = Convert.ToInt16(ConfigurationManager.AppSettings["Port"]);
            Subject = ConfigurationManager.AppSettings["Subject"];
            CCMail = ConfigurationManager.AppSettings["CCMail"];
            From = ConfigurationManager.AppSettings["From"];
            Password = ConfigurationManager.AppSettings["Pass"];
            FromName = ConfigurationManager.AppSettings["FromName"];
            IsHtml = true; // Default is html
            listToEmail = new string[0];
        }

        /// <summary>
        ///     Send mail bu mail server not gmail
        /// </summary>      
        public bool SendMail()
        {
            //specify the mail server address
            SmtpClient mailClient = new SmtpClient();
            mailClient.Host = this.SmtpServer;
            mailClient.Port = this.Port;
            mailClient.DeliveryMethod = SmtpDeliveryMethod.Network;            
            mailClient.UseDefaultCredentials = false;

            mailClient.Credentials = new System.Net.NetworkCredential(this.From, this.Password);
            mailClient.EnableSsl = true;
            try
            {
                //set the addresses
                MailAddress fromMail = new MailAddress(this.From, this.FromName);
                //create the mail message
                MailMessage message = new MailMessage(fromMail, ToMail);
                // send to multi email                                  
                foreach (var item in listToEmail)
                {
                    if (!String.IsNullOrEmpty(item))
                        message.To.Add(item);
                }
                message.SubjectEncoding = System.Text.Encoding.UTF8;

                //set the content
                message.Body = this.Body;
                message.Subject = this.Subject;
                message.IsBodyHtml = this.IsHtml;
                //the userstate can be any object. The object can be accessed in the callback method
                //in this example, we will just use the MailMessage object.
                object userState = message;

                //wire up the event for when the Async send is completed
                mailClient.SendCompleted += new SendCompletedEventHandler(SmtpClient_OnCompleted);

                // Attach file is true
                if (!String.IsNullOrEmpty(AttachmentPath))
                {
                    Attachment attachment = new Attachment(AttachmentPath);
                    message.Attachments.Add(attachment);
                }

                //Begin send
                mailClient.Send(message);
                return true;
            }
            catch (Exception e)
            {
                //Lưu lịch sử hoạt động
                BatDongSan.Models.HeThong.LinqHeThongDataContext context = new BatDongSan.Models.HeThong.LinqHeThongDataContext();
                BatDongSan.Models.HeThong.Sys_LichSuHoatDong lichSu = new BatDongSan.Models.HeThong.Sys_LichSuHoatDong();
                lichSu.controller = GenerateUtil.GetRouteData().Controller;
                lichSu.action = GenerateUtil.GetRouteData().Action;
                lichSu.ngayLap = DateTime.Now;
                lichSu.nguoiLap = "NV-217";
                lichSu.noiDung = e.Message;
                context.Sys_LichSuHoatDongs.InsertOnSubmit(lichSu);
                context.SubmitChanges();

                /*END*/
                Log4Net.WriteLog(log4net.Core.Level.Error, "Sorry, an error occurred while send mail: " + e.Message);
                return false;
            }
        }

        public static void SmtpClient_OnCompleted(object sender, AsyncCompletedEventArgs e)
        {
            //Get the Original MailMessage object
            MailMessage mail = (MailMessage)e.UserState;

            //write out the subject
            string address = mail.To.Select(d => d.Address).FirstOrDefault();

            if (e.Cancelled)
            {
                Log4Net.WriteLog(log4net.Core.Level.Error, "Send canceled for mail address [" + address + "].");
            }
            if (e.Error != null)
            {
                Log4Net.WriteLog(log4net.Core.Level.Error, "Error " + address + " occurred when sending mail [" + e.Error.ToString() + "]");
            }
            else
            {
                Log4Net.WriteLog(log4net.Core.Level.Info, "Message to address [" + address + "] sent.");
            }
        }

        public string GetBodyMailMessageToUser()
        {
            string ret = String.Empty;
            //get path
            string path = System.Configuration.ConfigurationManager.AppSettings["StorePath"].ToString() + "\\Files";
            // create reader & open file            
            TextReader tr = new StreamReader(Path.Combine(path, "MailMessageToUser.txt"));

            // read a line of text
            ret = tr.ReadToEnd();

            // close the stream
            tr.Close();

            return ret;
        }

        public string GetBodyMailMessageAuto()
        {
            string ret = String.Empty;
            //get path
            string path = System.Configuration.ConfigurationManager.AppSettings["StorePath"].ToString() + "\\Files";
            // create reader & open file            
            TextReader tr = new StreamReader(Path.Combine(path, "MailMessageAuto.txt"));

            // read a line of text
            ret = tr.ReadToEnd();

            // close the stream
            tr.Close();

            return ret;
        }


    }
}
