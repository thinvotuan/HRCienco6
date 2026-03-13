using System;
using System.Reflection;
using log4net.Config;
using System.IO;

using System.Web.Routing;
using System.Text.RegularExpressions;

namespace BatDongSan.Helper.Utils
{
    public static class Log4Net
    {
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        #region Constructors
        static Log4Net()
        {
            log4net.Config.XmlConfigurator.Configure();
        }
        #endregion

        #region Methods
        ///
        /// <param name="logLevel" />
        ///
        public static void WriteLog(log4net.Core.Level logLevel, String log)
        {
            if (logLevel.Name.Equals(log4net.Core.Level.Debug.Name))
                Log.Debug(" | " + log);
            else if (logLevel.Name.Equals(log4net.Core.Level.Error.Name))
                Log.Error(" | " + log);
            else if (logLevel.Name.Equals(log4net.Core.Level.Fatal.Name))
                Log.Fatal(" | " + log);
            else if (logLevel.Name.Equals(log4net.Core.Level.Info.Name))
                Log.Info(" | " + log);
            else if (logLevel.Name.Equals(log4net.Core.Level.Warn.Name))
                Log.Warn(" | " + log);
        }
        #endregion

        /// <summary>
        /// Tao chuoi ghi log
        /// </summary>
        /// <param name="username">User name</param>
        /// <param name="pageName">Page name</param>
        /// <param name="item">Item</param>
        /// <param name="actionKey">
        /// 1: Them moi
        /// 2: Sua
        /// 3: Xoa
        /// 4: In
        /// 5: Gui duyet
        /// 6: Duyet
        /// </param>
        /// <param name="othertext"></param>
        /// <returns></returns>
        public static string CreateMessageLog(string username, string pageName, string item, int actionKey, string othertext)
        {
            return username + " | " + pageName + " | " + CreateActionName(actionKey) + " | " + item + (string.IsNullOrEmpty(othertext) ? "" : " | " + othertext);
        }

        /// <summary>
        /// Tao ten action
        /// </summary>
        /// <param name="index">
        /// 1: Them moi
        /// 2: Sua
        /// 3: Xoa
        /// 4: In
        /// 5: Gui duyet
        /// 6: Duyet
        /// </param>
        /// <returns></returns>
        public static string CreateActionName(int index)
        {
            switch (index)
            {
                case 1:
                    return "Create";
                case 2:
                    return "Edit";
                case 3:
                    return "Delete";
                case 4:
                    return "Print";
                case 5:
                    return "Send approve";
                case 6:
                    return "Approved";
            }
            return "";
        }
    }
}
