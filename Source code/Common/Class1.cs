using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SP.SpCommonFun
{
    using log4net;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.CompilerServices;
    using System.Text;
    using System.Threading.Tasks;

    public static class AppLogger
    {
        /// <summary>
        /// Gets the logger for a class
        /// </summary>
        private static ILog logger = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);


        public static void InitLogger()
        {
            log4net.Config.XmlConfigurator.Configure();
        }



        public static string ArchiveDirectory { get; set; }
        public static bool IsDebugEnabled { get; set; }
        public static string LogFileName { get; set; }
        public static int LogFilesRetentionHours { get; set; }
        public static string WorkObject { get; set; }

        public static void ArchiveOldLogFiles() { }
        public static void ClearWorkObject() { }

        public static void Debug(string logmessage, [CallerMemberName] string methodName = "")
        {
            if (logger.IsDebugEnabled)
            {
                logger.Debug(GetFormattedMessage(methodName, logmessage));
            }
        }
        public static void Error(Exception exception, string message, [CallerMemberName] string functionName = "")
        {
            logger.Error(GetFormattedMessage(functionName, message), exception);
        }

        public static void Error(string message, [CallerMemberName] string functionName = "")
        {
            logger.Error(GetFormattedMessage(functionName, message));
        }
        public static void Info(string message, [CallerMemberName] string functionName = "")
        {

            logger.Info(GetFormattedMessage(functionName, message));
            Console.WriteLine(message + functionName);
        }

        public static void SqlLogger(string message, [CallerMemberName] string functionName = "")
        {
            logger.Info(GetFormattedMessage(functionName, message.Replace(Environment.NewLine, "")));
        }

        public static void SetWorkObjectID(string workObject) { }

        public static void Warning(string message, [CallerMemberName]string functionName = "")
        {
            logger.Warn(GetFormattedMessage(functionName, message));
        }

        public static void Fatal(Exception ex, string message, [CallerMemberName]string functionName = "")
        {
            logger.Fatal(GetFormattedMessage(functionName, message), ex);
        }

        public static void Fatal(string message, [CallerMemberName]string functionName = "")
        {
            logger.Fatal(GetFormattedMessage(functionName, message));
        }

        private static string GetFormattedMessage(string methodName, string message)
        {
            return methodName + " | " + message;
        }
    }

}
