using Electre_Customize_DotNet.Objects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Electre_Customize_DotNet.Helpers.Global;

namespace Electre_Customize_DotNet.Logs
{
    public static class Null_Group_IDs
    {

        public static string logpath;
        public static string logFilName;

        public static readonly string LogFolder = Path.Combine(GlobalVar.StrtCmd, "Logs");

        static Null_Group_IDs()
        {
            // Ensure the Logs directory exists
            if (!Directory.Exists(LogFolder))
            {
                Directory.CreateDirectory(LogFolder);
            }
        }

        private static readonly object NameLock = new object();

        // the log file name is created on first use - guarded because parallel Excel workers may log at the same moment
        public static string GetLogFilePath()
        {
            lock (NameLock)
            {
                return GetLogFilePathCore();
            }
        }

        private static string GetLogFilePathCore()
        {
            // If log file name already exists, return it
            if (!string.IsNullOrEmpty(logFilName))
            {
                return Path.Combine(LogFolder, logFilName);
            }

            // Otherwise, create a new log file name based on the current date
            logFilName = "Null Group IDs_" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".log";

            return Path.Combine(LogFolder, logFilName);
        }

        public static void WriteLog(string message)
        {
            LogFileHelper.AppendTimestamped(GetLogFilePath(), message);
        }

        public static void Info(string message)
        {
            WriteLog(message);
        }
        public static void DeletePreviousLogs()
        {
            LogFileHelper.DeleteMatching(LogFolder, "Null group IDS_*.log");
        }
    }
}