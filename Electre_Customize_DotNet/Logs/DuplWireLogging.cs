using Electre_Customize_DotNet.Objects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Electre_Customize_DotNet.MainOperation;
using Electre_Customize_DotNet.Helpers.Global;

namespace Electre_Customize_DotNet.Logs
{

    public static class DuplWireLogging
    {
        //public static string logpath;
        private static readonly string LogFolder = Path.Combine(GlobalVar.StrtCmd, "Logs");

        //static DuplWireLogging()
        //{
        //    // Ensure the Logs directory exists
        //    if (!Directory.Exists(LogFolder))
        //    {
        //        Directory.CreateDirectory(LogFolder);
        //    }
        //}

        private static string GetLogFilePath()
        {
            // Log file name based on the current date
           // string logFileName = "DuplicateWire_" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".log";
            string logFileName = "DuplicateWire_" + DateTime.Now.ToString("dd-MM-yyyy-HH-mm") + ".log";
            return Path.Combine(LogFolder, logFileName);
        }

        private static void WriteLog(string message)
        {
            LogFileHelper.AppendTimestamped(GetLogFilePath(), message);
        }

        public static void Info(string message)
        {
            WriteLog(message);
        }
        public static void DeletePreviousLogs()
        {
            LogFileHelper.DeleteMatching(LogFolder, "DuplicateWire_*.log");
        }
    }
}
