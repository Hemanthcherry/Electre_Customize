using Electre_Customize_DotNet.Objects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Electre_Customize_DotNet.MainOperation;

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
            string logFilePath = GetLogFilePath();
            string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}]  {message}";
            Console.WriteLine(logMessage);

            try
            {
                // Append the log message to the file
                File.AppendAllText(logFilePath, logMessage + Environment.NewLine);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write log: {ex.Message}");
            }
        }

        public static void Info(string message)
        {
            WriteLog(message);
        }
        public static void DeletePreviousLogs()
        {
            try
            {
                if (Directory.Exists(LogFolder)) // Check if the log directory exists
                {
                    foreach (string file in Directory.GetFiles(LogFolder, "DuplicateWire_*.log"))
                    {
                        File.Delete(file); // Delete each log file
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to delete old logs: {ex.Message}");

            }
        }
    }
}
