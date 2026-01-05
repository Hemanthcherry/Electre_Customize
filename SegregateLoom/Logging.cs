using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SegregateLoom
{
    public static class Logging
    {
        private static readonly string LogFolder = Path.Combine(GlobalVar.StrtCmd, "Logs");

        static Logging()
        {
            // Ensure the Logs directory exists
            if (!Directory.Exists(LogFolder))
            {
                Directory.CreateDirectory(LogFolder);
            }
        }

        private static string GetLogFilePath()
        {
            // Log file name based on the current date
            string logFileName = "MSLOOM_" + DateTime.Now.ToString("yyyy-MM-dd") + ".log";
            return Path.Combine(LogFolder, logFileName);
        }

        private static void WriteLog(string logLevel, string message)
        {
            string logFilePath = GetLogFilePath();
            string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{logLevel}] {message}";

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
            WriteLog("INFO", message);
        }

        public static void Warning(string message)
        {
            WriteLog("WARNING", message);
        }

        public static void Error(string message)
        {
            WriteLog("ERROR", message);
        }

        public static void Debug(string message)
        {
            WriteLog("DEBUG", message);
        }
    }
}
