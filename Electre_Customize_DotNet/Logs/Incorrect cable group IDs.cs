using Electre_Customize_DotNet.Objects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Electre_Customize_DotNet.Logs
{
    public static class Incorrect_cable_group_IDs
    {

        public static  string logpath;
        public static string logFilName;

        public static readonly string LogFolder = Path.Combine(GlobalVar.StrtCmd, "Logs");

         static Incorrect_cable_group_IDs()
         {
            // Ensure the Logs directory exists
            if (!Directory.Exists(LogFolder))
            {
                Directory.CreateDirectory(LogFolder);
            }
         }

        public static string GetLogFilePath()
        {

            // If log file name already exists, return it
            if (!string.IsNullOrEmpty(logFilName))
            {
                return Path.Combine(LogFolder, logFilName);
            }

            // Otherwise, create a new log file name based on the current date
            logFilName = "Incorrect cable group IDS_" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".log";

            return Path.Combine(LogFolder, logFilName);




        }

        public static void WriteLog(string message)
        {
            
              string  logFilePath = GetLogFilePath();
           
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
        public  static void DeletePreviousLogs()
        {
            try
            {
                if (Directory.Exists(LogFolder)) // Check if the log directory exists
                {
                    foreach (string file in Directory.GetFiles(LogFolder, "Incorrect cable group IDS_*.log"))
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