using System;
using System.IO;

namespace Electre_Customize_DotNet.Helpers.Global
{
    /// <summary>
    /// Shared file plumbing for the small single-purpose loggers (duplicate wires, incorrect cable
    /// group IDs, null group IDs). Each logger still decides its own log folder, file name and pattern.
    /// </summary>
    internal static class LogFileHelper
    {
        private static readonly object WriteLock = new object();   // parallel Excel workers log at the same time

        /// <summary>Writes "[yyyy-MM-dd HH:mm:ss]  message" to the console and appends it to the file. A failed write is only reported on the console.</summary>
        public static void AppendTimestamped(string logFilePath, string message)
        {
            string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}]  {message}";
            Console.WriteLine(logMessage);

            try
            {
                // Append the log message to the file
                lock (WriteLock)
                {
                    File.AppendAllText(logFilePath, logMessage + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write log: {ex.Message}");
            }
        }

        /// <summary>Deletes every file in the folder matching the pattern. Does nothing if the folder is missing; a failure is only reported on the console.</summary>
        public static void DeleteMatching(string logFolder, string searchPattern)
        {
            try
            {
                if (Directory.Exists(logFolder)) // Check if the log directory exists
                {
                    foreach (string file in Directory.GetFiles(logFolder, searchPattern))
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
