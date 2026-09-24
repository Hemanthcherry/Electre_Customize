using System;
using System.IO;
using System.Linq;
using Electre_Customize_DotNet.Logs;

namespace Electre_Customize_DotNet.Helpers.Global
{
    internal static class ReportFolderCleaner
    {
        /// <summary>Deletes the .xls/.xlsx/.txt/.csv files (any letter case) directly inside the folder.</summary>
        public static void DeleteExistingFiles(string directoryPath)
        {
            if (!Directory.Exists(directoryPath))
            {
                Logging.Warning("Directory does not exist: " + directoryPath);
                return;
            }

            string[] extensionsToDelete = { ".xls", ".xlsx", ".txt", ".csv" };

            try
            {
                var files = Directory.GetFiles(directoryPath)
                                     .Where(file => extensionsToDelete.Contains(Path.GetExtension(file), StringComparer.OrdinalIgnoreCase));

                foreach (var file in files)
                {
                    try
                    {
                        File.Delete(file);
                        Logging.Info("Deleted: " + file);
                    }
                    catch (Exception ex)
                    {
                        Logging.Error($"Error deleting file {file}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Logging.Error($"Failed to retrieve files: {ex.Message}");
            }
        }
    }
}
