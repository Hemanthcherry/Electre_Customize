using System.Configuration;
using System.Diagnostics;
using System.Text;
using Electre_Customize_DotNet.Logs;
using Electre_Customize_DotNet.Objects;
using Electre_Customize_DotNet.ReportUI;

namespace Electre_Customize_DotNet
{
    internal static class Program
    {
        // Define a mutex to check for a single instance of the application
        static Mutex mutex = new Mutex(true, "{8f2b37ff-df23-4319-93ae-1f7e6b5afbae}"); // Use a unique GUID

        [STAThread]
        static void Main(string[] args)
        {
            //Check other instance is running
            if (mutex.WaitOne(TimeSpan.Zero, true)) 
            {
                if (args.Length > 0)
                    GlobalVar.StrtCmd = args[0];

                //GlobalVar.StrtCmd = "C:\\ELECTRE\\ELECTRE_PROJECTS\\PANEL_DRAWING\\";
                Logging.Info($"Report extraction started from path {GlobalVar.StrtCmd}");

                GlobalVar.TempFolderGlobal = Path.Combine(GlobalVar.StrtCmd, ConfigurationManager.AppSettings["TempFolder"]);
                GlobalVar.DataExtractionGlobal = Path.Combine(GlobalVar.StrtCmd, ConfigurationManager.AppSettings["DataExtraction"]);
                GlobalVar.ReportFolderGlobal = Path.Combine(GlobalVar.StrtCmd, ConfigurationManager.AppSettings["ReportFolder"]);
              
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                ApplicationConfiguration.Initialize();
                Application.Run(new CustomReportWindow());
                mutex.ReleaseMutex();
            }

            else
            {
                MessageBox.Show("Already Running.", "Electre Report", MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
        }
    }
}