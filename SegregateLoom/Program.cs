using System.Configuration;
using System.Threading;

namespace SegregateLoom
{
    internal static class Program
    {

        static Mutex mutex = new Mutex(true, "{8f2b37ff-df23-4319-93ae-1f7e6b5afbaf}");
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            //Check other instance is running
            if (mutex.WaitOne(TimeSpan.Zero, true))
            {
                if (args.Length > 0)
                    GlobalVar.StrtCmd = args[0];

                GlobalVar.MsFolder = Path.Combine(args[0], "result");
               
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                // To customize application configuration such as set high DPI settings or default font,
                // see https://aka.ms/applicationconfiguration.
                ApplicationConfiguration.Initialize();
                Application.Run(new SegregateLoom());
                mutex.ReleaseMutex();
            }

            else
            {
                MessageBox.Show("Segregation Running.", "Loom Segregation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}