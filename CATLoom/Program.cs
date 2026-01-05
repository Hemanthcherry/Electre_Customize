namespace CATLoom
{
    internal static class Program
    {
        
        static Mutex mutex = new Mutex(true, "{8f2b37ff-df23-4319-93ae-1f7e6b5afbef}");
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

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                
                ApplicationConfiguration.Initialize();
                Application.Run(new CATLoomWindow());
                mutex.ReleaseMutex();
            }

            else
            {
                MessageBox.Show("CATLOOM Report generation.", "Loom Segregation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}