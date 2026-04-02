namespace NetworkInterfaceSwitcher
{
    //internal static class Program
    //{
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        //[STAThread]
        //static void Main()
        //{
        //    // To customize application configuration such as set high DPI settings or default font,
        //    // see https://aka.ms/applicationconfiguration.
        //    ApplicationConfiguration.Initialize();
        //    Application.Run(new NetworkInterfaceSwitcher());
        //}


    static class Program
    {
        [STAThread]
        static void Main()
        {
            // If running in an interactive desktop session, start the WinForms UI.
            // If running as a Windows Service (no interactive session), run the service.
            if (Environment.UserInteractive)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm());
            }
            else
            {
                // Run as service. To install the service, use sc.exe create or a service installer.
                System.ServiceProcess.ServiceBase[] ServicesToRun;
                ServicesToRun = new System.ServiceProcess.ServiceBase[]
                {
                    new NetworkInterfaceSwitcher.Service.NetworkInterfaceService()
                };
                System.ServiceProcess.ServiceBase.Run(ServicesToRun);
            }
        }
    }


}