using StartupHelper;

namespace TaskTrayWeek;

static class Program
{
    public static StartupManager StartupManager { get; private set; }
    /// <summary>
    ///  The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main()
    {
        StartupManager = new StartupManager("TaskTrayWeek", RegistrationScope.Global);
        // To customize application configuration such as set high DPI settings or default font,
        // see https://aka.ms/applicationconfiguration.
        ApplicationConfiguration.Initialize();
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new WeekApplicationContext());
    }
}