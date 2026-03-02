namespace AutoClickScenarioTool
{
    internal static class Program
    {
        /// <summary>
        ///  アプリケーションのエントリーポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());
        }
    }
}