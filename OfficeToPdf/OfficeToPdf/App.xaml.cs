using System.Configuration;
using System.Data;
using System.Windows;

namespace OfficeToPdf
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            List<string> paths = new();
            foreach(var arg in e.Args)
            {
                if (arg.Contains(";"))
                {
                    paths.AddRange(arg.Split(';'));
                }
                else
                {
                    paths.Add(arg);
                }
            }
        }

        private void Application_Exit(object sender, ExitEventArgs e)
        {

        }
    }

}
