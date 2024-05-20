using OfficeToPdf.Lib.Conf;
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
            Item.BindingParams = new();
        }

        private void Application_Exit(object sender, ExitEventArgs e)
        {
            Item.BindingParams.Setting.Save();
        }
    }
}
