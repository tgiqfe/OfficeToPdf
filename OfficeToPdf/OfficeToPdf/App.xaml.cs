using OfficeToPdf.Lib;
using System.Configuration;
using System.Data;
using System.IO;
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
            foreach (var arg in e.Args)
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
            Item.Documents = new();
            paths.Select(x => new DocumentItem(x)).
                ToList().
                ForEach(x => Item.Documents.Add(x));
            Item.Setting = Setting.Load();
        }

        private void Application_Exit(object sender, ExitEventArgs e)
        {
            Item.Setting.Save();
        }
    }

}
