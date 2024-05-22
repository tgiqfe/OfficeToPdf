using OfficeToPdf.Lib;
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
            if(e.Args.Length > 0)
            {
                var arguments = e.Args.Aggregate(new List<string>(), (list, arg) =>
                {
                    if (arg.Contains(";"))
                    {
                        list.AddRange(arg.Split(';'));
                    }
                    else
                    {
                        list.Add(arg);
                    }
                    return list;
                });
                Item.ConvertTargets = arguments.Select(arg => new ConvertTarget(arg)).ToList();
                Item.BindingParams = new();
            }
        }

        private void Application_Exit(object sender, ExitEventArgs e)
        {
            Item.BindingParams?.Setting.Save();
        }
    }
}
