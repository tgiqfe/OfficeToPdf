using OfficeToPdf.Lib;
using OfficeToPdf.Lib.Conf;
using System.Configuration;
using System.Data;
using System.Diagnostics;
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
            if (e.Args.Length == 0)
            {
                Shutdown();
                return;
            }

            if (e.Args.Length > 0)
            {
                using (var proc = new Process())
                {
                    proc.StartInfo.UseShellExecute = false;
                    proc.StartInfo.CreateNoWindow = false;
                    proc.StartInfo.WorkingDirectory = Path.Combine(
                        Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName), "script");

                    proc.StartInfo.FileName = "cmd";
                    proc.StartInfo.Arguments = "/c InitialSetup.bat";
                    proc.Start();
                    proc.WaitForExit();
                }

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
