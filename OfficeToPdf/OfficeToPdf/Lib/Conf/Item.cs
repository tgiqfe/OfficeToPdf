using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeToPdf.Lib.Conf
{
    internal class Item
    {
        public const string ProcessName = "OfficeToPdf";

        public static readonly string WorkingDirectory = Path.Combine(
            Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName));

        public static BindingParams BindingParams = null;
    }
}
