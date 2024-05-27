using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeToPdf.Lib.Conf
{
    internal class BindingParams
    {
        public Setting Setting { get; set; }

        public BindingParams()
        {
            this.Setting = Setting.Load();
        }
    }
}
