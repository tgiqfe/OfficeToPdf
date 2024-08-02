using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeToPdf.Lib
{
    internal enum DocumentStatus
    {
        None,
        Waiting,
        Converting,
        Completed,
        Error,
    }
}
