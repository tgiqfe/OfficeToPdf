using OfficeToPdf.Lib;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeToPdf
{
    internal class Item
    {
        public static ObservableCollection<DocumentItem> Documents { get; set; }

        public static Setting Setting { get; set; }
    }
}
