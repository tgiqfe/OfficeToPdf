using OfficeToPdf.Lib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace OfficeToPdf
{
    internal class DocumentItem : INotifyPropertyChanged
    {
        public string Name { get; set; }
        public string FilePath { get; set; }
        public string OutputPath
        {
            get
            {
                return Path.Combine(
                    Path.GetDirectoryName(this.FilePath),
                    Path.GetFileNameWithoutExtension(this.FilePath) + ".pdf");
            }
        }
        public DocumentType DocumentType { get; set; }
        public DocumentStatus Status { get; set; }
        public bool? LastCheck { get; set; }

        #region Inotify change

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        #endregion
    }
}
