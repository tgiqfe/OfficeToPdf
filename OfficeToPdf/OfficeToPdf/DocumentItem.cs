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

        public DocumentItem(string path)
        {
            this.Name = Path.GetFileName(path);
            this.FilePath = path;
            this.DocumentType = Path.GetExtension(path).ToLower() switch
            {
                ".doc" => DocumentType.Word,
                ".docx" => DocumentType.Word,
                ".xls" => DocumentType.Excel,
                ".xlsx" => DocumentType.Excel,
                ".ppt" => DocumentType.PowerPoint,
                ".pptx" => DocumentType.PowerPoint,
                _ => DocumentType.Unknown,
            };
            this.Status = DocumentStatus.Waiting;
        }

        #region Inotify change

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        #endregion
    }
}
