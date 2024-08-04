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
    public class DocumentItem : INotifyPropertyChanged
    {
        public string Name { get; private set; }
        public string NameWithoutExtension { get; private set; }
        public string Extension { get; private set; }
        public string FilePath { get; private set; }
        public string OutputPath
        {
            get
            {
                return Path.Combine(
                    Path.GetDirectoryName(this.FilePath),
                    Path.GetFileNameWithoutExtension(this.FilePath) + ".pdf");
            }
        }
        public DocumentType DocumentType { get; private set; }
        public DocumentStatus Status { get; set; }
        public bool? LastCheck { get; set; }

        public DocumentItem(string path)
        {
            this.Name = Path.GetFileName(path);
            this.NameWithoutExtension = Path.GetFileNameWithoutExtension(path);
            this.Extension = Path.GetExtension(path);
            this.FilePath = path;
            this.DocumentType = Path.GetExtension(path).ToLower() switch
            {
                ".xls" => DocumentType.Excel,
                ".xlsx" => DocumentType.Excel,
                ".ppt" => DocumentType.PowerPoint,
                ".pptx" => DocumentType.PowerPoint,
                ".doc" => DocumentType.Word,
                ".docx" => DocumentType.Word,
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
