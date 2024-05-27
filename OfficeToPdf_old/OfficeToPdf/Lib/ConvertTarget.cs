using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace OfficeToPdf.Lib
{
    class ConvertTarget
    {
        public string SourcePath { get; set; }
        public string DestinationPath { get; set; }
        public string Parent { get; set; }
        public string FileName { get; set; }
        public string FileNameWithoutExtension { get; set; }
        public string Extension { get; set; }

        public OfficeAppType AppType { get; set; }

        public ConvertTarget(string path)
        {
            this.SourcePath = path;

            this.Parent = Path.GetDirectoryName(path);
            this.FileName = Path.GetFileName(path);
            this.FileNameWithoutExtension = Path.GetFileNameWithoutExtension(path);
            this.Extension = Path.GetExtension(path);

            this.DestinationPath = Path.Combine(this.Parent, this.FileNameWithoutExtension + ".pdf");

            this.AppType = Extension switch
            {
                ".xls" or ".xlsx" => OfficeAppType.Excel,
                ".doc" or ".docx" => OfficeAppType.Word,
                ".ppt" or ".pptx" => OfficeAppType.PowerPoint,
                _ => OfficeAppType.None,
            };
        }
    }
}
