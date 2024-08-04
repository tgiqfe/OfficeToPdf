using MaterialDesignThemes.Wpf;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace OfficeToPdf.Lib
{
    internal class DocumentTypeToIconConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is DocumentType type)
            {
                return type switch
                {
                    DocumentType.Excel => PackIconKind.FileExcel,
                    DocumentType.PowerPoint => PackIconKind.FilePowerpoint,
                    DocumentType.Word => PackIconKind.FileWord,
                    _ => PackIconKind.FileQuestion,
                };
            }
            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return new NotImplementedException();
        }
    }
}
