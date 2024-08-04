using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace OfficeToPdf.Lib
{
    internal class DocumentTypeToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is DocumentType)
            {
                DocumentType type = (DocumentType)value;
                switch (type)
                {
                    case DocumentType.Excel:
                        return "#107C41";
                    case DocumentType.PowerPoint:
                        return "#C43E1C";
                    case DocumentType.Word:
                        return "#185ABD";
                    default:
                        return "#222222";
                }
            }
            return "#000000";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return new NotImplementedException();
        }
    }
}
