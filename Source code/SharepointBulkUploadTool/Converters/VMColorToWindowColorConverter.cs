using SharepointBulkUploadTool.ViewModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Media;

namespace SharepointBulkUploadTool.Converters
{
    public class VMColorToWindowColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            VMColor color =  (VMColor) value;

            switch (color)
            {
                case VMColor.Red:
                    return Brushes.Red;
                case VMColor.Black:
                    return Brushes.Black;
                case VMColor.Green:
                    return Brushes.Green;
            }

            return Binding.DoNothing;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
