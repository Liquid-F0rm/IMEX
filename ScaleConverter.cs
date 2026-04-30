using System;
using System.Globalization;
using System.Windows.Data;

namespace WpfImageToText
{
    public class ScaleConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Exemple de conversion, à adapter selon vos besoins
            double scale = System.Convert.ToDouble(value);
            return scale * 2;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            double scale = System.Convert.ToDouble(value);
            return scale / 2;
        }
    }
}