using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace ExcelDropViewer
{
    public sealed class ResultToBackgroundConverter : IValueConverter
    {
        private static readonly SolidColorBrush NgBackground = new(Color.FromRgb(255, 228, 228));

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var text = System.Convert.ToString(value, culture)?.Trim();
            return string.Equals(text, "NG", StringComparison.OrdinalIgnoreCase)
                ? NgBackground
                : Brushes.Transparent;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}
