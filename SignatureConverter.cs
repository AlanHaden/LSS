using System.Globalization;
using System;
using System.IO;
using System.Windows.Data;
using System.Windows.Media.Imaging;

namespace LSS
{
    public class SignatureConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is byte[] imageData && imageData.Length > 0)
            {
                return "Signature Saved";
            }

            return "None";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}

