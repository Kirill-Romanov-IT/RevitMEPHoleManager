using System;
using System.Globalization;
using System.Windows.Data;
using Autodesk.Revit.DB;

namespace RevitMEPHoleManager
{
    /// <summary>
    /// Конвертер для отображения bool как "Да"/"Нет"
    /// </summary>
    public class BoolToYesNoConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value is bool b && b ? "Да" : "Нет";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value?.ToString() == "Да";
        }
    }

    /// <summary>
    /// Класс для представления прямоугольника отверстия
    /// </summary>
    public class HoleRectangle
    {
        public ElementId HostId { get; set; }
        public ElementId MepId { get; set; }
        public double CenterX { get; set; }  // в мм
        public double CenterY { get; set; }  // в мм  
        public double CenterZ { get; set; }  // в мм
        public double Width { get; set; }    // в мм
        public double Height { get; set; }   // в мм
        public double LeftEdge { get; set; } // в мм
        public double RightEdge { get; set; } // в мм
        public double BottomEdge { get; set; } // в мм
        public double TopEdge { get; set; } // в мм
    }

    /// <summary>
    /// Класс для результата пересечения
    /// </summary>
    public class IntersectionPoint
    {
        public double CenterX { get; set; }  // в мм
        public double CenterY { get; set; }  // в мм
        public double CenterZ { get; set; }  // в мм
    }
}
