using System;
using System.Diagnostics;
using Autodesk.Revit.DB;

namespace RevitMEPHoleManager
{
    /// <summary>
    /// Класс для расчета размеров отверстий на основе размеров семейств
    /// </summary>
    public static class HoleSizeCalculator
    {
        /// <summary>
        /// Получает ширину отверстия в мм
        /// </summary>
        public static double GetHoleWidth(FamilyInstance hole)
        {
            try
            {
                // Сначала пробуем параметры типоразмера (FamilySymbol)
                var symbol = hole.Symbol;
                if (symbol != null)
                {
                    var widthParam = symbol.LookupParameter("Ширина") ?? 
                                   symbol.LookupParameter("Width") ??
                                   symbol.get_Parameter(BuiltInParameter.FAMILY_WIDTH_PARAM);
                    
                    if (widthParam != null && widthParam.HasValue)
                    {
                        double widthValue = UnitUtils.ConvertFromInternalUnits(widthParam.AsDouble(), UnitTypeId.Millimeters);
                        Debug.WriteLine($"GetHoleWidth: Hole {hole.Id}, Symbol: {symbol.Name}, Width: {widthValue:F0}мм");
                        return widthValue;
                    }
                }
                
                // Затем пробуем параметры экземпляра
                var instanceWidthParam = hole.LookupParameter("Ширина") ?? 
                                       hole.LookupParameter("Width") ??
                                       hole.get_Parameter(BuiltInParameter.FAMILY_WIDTH_PARAM);
                
                if (instanceWidthParam != null && instanceWidthParam.HasValue)
                {
                    double widthValue = UnitUtils.ConvertFromInternalUnits(instanceWidthParam.AsDouble(), UnitTypeId.Millimeters);
                    Debug.WriteLine($"GetHoleWidth: Hole {hole.Id}, Instance param, Width: {widthValue:F0}мм");
                    return widthValue;
                }
                
                // Если ничего не найдено, пытаемся извлечь из имени типоразмера
                string typeName = symbol?.Name ?? "";
                System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(typeName, @"(\d+)×(\d+)");
                if (match.Success && double.TryParse(match.Groups[1].Value, out double widthFromName))
                {
                    Debug.WriteLine($"GetHoleWidth: Hole {hole.Id}, From name '{typeName}', Width: {widthFromName:F0}мм");
                    return widthFromName;
                }
                
                Debug.WriteLine($"GetHoleWidth ERROR: Hole {hole.Id}, Type: {typeName} - не удалось получить ширину");
                throw new InvalidOperationException($"Не удалось получить ширину отверстия {hole.Id}, тип: {typeName}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"GetHoleWidth EXCEPTION: Hole {hole.Id}, Error: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Получает высоту отверстия в мм
        /// </summary>
        public static double GetHoleHeight(FamilyInstance hole)
        {
            try
            {
                // Сначала пробуем параметры типоразмера (FamilySymbol)
                var symbol = hole.Symbol;
                if (symbol != null)
                {
                    var heightParam = symbol.LookupParameter("Высота") ?? 
                                    symbol.LookupParameter("Height") ??
                                    symbol.get_Parameter(BuiltInParameter.FAMILY_HEIGHT_PARAM);
                    
                    if (heightParam != null && heightParam.HasValue)
                    {
                        double heightValue = UnitUtils.ConvertFromInternalUnits(heightParam.AsDouble(), UnitTypeId.Millimeters);
                        Debug.WriteLine($"GetHoleHeight: Hole {hole.Id}, Symbol: {symbol.Name}, Height: {heightValue:F0}мм");
                        return heightValue;
                    }
                }
                
                // Затем пробуем параметры экземпляра
                var instanceHeightParam = hole.LookupParameter("Высота") ?? 
                                        hole.LookupParameter("Height") ??
                                        hole.get_Parameter(BuiltInParameter.FAMILY_HEIGHT_PARAM);
                
                if (instanceHeightParam != null && instanceHeightParam.HasValue)
                {
                    double heightValue = UnitUtils.ConvertFromInternalUnits(instanceHeightParam.AsDouble(), UnitTypeId.Millimeters);
                    Debug.WriteLine($"GetHoleHeight: Hole {hole.Id}, Instance param, Height: {heightValue:F0}мм");
                    return heightValue;
                }
                
                // Если ничего не найдено, пытаемся извлечь из имени типоразмера
                string typeName = symbol?.Name ?? "";
                System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(typeName, @"(\d+)×(\d+)");
                if (match.Success && double.TryParse(match.Groups[2].Value, out double heightFromName))
                {
                    Debug.WriteLine($"GetHoleHeight: Hole {hole.Id}, From name '{typeName}', Height: {heightFromName:F0}мм");
                    return heightFromName;
                }
                
                Debug.WriteLine($"GetHoleHeight ERROR: Hole {hole.Id}, Type: {typeName} - не удалось получить высоту");
                throw new InvalidOperationException($"Не удалось получить высоту отверстия {hole.Id}, тип: {typeName}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"GetHoleHeight EXCEPTION: Hole {hole.Id}, Error: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Получает локальную позицию отверстия относительно хоста
        /// </summary>
        public static XYZ GetLocalPosition(FamilyInstance hole)
        {
            if (hole == null)
            {
                System.Diagnostics.Debug.WriteLine("GetLocalPosition: hole is null");
                return null;
            }

            try
            {
                var location = hole.Location;
                if (location is LocationPoint locationPoint)
                {
                    return locationPoint.Point;
                }
                else if (location is LocationCurve locationCurve)
                {
                    // Для элементов с кривой локацией берем середину
                    var curve = locationCurve.Curve;
                    return curve.Evaluate(0.5, true); // параметр 0.5 = середина кривой
                }
                else
                {
                    // Fallback: центр bounding box
                    var bbox = hole.get_BoundingBox(null);
                    if (bbox != null)
                    {
                        return (bbox.Min + bbox.Max) / 2.0;
                    }
                }
                
                System.Diagnostics.Debug.WriteLine($"GetLocalPosition: не удалось получить позицию для {hole.Id}");
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetLocalPosition ERROR: Hole {hole.Id}, Error: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Универсальный метод для установки размеров семейства
        /// </summary>
        public static void SetSize(FamilySymbol familySymbol, double widthMm, double heightMm)
        {
            try
            {
                // Ищем параметры ширины
                var widthParam = FindByNames(familySymbol, "Ширина", "Width", "FAMILY_WIDTH_PARAM");
                if (widthParam != null && !widthParam.IsReadOnly)
                {
                    double widthFt = UnitUtils.ConvertToInternalUnits(widthMm, UnitTypeId.Millimeters);
                    widthParam.Set(widthFt);
                }
                
                // Ищем параметры высоты
                var heightParam = FindByNames(familySymbol, "Высота", "Height", "FAMILY_HEIGHT_PARAM");
                if (heightParam != null && !heightParam.IsReadOnly)
                {
                    double heightFt = UnitUtils.ConvertToInternalUnits(heightMm, UnitTypeId.Millimeters);
                    heightParam.Set(heightFt);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SetSize ERROR: {ex.Message}");
            }
        }

        /// <summary>
        /// Устанавливает параметр глубины для экземпляра семейства
        /// </summary>
        public static void SetDepthParam(FamilyInstance instance, double depthMm)
        {
            try
            {
                var depthParam = FindByNames(instance, "Глубина", "Depth", "Толщина", "Thickness");
                if (depthParam != null && !depthParam.IsReadOnly)
                {
                    double depthFt = UnitUtils.ConvertToInternalUnits(depthMm, UnitTypeId.Millimeters);
                    depthParam.Set(depthFt);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SetDepthParam ERROR: {ex.Message}");
            }
        }

        /// <summary>
        /// Поиск параметра по нескольким возможным именам
        /// </summary>
        private static Parameter FindByNames(Element element, params string[] names)
        {
            foreach (string name in names)
            {
                var param = element.LookupParameter(name);
                if (param != null) return param;
                
                // Также пробуем как BuiltInParameter если имя содержит PARAM
                if (name.Contains("PARAM"))
                {
                    try
                    {
                        if (Enum.TryParse<BuiltInParameter>(name, out BuiltInParameter builtIn))
                        {
                            param = element.get_Parameter(builtIn);
                            if (param != null) return param;
                        }
                    }
                    catch
                    {
                        // Игнорируем ошибки парсинга enum
                    }
                }
            }
            return null;
        }
    }
}
