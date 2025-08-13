using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;

namespace RevitMEPHoleManager
{
    /// <summary>
    /// Тип сечения инженерной трассы
    /// </summary>
    internal enum ShapeKind
    {
        Unknown,
        Round,
        Rect,
        Square,
        Tray
    }

    /// <summary>
    /// Вспомогательные методы для извлечения габаритов MEP‑элементов (мм).
    /// Поддерживает версии Revit 2020 – 2025.
    /// </summary>
    internal static class SizeHelper
    {
        /// <summary>
        /// Возвращает true, если размер получен, и заполняет W/H (мм) + тип сечения
        /// </summary>
        public static bool TryGetSizes(Element mep,
                                       out double wMm,
                                       out double hMm,
                                       out ShapeKind shape)
        {
            wMm = hMm = 0;
            shape = ShapeKind.Unknown;
            if (mep == null) return false;

            switch (mep.Category.Id.IntegerValue)
            {
                /* 1. трубы  ──────────────────────────────────── */
                case (int)BuiltInCategory.OST_PipeCurves:
                    Parameter pDN = mep.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                    if (pDN == null || !pDN.HasValue) return false;

                    double dnMm = UnitUtils.ConvertFromInternalUnits(pDN.AsDouble(), UnitTypeId.Millimeters);
                    wMm = hMm = dnMm;
                    shape = ShapeKind.Round;
                    return dnMm > 0;

                /* 2. воздуховоды ─────────────────────────────── */
                case (int)BuiltInCategory.OST_DuctCurves:
                    return TryGetDuctSizes(mep, out wMm, out hMm, out shape);

                /* 3. кабель-лотки ─────────────────────────────── */
                case (int)BuiltInCategory.OST_CableTray:
                case (int)BuiltInCategory.OST_CableTrayRun: // Also handle Cable Tray Runs
                    Parameter cw = mep.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM) ?? mep.get_Parameter(BuiltInParameter.RBS_CABLETRAYRUN_WIDTH_PARAM);
                    Parameter ch = mep.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM) ?? mep.get_Parameter(BuiltInParameter.RBS_CABLETRAYRUN_HEIGHT_PARAM);
                    if (cw == null || ch == null || !cw.HasValue || !ch.HasValue) return false;
                    wMm = UnitUtils.ConvertFromInternalUnits(cw.AsDouble(), UnitTypeId.Millimeters);
                    hMm = UnitUtils.ConvertFromInternalUnits(ch.AsDouble(), UnitTypeId.Millimeters);
                    shape = ShapeKind.Tray;
                    return wMm > 0 && hMm > 0;
            }
            return false;
        }

        /// <summary>
        /// Улучшенный метод для получения размеров воздуховодов с множественными fallback'ами
        /// </summary>
        private static bool TryGetDuctSizes(Element mep, out double wMm, out double hMm, out ShapeKind shape)
        {
            wMm = hMm = 0;
            shape = ShapeKind.Unknown;

            try
            {
                // ══════════════════════════════════════════════════════════════════
                // МЕТОД 1: Через тип воздуховода (самый надежный)
                // ══════════════════════════════════════════════════════════════════
                var ductType = mep.Document.GetElement(mep.GetTypeId()) as MEPCurveType;
                if (ductType != null)
                {
                    // Круглые воздуховоды
                    if (ductType.Shape == ConnectorProfileType.Round)
                    {
                        Parameter pD = mep.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                        if (pD != null && pD.HasValue && pD.AsDouble() > 0)
                        {
                            double diamMm = UnitUtils.ConvertFromInternalUnits(pD.AsDouble(), UnitTypeId.Millimeters);
                            wMm = hMm = diamMm;
                            shape = ShapeKind.Round;
                            return true;
                        }
                    }
                    // Прямоугольные и овальные воздуховоды
                    else if (ductType.Shape == ConnectorProfileType.Rectangular || 
                             ductType.Shape == ConnectorProfileType.Oval)
                    {
                        Parameter pW = mep.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                        Parameter pH = mep.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                        
                        if (pW != null && pH != null && pW.HasValue && pH.HasValue && 
                            pW.AsDouble() > 0 && pH.AsDouble() > 0)
                        {
                            wMm = UnitUtils.ConvertFromInternalUnits(pW.AsDouble(), UnitTypeId.Millimeters);
                            hMm = UnitUtils.ConvertFromInternalUnits(pH.AsDouble(), UnitTypeId.Millimeters);
                            
                            // Для овальных считаем как прямоугольные (безопаснее)
                            if (ductType.Shape == ConnectorProfileType.Oval)
                            {
                                shape = ShapeKind.Rect;
                            }
                            else
                            {
                                shape = Math.Abs(wMm - hMm) < 1.0 ? ShapeKind.Square : ShapeKind.Rect;
                            }
                            
                            // ДИАГНОСТИКА
                            System.Diagnostics.Debug.WriteLine($"TryGetDuctSizes (МЕТОД 1): W={wMm:F0}, H={hMm:F0}, Shape={shape}");
                            return true;
                        }
                    }
                }

                // ══════════════════════════════════════════════════════════════════
                // МЕТОД 2: Прямое чтение параметров экземпляра (fallback)
                // ══════════════════════════════════════════════════════════════════
                
                // Попытка получить диаметр для круглого
                Parameter dirDiam = mep.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                if (dirDiam != null && dirDiam.HasValue && dirDiam.AsDouble() > 0)
                {
                    double diamMm = UnitUtils.ConvertFromInternalUnits(dirDiam.AsDouble(), UnitTypeId.Millimeters);
                    wMm = hMm = diamMm;
                    shape = ShapeKind.Round;
                    return true;
                }

                // Попытка получить ширину и высоту
                Parameter dirW = mep.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                Parameter dirH = mep.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                
                if (dirW != null && dirH != null && dirW.HasValue && dirH.HasValue && 
                    dirW.AsDouble() > 0 && dirH.AsDouble() > 0)
                {
                    wMm = UnitUtils.ConvertFromInternalUnits(dirW.AsDouble(), UnitTypeId.Millimeters);
                    hMm = UnitUtils.ConvertFromInternalUnits(dirH.AsDouble(), UnitTypeId.Millimeters);
                    shape = Math.Abs(wMm - hMm) < 1.0 ? ShapeKind.Square : ShapeKind.Rect;
                    return true;
                }

                // ══════════════════════════════════════════════════════════════════
                // МЕТОД 3: Через коннекторы воздуховода (последний резерв)
                // ══════════════════════════════════════════════════════════════════
                if (mep is Duct duct)
                {
                    var connectorSet = duct.ConnectorManager?.Connectors;
                    if (connectorSet != null)
                    {
                        foreach (Connector conn in connectorSet)
                        {
                            if (conn.Shape == ConnectorProfileType.Round && conn.Radius > 0)
                            {
                                double radiusMm = UnitUtils.ConvertFromInternalUnits(conn.Radius, UnitTypeId.Millimeters);
                                wMm = hMm = radiusMm * 2;
                                shape = ShapeKind.Round;
                                return true;
                            }
                            else if ((conn.Shape == ConnectorProfileType.Rectangular || 
                                     conn.Shape == ConnectorProfileType.Oval) && 
                                     conn.Width > 0 && conn.Height > 0)
                            {
                                wMm = UnitUtils.ConvertFromInternalUnits(conn.Width, UnitTypeId.Millimeters);
                                hMm = UnitUtils.ConvertFromInternalUnits(conn.Height, UnitTypeId.Millimeters);
                                shape = Math.Abs(wMm - hMm) < 1.0 ? ShapeKind.Square : ShapeKind.Rect;
                                return true;
                            }
                        }
                    }
                }

                // ══════════════════════════════════════════════════════════════════
                // МЕТОД 4: Анализ геометрии (крайний случай)
                // ══════════════════════════════════════════════════════════════════
                var bbox = mep.get_BoundingBox(null);
                if (bbox != null)
                {
                    double bboxW = UnitUtils.ConvertFromInternalUnits(bbox.Max.X - bbox.Min.X, UnitTypeId.Millimeters);
                    double bboxH = UnitUtils.ConvertFromInternalUnits(bbox.Max.Y - bbox.Min.Y, UnitTypeId.Millimeters);
                    
                    // Используем меньшие размеры (исключаем длину)
                    if (bboxW > 50 && bboxH > 50) // минимальные разумные размеры
                    {
                        wMm = Math.Min(bboxW, bboxH);
                        hMm = Math.Max(bboxW, bboxH);
                        shape = Math.Abs(wMm - hMm) < 50 ? ShapeKind.Square : ShapeKind.Rect;
                        return true;
                    }
                }
            }
            catch (System.Exception)
            {
                // Подавляем ошибки и возвращаем false
            }

            return false;
        }

        /// <summary>
        /// Диагностический метод для получения информации о воздуховоде (для отладки)
        /// </summary>
        public static string GetDuctDiagnostics(Element mep)
        {
            if (mep?.Category?.Id?.IntegerValue != (int)BuiltInCategory.OST_DuctCurves)
                return "Не воздуховод";

            var diagnostics = new System.Text.StringBuilder();
            diagnostics.AppendLine($"Воздуховод ID: {mep.Id.IntegerValue}");

            try
            {
                // Информация о типе
                var ductType = mep.Document.GetElement(mep.GetTypeId()) as MEPCurveType;
                if (ductType != null)
                {
                    diagnostics.AppendLine($"Тип: {ductType.Name}");
                    diagnostics.AppendLine($"Форма: {ductType.Shape}");
                }
                else
                {
                    diagnostics.AppendLine("Тип: NULL");
                }

                // Проверяем параметры
                var pDiam = mep.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                if (pDiam != null)
                {
                    diagnostics.AppendLine($"Диаметр: {(pDiam.HasValue ? $"{UnitUtils.ConvertFromInternalUnits(pDiam.AsDouble(), UnitTypeId.Millimeters):F0} мм" : "Нет значения")}");
                }

                var pWidth = mep.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                var pHeight = mep.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                
                if (pWidth != null)
                {
                    diagnostics.AppendLine($"Ширина: {(pWidth.HasValue ? $"{UnitUtils.ConvertFromInternalUnits(pWidth.AsDouble(), UnitTypeId.Millimeters):F0} мм" : "Нет значения")}");
                }
                if (pHeight != null)
                {
                    diagnostics.AppendLine($"Высота: {(pHeight.HasValue ? $"{UnitUtils.ConvertFromInternalUnits(pHeight.AsDouble(), UnitTypeId.Millimeters):F0} мм" : "Нет значения")}");
                }

                // Информация о коннекторах
                if (mep is Duct duct && duct.ConnectorManager?.Connectors != null)
                {
                    int connCount = 0;
                    foreach (Connector conn in duct.ConnectorManager.Connectors)
                    {
                        connCount++;
                        diagnostics.AppendLine($"Коннектор {connCount}: {conn.Shape}, R={UnitUtils.ConvertFromInternalUnits(conn.Radius, UnitTypeId.Millimeters):F0}мм, W×H={UnitUtils.ConvertFromInternalUnits(conn.Width, UnitTypeId.Millimeters):F0}×{UnitUtils.ConvertFromInternalUnits(conn.Height, UnitTypeId.Millimeters):F0}мм");
                    }
                }
            }
            catch (System.Exception ex)
            {
                diagnostics.AppendLine($"Ошибка: {ex.Message}");
            }

            return diagnostics.ToString();
        }
    }
}
