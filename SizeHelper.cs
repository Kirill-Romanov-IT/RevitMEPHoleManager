using Autodesk.Revit.DB;

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
                    var ductType = mep.Document.GetElement(mep.GetTypeId()) as MEPCurveType;
                    if (ductType?.Shape == ConnectorProfileType.Round)
                    {
                        Parameter pD = mep.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                        if (pD == null || !pD.HasValue) return false;
                        wMm = hMm = UnitUtils.ConvertFromInternalUnits(pD.AsDouble(), UnitTypeId.Millimeters);
                        shape = ShapeKind.Round;
                    }
                    else // Rectangular or Oval
                    {
                        Parameter pW = mep.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                        Parameter pH = mep.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                        if (pW == null || pH == null || !pW.HasValue || !pH.HasValue) return false;
                        wMm = UnitUtils.ConvertFromInternalUnits(pW.AsDouble(), UnitTypeId.Millimeters);
                        hMm = UnitUtils.ConvertFromInternalUnits(pH.AsDouble(), UnitTypeId.Millimeters);
                        shape = wMm.Equals(hMm) ? ShapeKind.Square : ShapeKind.Rect;
                    }
                    return wMm > 0 && hMm > 0;

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
    }
}
