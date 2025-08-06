using Autodesk.Revit.DB;

namespace RevitMEPHoleManager
{
    /// <summary>
    /// Вспомогательные методы для извлечения габаритов MEP‑элементов (мм).
    /// Поддерживает версии Revit 2020 – 2025.
    /// </summary>
    internal static class SizeHelper
    {
        // ────────────────────────────────
        //  Круглая геометрия (⌀)
        // ────────────────────────────────
        public static bool TryGetRoundDiameter(Element mep, out double diaMm)
        {
            diaMm = 0;
            if (mep == null) return false;

            int catId = mep.Category.Id.IntegerValue;

            // ── Трубы ─────────────────────
            if (catId == (int)BuiltInCategory.OST_PipeCurves)
            {
                Parameter p =
                    mep.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER) ??
                    mep.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);

                return ExtractMm(p, out diaMm);
            }

            // ── Воздуховоды круглого сечения ─
            if (catId == (int)BuiltInCategory.OST_DuctCurves)
            {
                if (mep.Document.GetElement(mep.GetTypeId()) is MEPCurveType ductType &&
                    ductType.Shape == ConnectorProfileType.Round)
                {
                    Parameter p = mep.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                    return ExtractMm(p, out diaMm);
                }
            }

            return false;
        }

        // ────────────────────────────────
        //  Прямоугольная геометрия (w × h)
        // ────────────────────────────────
        public static bool TryGetRectSize(Element mep, out double wMm, out double hMm)
        {
            wMm = hMm = 0;
            if (mep == null) return false;

            int catId = mep.Category.Id.IntegerValue;

            // ── Прямоугольные / овальные воздуховоды ─
            if (catId == (int)BuiltInCategory.OST_DuctCurves)
            {
                if (mep.Document.GetElement(mep.GetTypeId()) is MEPCurveType ductType &&
                    ductType.Shape != ConnectorProfileType.Round) // Rect / Oval
                {
                    Parameter pW = mep.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                    Parameter pH = mep.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                    return ExtractMm(pW, out wMm) & ExtractMm(pH, out hMm);
                }
            }

            // ── Кабель‑лотки (старые) ──────
            if (catId == (int)BuiltInCategory.OST_CableTray)
            {
                Parameter pW = mep.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM);
                Parameter pH = mep.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM);
                return ExtractMm(pW, out wMm) & ExtractMm(pH, out hMm);
            }

            // ── Кабель‑лотки «Run» (Revit 2022+) ─
            if (catId == (int)BuiltInCategory.OST_CableTrayRun)
            {
                Parameter pW = mep.get_Parameter(BuiltInParameter.RBS_CABLETRAYRUN_WIDTH_PARAM);
                Parameter pH = mep.get_Parameter(BuiltInParameter.RBS_CABLETRAYRUN_HEIGHT_PARAM);
                return ExtractMm(pW, out wMm) & ExtractMm(pH, out hMm);
            }

            return false;
        }

        // ────────────────────────────────
        //  Вспомогательный метод извлечения мм
        // ────────────────────────────────
        private static bool ExtractMm(Parameter p, out double valueMm)
        {
            valueMm = 0;
            if (p == null || !p.HasValue) return false;

            valueMm = UnitUtils.ConvertFromInternalUnits(p.AsDouble(), UnitTypeId.Millimeters);
            return valueMm > 0;
        }
    }
}