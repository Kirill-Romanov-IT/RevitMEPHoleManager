using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace RevitMEPHoleManager
{
    /// <summary>
    /// Одна строка для таблицы пересечений.
    /// </summary>
    internal class IntersectRow
    {
        public int HostId { get; set; }
        public int MepId { get; set; }

        public string Host { get; set; }   // Стена / Перекрытие
        public string Mep { get; set; }   // Труба / Воздуховод / Лоток
        public string Shape { get; set; }   // Круг / Прямоуг.

        public double ElemWidthMm { get; set; }   // Габариты трассы
        public double ElemHeightMm { get; set; }

        public double HoleWidthMm { get; set; }   // Итоговые габариты отверстия
        public double HoleHeightMm { get; set; }

        public string HoleTypeName { get; set; }   // Имя типоразмера (семейства)
    }

    /// <summary>
    /// Анализ пересечений + сводка для таблицы.
    /// </summary>
    internal static class IntersectionStats
    {
        /// <returns>(wRnd, wRec, fRnd, fRec, rows)</returns>
        public static (int wRnd, int wRec, int fRnd, int fRec, List<IntersectRow> rows)
            Analyze(IEnumerable<Element> hosts,
                    IEnumerable<(Element elem, Transform tx)> mepList,
                    double clearanceMm)
        {
            int wRnd = 0, wRec = 0, fRnd = 0, fRec = 0;
            var rows = new List<IntersectRow>();

            bool Intersects(BoundingBoxXYZ a, BoundingBoxXYZ b) =>
                !(a.Max.X < b.Min.X || a.Min.X > b.Max.X ||
                  a.Max.Y < b.Min.Y || a.Min.Y > b.Max.Y ||
                  a.Max.Z < b.Min.Z || a.Min.Z > b.Max.Z);

            foreach (Element host in hosts)
            {
                var hBox = host.get_BoundingBox(null);
                if (hBox == null) continue;

                bool isWall = host is Wall;
                bool isFloor = host is Floor;
                string hostLbl = isWall ? "Стена" : "Перекрытие";

                foreach ((Element mep, Transform tx) in mepList)
                {
                    var bb = mep.get_BoundingBox(null);
                    if (bb == null) continue;

                    // bb MEP → координаты хоста
                    var bbHost = new BoundingBoxXYZ
                    {
                        Min = tx.OfPoint(bb.Min),
                        Max = tx.OfPoint(bb.Max)
                    };

                    if (!Intersects(hBox, bbHost)) continue;

                    bool isRound = false;
                    string mepLbl, shapeLbl;

                    switch ((BuiltInCategory)mep.Category.Id.IntegerValue)
                    {
                        case BuiltInCategory.OST_PipeCurves:
                            isRound = true; mepLbl = "Труба"; break;

                        case BuiltInCategory.OST_DuctCurves:
                            var t = mep.Document.GetElement(mep.GetTypeId()) as MEPCurveType;
                            isRound = t?.Shape == ConnectorProfileType.Round;
                            mepLbl = "Воздуховод"; break;

                        case BuiltInCategory.OST_CableTray:
                            mepLbl = "Лоток"; break;

                        default:
                            mepLbl = "—"; break;
                    }

                    shapeLbl = isRound ? "Круг" : "Прямоуг.";

                    if (isWall) { if (isRound) wRnd++; else wRec++; }
                    if (isFloor) { if (isRound) fRnd++; else fRec++; }

                    // --- исходные размеры трассы ---
                    double elemW = 0, elemH = 0;
                    if (isRound && SizeHelper.TryGetRoundDiameter(mep, out double d))
                    {
                        elemW = elemH = d;
                    }
                    else if (!isRound &&
                             SizeHelper.TryGetRectSize(mep, out double w, out double h))
                    {
                        elemW = w; elemH = h;
                    }

                    // размеры готового отверстия
                    Calculaters.GetHoleSize(
                        isRound,
                        elemW,
                        elemH,
                        clearanceMm,
                        out double holeW,
                        out double holeH,
                        out string holeType);

                    rows.Add(new IntersectRow
                    {
                        HostId = host.Id.IntegerValue,
                        MepId = mep.Id.IntegerValue,

                        Host = hostLbl,
                        Mep = mepLbl,
                        Shape = shapeLbl,

                        ElemWidthMm = elemW,
                        ElemHeightMm = elemH,

                        HoleWidthMm = holeW,
                        HoleHeightMm = holeH,
                        HoleTypeName = holeType
                    });
                }
            }

            return (wRnd, wRec, fRnd, fRec, rows);
        }
    }
}
