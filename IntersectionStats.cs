using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace RevitMEPHoleManager
{
    /// <summary>
    /// Одна строка для таблицы пересечений + координата центра.
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

        public double WidthFt { get; set; }   // DN-Ø или ширина лотка  (футы)
        public double HeightFt { get; set; }   // DN-Ø или высота лотка (футы)

        public double HoleWidthMm { get; set; }   // Итоговые габариты отверстия
        public double HoleHeightMm { get; set; }
        public string HoleTypeName { get; set; }

        // ▼ NEW для кластеризации --------------------------------------
        public XYZ Center { get; set; }   // координата центра, футы
        public bool IsMerged { get; set; }   // true, если это кластер
        public Guid ClusterId { get; set; }   // Id кластера
        public double CenterZft { get; set; }      // удобнее, чем XYZ Center
        public double? GapMm { get; set; }   // расстояние до соседа (< mergeDist) либо null
        // ---------------------------------------------------------------
    }

    /// <summary>
    /// Анализ пересечений + сводка для таблицы.
    /// </summary>
    internal static class IntersectionStats
    {
        /// <returns>(wRnd, wRec, fRnd, fRec, rows, hostStats)</returns>
        public static (int wRnd, int wRec, int fRnd, int fRec, List<IntersectRow> rows, List<HostStatRow> hostStats)
            Analyze(IEnumerable<Element> hosts,
                    IEnumerable<(Element elem, Transform tx)> mepList,
                    double clearanceMm)
        {
            int wRnd = 0, wRec = 0, fRnd = 0, fRec = 0;
            var rows = new List<IntersectRow>();
            var hostDict = new Dictionary<int, HostStatRow>();

            // пересекаются ли два bounding‑box'а
            bool Intersects(BoundingBoxXYZ a, BoundingBoxXYZ b, out XYZ ctr)
            {
                double minX = Math.Max(a.Min.X, b.Min.X);
                double minY = Math.Max(a.Min.Y, b.Min.Y);
                double minZ = Math.Max(a.Min.Z, b.Min.Z);
                double maxX = Math.Min(a.Max.X, b.Max.X);
                double maxY = Math.Min(a.Max.Y, b.Max.Y);
                double maxZ = Math.Min(a.Max.Z, b.Max.Z);
                bool hit = (minX <= maxX) && (minY <= maxY) && (minZ <= maxZ);
                ctr = hit ? new XYZ((minX + maxX) / 2,
                                    (minY + maxY) / 2,
                                    (minZ + maxZ) / 2) : null;
                return hit;
            }

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

                    // bb → координаты хоста
                    var bbHost = new BoundingBoxXYZ
                    {
                        Min = tx.OfPoint(bb.Min),
                        Max = tx.OfPoint(bb.Max)
                    };

                    if (!Intersects(hBox, bbHost, out XYZ center)) continue;

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

                    // сводка по конкретному хосту
                    if (!hostDict.TryGetValue(host.Id.IntegerValue, out var hs))
                    {
                        hs = new HostStatRow
                        {
                            HostId = host.Id.IntegerValue,
                            HostName = hostLbl
                        };
                        hostDict.Add(hs.HostId, hs);
                    }
                    if (isRound) hs.Round++; else hs.Rect++;

                    // размеры трассы
                    double elemW = 0, elemH = 0;
                    if (isRound && SizeHelper.TryGetRoundDiameter(mep, out double d))
                    {
                        elemW = elemH = d;
                    }
                    else if (!isRound && SizeHelper.TryGetRectSize(mep, out double w, out double h))
                    {
                        elemW = w; elemH = h;
                    }

                    // размеры в футах для MBR расчётов
                    double widthFt = UnitUtils.ConvertToInternalUnits(elemW, UnitTypeId.Millimeters);
                    double heightFt = UnitUtils.ConvertToInternalUnits(elemH, UnitTypeId.Millimeters);

                    // размеры отверстия
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

                        WidthFt = widthFt,
                        HeightFt = heightFt,

                        HoleWidthMm = holeW,
                        HoleHeightMm = holeH,
                        HoleTypeName = holeType,

                        Center = center,
                        CenterZft = center.Z,
                        IsMerged = false,
                        ClusterId = Guid.Empty
                    });
                }
            }

            return (wRnd, wRec, fRnd, fRec, rows, hostDict.Values.ToList());
        }
    }
}
