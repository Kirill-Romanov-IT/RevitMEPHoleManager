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
        public XYZ Center { get; set; }      // уже есть – центр конкретного MEP
        public XYZ GroupCtr { get; set; }     // ⬅ будет заполнен для «кустов»
        public bool IsMerged { get; set; }   // true, если это кластер
        public Guid ClusterId { get; set; }   // Id кластера
        public double CenterXft { get; set; }      // X-координата центра (футы)
        public double CenterYft { get; set; }      // Y-координата центра (футы)
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
            var hostMidPoint = new Dictionary<int, XYZ>();

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

                // сохраняем центр хоста для дедупликации
                XYZ hostCenter = new XYZ(
                    (hBox.Min.X + hBox.Max.X) / 2,
                    (hBox.Min.Y + hBox.Max.Y) / 2,
                    (hBox.Min.Z + hBox.Max.Z) / 2);
                hostMidPoint[host.Id.IntegerValue] = hostCenter;

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

                    if (!SizeHelper.TryGetSizes(mep, out double elemW, out double elemH, out ShapeKind kind))
                        continue; // элемент без размеров пропускаем

                    double wFt = UnitUtils.ConvertToInternalUnits(elemW, UnitTypeId.Millimeters);
                    double hFt = UnitUtils.ConvertToInternalUnits(elemH, UnitTypeId.Millimeters);

                    // размеры отверстия
                    Calculaters.GetHoleSize(
                        kind == ShapeKind.Round,
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
                        Mep = mep.Category.Name,
                        Shape = kind.ToString(),
                        WidthFt = wFt,
                        HeightFt = hFt,
                        Center = center,
                        CenterXft = center.X,
                        CenterYft = center.Y,
                        CenterZft = center.Z,
                        ElemWidthMm = elemW,
                        ElemHeightMm = elemH,
                        HoleWidthMm = holeW,
                        HoleHeightMm = holeH,
                        HoleTypeName = holeType
                    });
                }
            }

            // убираем повторы «та же труба – та же стена»
            rows = rows
                .GroupBy(r => new { r.HostId, r.MepId })  // ключ = грань-хост + труба
                .Select(g =>
                {
                    // берём пересечение, чья точка ближе к центру стены
                    return g.OrderBy(r => r.Center.DistanceTo(hostMidPoint[r.HostId]))
                            .First();
                })
                .ToList();

            return (wRnd, wRec, fRnd, fRec, rows, hostDict.Values.ToList());
        }
    }
}
