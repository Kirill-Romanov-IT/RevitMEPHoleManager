using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace RevitMEPHoleManager
{
    /// <summary>
    /// Сервис объединения соседних отверстий в один «кластер» при
    /// расстоянии между центрами меньше заданного порога.
    /// </summary>
    internal static class MergeService
    {
        private const double FtPerMm = 1 / 304.8;   // фут на миллиметр

        /// <param name="rows">Список найденных пересечений</param>
        /// <param name="maxGapMm">Максимальный зазор между центрами, мм</param>
        /// <returns>Новый список: одиночные + кластерные (IsMerged = true)</returns>
        public static IEnumerable<IntersectRow> Merge(IEnumerable<IntersectRow> rows, double maxGapMm)
        {
            if (rows == null) return Enumerable.Empty<IntersectRow>();
            if (maxGapMm <= 0) return rows;               // объединение отключено

            double gapFt = maxGapMm * FtPerMm;
            var result = new List<IntersectRow>();

            foreach (var hostGrp in rows.GroupBy(r => r.HostId))
            {
                var pending = hostGrp.ToList();

                while (pending.Any())
                {
                    //–– стартовая точка нового кластера
                    var seed = pending[0];
                    var cluster = new List<IntersectRow> { seed };
                    pending.RemoveAt(0);

                    bool expanded;
                    do
                    {
                        expanded = false;
                        // текущие XY‑границы кластера (по центрам)
                        double minX = cluster.Min(r => r.Center.X);
                        double maxX = cluster.Max(r => r.Center.X);
                        double minY = cluster.Min(r => r.Center.Y);
                        double maxY = cluster.Max(r => r.Center.Y);

                        foreach (var r in pending.ToList())
                        {
                            if (r.Center.X >= minX - gapFt && r.Center.X <= maxX + gapFt &&
                                r.Center.Y >= minY - gapFt && r.Center.Y <= maxY + gapFt)
                            {
                                cluster.Add(r);
                                pending.Remove(r);
                                expanded = true;
                            }
                        }
                    } while (expanded);

                    if (cluster.Count == 1)   // одиночное отверстие
                    {
                        result.Add(seed);
                        continue;
                    }

                    //–– строим минимальный прямоугольник по габаритам отверстий
                    double minXb = cluster.Min(r => r.Center.X - (r.HoleWidthMm * FtPerMm) / 2);
                    double maxXb = cluster.Max(r => r.Center.X + (r.HoleWidthMm * FtPerMm) / 2);
                    double minYb = cluster.Min(r => r.Center.Y - (r.HoleHeightMm * FtPerMm) / 2);
                    double maxYb = cluster.Max(r => r.Center.Y + (r.HoleHeightMm * FtPerMm) / 2);

                    double holeWmm = (maxXb - minXb) / FtPerMm;
                    double holeHmm = (maxYb - minYb) / FtPerMm;

                    static double Up5(double v) => Math.Ceiling(v / 5.0) * 5.0;
                    holeWmm = Up5(holeWmm);
                    holeHmm = Up5(holeHmm);

                    var merged = new IntersectRow
                    {
                        HostId = seed.HostId,
                        MepId = seed.MepId,
                        Host = seed.Host,
                        Mep = "Кластер",
                        Shape = "Прямоуг.",

                        ElemWidthMm = holeWmm,
                        ElemHeightMm = holeHmm,

                        HoleWidthMm = holeWmm,
                        HoleHeightMm = holeHmm,
                        HoleTypeName = $"Прям. {holeWmm}×{holeHmm}",

                        Center = new XYZ((minXb + maxXb) / 2, (minYb + maxYb) / 2, seed.Center.Z),
                        IsMerged = true,
                        ClusterId = Guid.NewGuid()
                    };

                    result.Add(merged);
                }
            }

            return result;
        }
    }
}
