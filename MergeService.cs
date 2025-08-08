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
        /// <param name="clearanceMm">Зазор вокруг элемента, мм</param>
        /// <param name="log">Логгер для отслеживания процесса</param>
        /// <returns>Новый список: одиночные + кластерные (IsMerged = true)</returns>
        public static IEnumerable<IntersectRow> Merge(IEnumerable<IntersectRow> rows, double maxGapMm, double clearanceMm, HoleLogger log)
        {
            if (rows == null) return Enumerable.Empty<IntersectRow>();
            if (maxGapMm <= 0) return rows;               // объединение отключено

            double gapFt = maxGapMm * FtPerMm;
            var result = new List<IntersectRow>();
            int clusterN = 1;

            foreach (var hostGrp in rows.GroupBy(r => r.HostId))
            {
                // ➊ убираем дубли одной и той же трубы
                var uniq = hostGrp
                    .GroupBy(r => r.MepId)
                    .Select(g => g.First())        // берём первую строку трубы
                    .ToList();

                // дальше uniq вместо hostGrp.ToList();
                while (uniq.Any())
                {
                    //–– стартовая точка нового кластера
                    var seed = uniq[0];
                    var cluster = new List<IntersectRow> { seed };
                    uniq.RemoveAt(0);

                    bool expanded;
                    do
                    {
                        expanded = false;
                        // текущие XY‑границы кластера (по центрам)
                        double minX = cluster.Min(r => r.Center.X);
                        double maxX = cluster.Max(r => r.Center.X);
                        double minY = cluster.Min(r => r.Center.Y);
                        double maxY = cluster.Max(r => r.Center.Y);

                        foreach (var r in uniq.ToList())
                        {
                            if (r.Center.X >= minX - gapFt && r.Center.X <= maxX + gapFt &&
                                r.Center.Y >= minY - gapFt && r.Center.Y <= maxY + gapFt)
                            {
                                cluster.Add(r);
                                uniq.Remove(r);
                                expanded = true;
                            }
                        }
                    } while (expanded);

                    if (cluster.Count == 1)   // одиночное отверстие
                    {
                        result.Add(seed);
                        continue;
                    }

                    // Создаём объединённую строку с улучшенным расчётом размеров
                    var merged = BuildMerged(cluster, clearanceMm, log, clusterN++, seed.HostId);
                    result.Add(merged);
                }
            }

            return result;
        }

        /// <summary>
        /// Создаёт объединённую строку для кластера с улучшенным расчётом размеров.
        /// </summary>
        /// <param name="cluster">Список элементов кластера</param>
        /// <param name="clearanceMm">Зазор вокруг элемента, мм</param>
        /// <param name="log">Логгер для отслеживания процесса</param>
        /// <param name="clusterNum">Номер кластера</param>
        /// <param name="hostId">ID хост-элемента</param>
        /// <returns>Объединённая строка IntersectRow</returns>
        private static IntersectRow BuildMerged(List<IntersectRow> cluster, double clearanceMm, HoleLogger log, int clusterNum, int hostId)
        {
            const double ftPerMm = 1 / 304.8;

            log.HR();
            log.Add($"Хост {hostId}   кластер №{clusterNum}");

            /*────────────────  X-направление  (ширина)  ────────────────*/
            var xSort = cluster.GroupBy(c => c.MepId).Select(g => g.First())
                               .OrderBy(r => r.Center.X).ToList();

            log.Add($" Id   DN,мм   Xцентр");
            foreach (var p in xSort)
                log.Add($"{p.MepId,6}  {p.WidthFt / ftPerMm,5:F0}   {p.Center.X,8:F1}");

            double sumWidthFt = 0, sumGapXFt = 0;
            for (int i = 0; i < xSort.Count; i++)
            {
                sumWidthFt += xSort[i].WidthFt;                               // Ø/шир W
                if (i < xSort.Count - 1)
                {
                    double gap = (xSort[i + 1].Center.X - xSort[i + 1].WidthFt / 2) -
                                 (xSort[i].Center.X     + xSort[i].WidthFt / 2);
                    gap = Math.Max(0, gap);
                    sumGapXFt += gap;

                    log.Add($"gap-X {i}-{i + 1} = {gap / ftPerMm:F0} мм");
                }
            }

            double holeWmm = (sumWidthFt + sumGapXFt) / ftPerMm + 2 * clearanceMm;
            log.Add($"Σ W = {sumWidthFt / ftPerMm:F0}   Σ gap-X = {sumGapXFt / ftPerMm:F0}");

            /*────────────────  Y-направление  (высота) ────────────────*/
            var ySort = cluster.GroupBy(c => c.MepId).Select(g => g.First())
                               .OrderBy(r => r.Center.Y).ToList();

            double sumHeightFt = 0, sumGapYFt = 0;
            for (int i = 0; i < ySort.Count; i++)
            {
                sumHeightFt += ySort[i].HeightFt;                             // Ø/выс H
                if (i < ySort.Count - 1)
                {
                    double gap = (ySort[i + 1].Center.Y - ySort[i + 1].HeightFt / 2) -
                                 (ySort[i].Center.Y     + ySort[i].HeightFt / 2);
                    gap = Math.Max(0, gap);
                    sumGapYFt += gap;

                    log.Add($"gap-Y {i}-{i + 1} = {gap / ftPerMm:F0} мм");
                }
            }

            double holeHmm = (sumHeightFt + sumGapYFt) / ftPerMm + 2 * clearanceMm;
            log.Add($"Σ H = {sumHeightFt / ftPerMm:F0}   Σ gap-Y = {sumGapYFt / ftPerMm:F0}");
            log.Add($"↦ Ширина  = ΣW + Σgap-X + 2×clr = {holeWmm:F0} мм");
            log.Add($"↦ Высота  = ΣH + Σgap-Y + 2×clr = {holeHmm:F0} мм");
            log.HR();

            /*──────────────  центр отверстия  ─────────────*/
            double minX = xSort.First().Center.X - xSort.First().WidthFt  / 2;
            double maxX = xSort.Last() .Center.X + xSort.Last() .WidthFt  / 2;
            double minY = ySort.First().Center.Y - ySort.First().HeightFt / 2;
            double maxY = ySort.Last() .Center.Y + ySort.Last() .HeightFt / 2;

            var row = cluster[0];
            row.HoleWidthMm  = holeWmm;
            row.HoleHeightMm = holeHmm;
            row.HoleTypeName = SafeTypeName(holeWmm, holeHmm);
            row.GroupCtr     = new XYZ((minX + maxX) / 2, (minY + maxY) / 2, row.Center.Z);
            row.IsMerged     = true;
            return row;
        }



        /// <summary>
        /// Создаёт безопасное имя типоразмера для семейства
        /// </summary>
        private static string SafeTypeName(double wMm, double hMm)
        {
            // Округляем до целых и убираем спецсимволы
            int w = (int)Math.Ceiling(wMm);
            int h = (int)Math.Ceiling(hMm);
            return $"{w}x{h}";  // ASCII символы для совместимости
        }
    }
}
