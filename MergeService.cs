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
        private const double MmPerFt = 304.8;       // миллиметр на фут

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
                               .OrderBy(r => r.LocalCtr.X).ToList();

            log.Add($" Id   DN,мм   Xлокал");
            foreach (var p in xSort)
                log.Add($"{p.MepId,6}  {p.WidthLocFt / ftPerMm,5:F0}   {p.LocalCtr.X,8:F1}");

            double sumWidthFt = 0, sumGapXFt = 0;
            for (int i = 0; i < xSort.Count; i++)
            {
                sumWidthFt += xSort[i].WidthLocFt;                               // Ø/шир W
                if (i < xSort.Count - 1)
                {
                    double gap = (xSort[i + 1].LocalCtr.X - xSort[i + 1].WidthLocFt / 2) -
                                 (xSort[i].LocalCtr.X     + xSort[i].WidthLocFt / 2);
                    gap = Math.Max(0, gap);
                    sumGapXFt += gap;

                    log.Add($"gap-X {i}-{i + 1} = {gap * MmPerFt:F0} мм");
                }
            }

            // Итоговая формула габаритов отверстия - суммирование диаметров/ширин
            double holeWmm = sumWidthFt * MmPerFt + 2 * clearanceMm;
            
            log.Add($"Σ ширин = {sumWidthFt * MmPerFt:F0} мм  +  2×зазор = {2 * clearanceMm:F0} мм");

            /*────────────────  Y-направление  (высота) ────────────────*/
            var ySort = cluster.GroupBy(c => c.MepId).Select(g => g.First())
                               .OrderBy(r => r.LocalCtr.Y).ToList();

            double sumHeightFt = 0, sumGapYFt = 0;
            for (int i = 0; i < ySort.Count; i++)
            {
                sumHeightFt += ySort[i].HeightLocFt;                             // Ø/выс H
                if (i < ySort.Count - 1)
                {
                    double gap = (ySort[i + 1].LocalCtr.Y - ySort[i + 1].HeightLocFt / 2) -
                                 (ySort[i].LocalCtr.Y     + ySort[i].HeightLocFt / 2);
                    gap = Math.Max(0, gap);
                    sumGapYFt += gap;

                    log.Add($"gap-Y {i}-{i + 1} = {gap * MmPerFt:F0} мм");
                }
            }

            // Итоговая формула габаритов отверстия для Y - суммирование высот
            double holeHmm = sumHeightFt * MmPerFt + 2 * clearanceMm;
            
            log.Add($"Σ высот = {sumHeightFt * MmPerFt:F0} мм  +  2×зазор = {2 * clearanceMm:F0} мм");
            log.Add($"↦ Ширина отверстия = {holeWmm:F0} мм");
            log.Add($"↦ Высота отверстия = {holeHmm:F0} мм");
            log.HR();

            /*──────────────  центр отверстия  ─────────────*/
            // Центр кластера - среднее арифметическое всех центров
            double avgX = cluster.Average(r => r.Center.X);
            double avgY = cluster.Average(r => r.Center.Y);
            double avgZ = cluster.Average(r => r.Center.Z);

            var row = cluster[0];
            row.HoleWidthMm  = holeWmm;
            row.HoleHeightMm = holeHmm;
            row.HoleTypeName = SafeTypeName(holeWmm, holeHmm);
            row.GroupCtr     = new XYZ(avgX, avgY, avgZ);
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
