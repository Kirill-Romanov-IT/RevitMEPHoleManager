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

            // ❶ сортируем слева-на-право и убираем дубли по MepId
            var pipes = cluster.GroupBy(c => c.MepId)
                               .Select(g => g.First())
                               .OrderBy(r => r.Center.X)
                               .ToList();

            log.HR();
            log.Add($"Хост {hostId}   кластер №{clusterNum}");
            log.Add($" Id   DN,мм   Xцентр");
            foreach (var p in pipes)
                log.Add($"{p.MepId,6}  {p.WidthFt / ftPerMm,5:F0}   {p.Center.X,8:F1}");

            // ❷ считаем Ø-сумму и промежутки
            double sumDnFt = 0;
            double sumGapFt = 0;
            for (int i = 0; i < pipes.Count; i++)
            {
                sumDnFt += pipes[i].WidthFt;                 // Ø_i
            }

            for (int i = 0; i < pipes.Count - 1; i++)
            {
                var a = pipes[i];
                var b = pipes[i + 1];

                double dx = b.Center.X - a.Center.X;
                double dy = b.Center.Y - a.Center.Y;     // ← учитываем смещение по Y
                double d2D_Ft = Math.Sqrt(dx * dx + dy * dy);

                double gapFt = d2D_Ft - (a.WidthFt / 2) - (b.WidthFt / 2);
                gapFt = Math.Max(0, gapFt);              // отрицательные → 0

                sumGapFt += gapFt;

                double gapMm = UnitUtils.ConvertFromInternalUnits(gapFt, UnitTypeId.Millimeters);
                log.Add($"gap {i}-{i + 1} = {gapMm:F0} мм");
            }

            // ❸ переводим в мм
            double sumDnMm = sumDnFt / ftPerMm;
            double sumGapMm = sumGapFt / ftPerMm;

            // ❹ итоговая ширина (двойной clearance!)
            double holeWmm = sumDnMm + sumGapMm + 2 * clearanceMm;

            log.Add($"DN Σ  = {sumDnMm:F0} мм");
            log.Add($"Gap Σ = {sumGapMm:F0} мм");
            log.Add($"Clearance = {clearanceMm} мм");
            log.Add($"Ширина = {sumDnMm:F0} + {sumGapMm:F0} + 2×{clearanceMm} = {holeWmm:F0} мм");

            /* ➋  Высота – берём максимальную + зазор */
            double maxHeightFt = cluster.Max(r => r.HeightFt);
            double holeHmm = maxHeightFt / ftPerMm + clearanceMm;

            /* ➌  Центр отверстия – середина между первой и последней трубой */
            double left = pipes.First().Center.X - pipes.First().WidthFt / 2;
            double right = pipes.Last().Center.X + pipes.Last().WidthFt / 2;
            double ctrX = (left + right) / 2;
            double ctrY = cluster.Average(r => r.Center.Y);
            double ctrZ = cluster.Average(r => r.Center.Z);
            XYZ groupCtr = new XYZ(ctrX, ctrY, ctrZ);

            log.Add($"Центр X = {(ctrX / ftPerMm):F0} мм");

            /* ➍  Заполняем строку */
            var row = pipes[0];
            row.IsMerged = true;
            row.HoleWidthMm = holeWmm;
            row.HoleHeightMm = holeHmm;
            row.HoleTypeName = SafeTypeName(holeWmm, holeHmm);   // «350x200»
            row.GroupCtr = groupCtr;                         // ⬅ новый центр

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
