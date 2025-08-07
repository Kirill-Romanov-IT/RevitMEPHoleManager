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
        /// <returns>Новый список: одиночные + кластерные (IsMerged = true)</returns>
        public static IEnumerable<IntersectRow> Merge(IEnumerable<IntersectRow> rows, double maxGapMm, double clearanceMm)
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

                    // Создаём объединённую строку с улучшенным расчётом размеров
                    var merged = BuildMerged(cluster, clearanceMm);
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
        /// <returns>Объединённая строка IntersectRow</returns>
        private static IntersectRow BuildMerged(List<IntersectRow> cluster, double clearanceMm)
        {
            const double ftPerMm = 1 / 304.8;

            /* ➊  Сортируем по горизонтальной оси стены (X) */
            var byX = cluster.OrderBy(r => r.Center.X).ToList();

            double sumWidthFt = 0;
            double sumGapsFt = 0;

            for (int i = 0; i < byX.Count; i++)
            {
                double wFt = byX[i].WidthFt;           // DN или ширина лотка
                sumWidthFt += wFt;

                if (i < byX.Count - 1)
                {
                    double rightEdge = byX[i].Center.X + wFt / 2;
                    double leftEdge = byX[i + 1].Center.X - byX[i + 1].WidthFt / 2;
                    double gapFt = leftEdge - rightEdge;          // «чистый» зазор
                    sumGapsFt += Math.Max(gapFt, 0);
                }
            }

            /* ➋  Высота – берём максимальную + зазор */
            double maxHeightFt = cluster.Max(r => r.HeightFt);
            double holeHmm = maxHeightFt / ftPerMm + clearanceMm;

            /* ➌  Ширина – сумма Ø + зазоры + единый clearance */
            double holeWmm = (sumWidthFt + sumGapsFt) / ftPerMm + clearanceMm;

            /* ➍  Центр группы (по X) */
            double minX = byX.First().Center.X - byX.First().WidthFt / 2;
            double maxX = byX.Last().Center.X + byX.Last().WidthFt / 2;
            double ctrX = (minX + maxX) / 2;
            double ctrY = cluster.Average(r => r.Center.Y);
            double ctrZ = cluster.Average(r => r.Center.Z);
            XYZ groupCtr = new XYZ(ctrX, ctrY, ctrZ);

            /* ➎  Заполняем строку */
            var row = byX[0];
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
