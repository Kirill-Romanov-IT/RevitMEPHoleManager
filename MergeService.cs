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

            /* ── границы кластера (как было) ───────────────────────────────── */
            double minX = cluster.Min(r => r.Center.X - r.WidthFt / 2);
            double maxX = cluster.Max(r => r.Center.X + r.WidthFt / 2);
            double minY = cluster.Min(r => r.Center.Y - r.HeightFt / 2);
            double maxY = cluster.Max(r => r.Center.Y + r.HeightFt / 2);

            double holeWmm = (maxY - minY) / ftPerMm + clearanceMm;   // ширина (Y)
            double holeHmm = (maxX - minX) / ftPerMm + clearanceMm;   // высота  (X)

            /* ── GapMm уже вычислен до объединения ─────────────────────── */
            double? gapMm = cluster.FirstOrDefault(r => r.GapMm.HasValue)?.GapMm;

            /* ── Центр кластера (среднее геометрическое центров) ──────────── */
            double centerX = cluster.Average(r => r.Center.X);
            double centerY = cluster.Average(r => r.Center.Y);
            double centerZ = cluster.Average(r => r.Center.Z);

            /* ── собираем итоговую строку ─────────────────────────────────── */
            var row = cluster[0];
            return new IntersectRow
            {
                HostId = row.HostId,
                MepId = row.MepId,
                Host = row.Host,
                Mep = "Кластер",
                Shape = "Прямоуг.",

                ElemWidthMm = row.ElemWidthMm,  // сохраняем оригинальные для отображения
                ElemHeightMm = row.ElemHeightMm,

                WidthFt = (maxX - minX),   // общая ширина кластера в футах
                HeightFt = (maxY - minY),  // общая высота кластера в футах

                HoleWidthMm = holeWmm,
                HoleHeightMm = holeHmm,
                HoleTypeName = SafeTypeName(holeWmm, holeHmm),   // «350x200»

                Center = new XYZ(centerX, centerY, centerZ),
                CenterZft = centerZ,
                GapMm = gapMm,        // покажем в гриде
                IsMerged = true,
                ClusterId = Guid.NewGuid()
            };
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
