using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace RevitMEPHoleManager
{
    /// <summary>
    /// Класс для работы с геометрией отверстий и их объединением
    /// </summary>
    internal static class HoleGeometry
    {
        /// <summary>
        /// Представляет прямоугольное отверстие в 2D пространстве стены
        /// </summary>
        public class HoleRect
        {
            public int MepId { get; set; }
            public int HostId { get; set; }
            public XYZ Center { get; set; }          // 3D центр в мировых координатах
            public XYZ LocalCenter { get; set; }     // 2D центр в локальных координатах стены
            public double Width { get; set; }        // ширина в мм
            public double Height { get; set; }       // высота в мм
            public IntersectRow OriginalRow { get; set; }  // ссылка на исходную строку

            // Границы прямоугольника в локальных координатах стены
            public double MinX => LocalCenter.X - Width / 2.0 / 304.8;
            public double MaxX => LocalCenter.X + Width / 2.0 / 304.8;
            public double MinY => LocalCenter.Y - Height / 2.0 / 304.8;
            public double MaxY => LocalCenter.Y + Height / 2.0 / 304.8;

            /// <summary>
            /// Проверяет пересечение с другим отверстием
            /// </summary>
            public bool IntersectsWith(HoleRect other)
            {
                return !(MaxX < other.MinX || MinX > other.MaxX || 
                        MaxY < other.MinY || MinY > other.MaxY);
            }

            /// <summary>
            /// Вычисляет объединяющий прямоугольник (union) по крайним точкам всех отверстий
            /// </summary>
            public static HoleRect Union(IEnumerable<HoleRect> holes)
            {
                var holeList = holes.ToList();
                if (!holeList.Any()) return null;

                // Находим крайние точки ВСЕХ отверстий (не центры, а границы!)
                double minX = holeList.Min(h => h.MinX);  // самая левая граница
                double maxX = holeList.Max(h => h.MaxX);  // самая правая граница  
                double minY = holeList.Min(h => h.MinY);  // самая нижняя граница
                double maxY = holeList.Max(h => h.MaxY);  // самая верхняя граница

                var firstHole = holeList.First();
                
                // Центр объединенного отверстия - геометрический центр общего прямоугольника
                double unionCenterX = (minX + maxX) / 2;
                double unionCenterY = (minY + maxY) / 2;
                
                // Размеры объединенного отверстия - расстояние между крайними точками
                double unionWidthFt = maxX - minX;   // ширина в футах
                double unionHeightFt = maxY - minY;  // высота в футах
                
                return new HoleRect
                {
                    HostId = firstHole.HostId,
                    MepId = -1, // объединенное отверстие
                    LocalCenter = new XYZ(unionCenterX, unionCenterY, firstHole.LocalCenter.Z),
                    Center = new XYZ(
                        holeList.Average(h => h.Center.X),
                        holeList.Average(h => h.Center.Y),
                        holeList.Average(h => h.Center.Z)
                    ),
                    Width = unionWidthFt * 304.8,   // футы → мм
                    Height = unionHeightFt * 304.8, // футы → мм
                    OriginalRow = firstHole.OriginalRow
                };
            }
        }

        /// <summary>
        /// Новый алгоритм объединения отверстий по геометрическому пересечению
        /// </summary>
        /// <param name="rows">Исходные пересечения</param>
        /// <param name="clearanceMm">Зазор вокруг элементов</param>
        /// <param name="log">Логгер</param>
        /// <returns>Список окончательных отверстий (одиночных и объединенных)</returns>
        public static IEnumerable<IntersectRow> MergeByIntersection(
            IEnumerable<IntersectRow> rows, 
            double clearanceMm, 
            HoleLogger log)
        {
            if (rows == null) return Enumerable.Empty<IntersectRow>();

            var result = new List<IntersectRow>();
            log.Add("═══ НОВЫЙ АЛГОРИТМ ОБЪЕДИНЕНИЯ ОТВЕРСТИЙ ═══");

            foreach (var hostGroup in rows.GroupBy(r => r.HostId))
            {
                log.Add($"Host ID: {hostGroup.Key}");
                
                // Шаг 1: Создаем виртуальные отверстия для каждого пересечения
                var holes = hostGroup.Select(row => new HoleRect
                {
                    MepId = row.MepId,
                    HostId = row.HostId,
                    Center = row.Center,
                    LocalCenter = row.LocalCtr ?? row.Center, // используем локальные координаты если есть
                    Width = row.HoleWidthMm > 0 ? row.HoleWidthMm : row.ElemWidthMm + 2 * clearanceMm,
                    Height = row.HoleHeightMm > 0 ? row.HoleHeightMm : row.ElemHeightMm + 2 * clearanceMm,
                    OriginalRow = row
                }).ToList();

                log.Add($"Создано {holes.Count} виртуальных отверстий:");
                foreach (var hole in holes)
                {
                    log.Add($"  MEP {hole.MepId}: {hole.Width:F0}×{hole.Height:F0} границы[{hole.MinX * 304.8:F0}, {hole.MinY * 304.8:F0}] - [{hole.MaxX * 304.8:F0}, {hole.MaxY * 304.8:F0}]");
                }

                // Шаг 2: Находим группы пересекающихся отверстий
                var clusters = FindIntersectingClusters(holes, log);
                
                // Шаг 3: Создаем результирующие отверстия
                foreach (var cluster in clusters)
                {
                    if (cluster.Count == 1)
                    {
                        // Одиночное отверстие
                        result.Add(cluster[0].OriginalRow);
                        log.Add($"  Одиночное: MEP {cluster[0].MepId}, {cluster[0].Width:F0}×{cluster[0].Height:F0}");
                    }
                    else
                    {
                        // Объединенное отверстие
                        var unionHole = HoleRect.Union(cluster);
                        var mergedRow = CreateMergedRow(cluster, unionHole, log);
                        result.Add(mergedRow);
                        
                        log.Add($"  Объединенное: {cluster.Count} отверстий → {unionHole.Width:F0}×{unionHole.Height:F0}");
                        log.Add($"    Крайние точки: X[{cluster.Min(h => h.MinX * 304.8):F0}..{cluster.Max(h => h.MaxX * 304.8):F0}] Y[{cluster.Min(h => h.MinY * 304.8):F0}..{cluster.Max(h => h.MaxY * 304.8):F0}]");
                        
                        foreach (var hole in cluster)
                        {
                            log.Add($"    - MEP {hole.MepId}: {hole.Width:F0}×{hole.Height:F0} центр({hole.LocalCenter.X * 304.8:F0}, {hole.LocalCenter.Y * 304.8:F0})");
                        }
                    }
                }
                
                log.HR();
            }

            return result;
        }

        /// <summary>
        /// Находит группы пересекающихся отверстий (алгоритм Union-Find)
        /// </summary>
        private static List<List<HoleRect>> FindIntersectingClusters(List<HoleRect> holes, HoleLogger log)
        {
            var clusters = new List<List<HoleRect>>();
            var processed = new HashSet<HoleRect>();

            foreach (var hole in holes)
            {
                if (processed.Contains(hole)) continue;

                var cluster = new List<HoleRect>();
                var queue = new Queue<HoleRect>();
                queue.Enqueue(hole);
                processed.Add(hole);

                // BFS для поиска всех связанных отверстий
                while (queue.Count > 0)
                {
                    var current = queue.Dequeue();
                    cluster.Add(current);

                    // Проверяем пересечения с остальными отверстиями
                    foreach (var other in holes)
                    {
                        if (processed.Contains(other)) continue;

                        if (current.IntersectsWith(other))
                        {
                            queue.Enqueue(other);
                            processed.Add(other);
                            log.Add($"    Найдено пересечение: MEP {current.MepId} ∩ MEP {other.MepId}");
                        }
                    }
                }

                clusters.Add(cluster);
            }

            return clusters;
        }

        /// <summary>
        /// Создает объединенную IntersectRow из кластера отверстий
        /// </summary>
        private static IntersectRow CreateMergedRow(List<HoleRect> cluster, HoleRect unionHole, HoleLogger log)
        {
            var firstRow = cluster[0].OriginalRow;
            
            // Копируем первую строку и обновляем размеры
            var mergedRow = new IntersectRow
            {
                HostId = firstRow.HostId,
                MepId = firstRow.MepId, // ID первого элемента в кластере
                Host = firstRow.Host,
                Mep = $"Кластер ({cluster.Count})",
                Shape = "Cluster",
                WidthFt = firstRow.WidthFt,
                HeightFt = firstRow.HeightFt,
                Center = unionHole.Center,
                CenterXft = unionHole.Center.X,
                CenterYft = unionHole.Center.Y,
                CenterZft = unionHole.Center.Z,
                ElemWidthMm = firstRow.ElemWidthMm,
                ElemHeightMm = firstRow.ElemHeightMm,
                HoleWidthMm = unionHole.Width,
                HoleHeightMm = unionHole.Height,
                HoleTypeName = $"Прям. {Math.Ceiling(unionHole.Width)}×{Math.Ceiling(unionHole.Height)}",
                PipeDir = firstRow.PipeDir,
                LocalCtr = unionHole.LocalCenter,
                WidthLocFt = UnitUtils.ConvertToInternalUnits(unionHole.Width, UnitTypeId.Millimeters),
                HeightLocFt = UnitUtils.ConvertToInternalUnits(unionHole.Height, UnitTypeId.Millimeters),
                AxisLocal = firstRow.AxisLocal,
                IsDiagonal = firstRow.IsDiagonal,
                IsMerged = true,
                GroupCtr = unionHole.Center
            };

            return mergedRow;
        }
    }
}
