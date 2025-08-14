using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace RevitMEPHoleManager
{
    /// <summary>
    /// Класс для управления объединением отверстий
    /// </summary>
    public static class HoleMergeManager
    {
        /// <summary>
        /// Анализирует и объединяет размещенные отверстия
        /// </summary>
        public static int AnalyzeAndMergeHoles(Document doc, Family holeFamily, double mergeThresholdMm, HoleLogger log)
        {
            if (holeFamily == null)
            {
                log.Add("❌ Семейство отверстий не выбрано");
                return 0;
            }

            log.Add($"Анализируем семейство: {holeFamily.Name}");

            // Собираем все размещенные отверстия этого семейства
            var placedHoles = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol?.Family?.Id == holeFamily.Id)
                .ToList();

            if (placedHoles.Count == 0)
            {
                log.Add("❌ Размещенные отверстия не найдены");
                return 0;
            }

            log.Add($"Найдено размещенных отверстий: {placedHoles.Count}");
            log.Add($"Порог объединения: {mergeThresholdMm:F0}мм");

            // Детальная информация об отверстиях
            foreach (var hole in placedHoles)
            {
                if (hole?.Host?.Id == null || hole.Id == null)
                {
                    log.Add($"⚠️ Пропуск отверстия с null Host или Id");
                    continue;
                }

                var pos = HoleSizeCalculator.GetLocalPosition(hole);
                var hostId = hole.Host?.Id.IntegerValue ?? -1;
                var width = HoleSizeCalculator.GetHoleWidth(hole);
                var height = HoleSizeCalculator.GetHoleHeight(hole);
                log.Add($"  Отверстие {hole.Id}: Host={hostId}, размер={width:F0}×{height:F0}, позиция=({pos?.X * 304.8:F0}, {pos?.Y * 304.8:F0}, {pos?.Z * 304.8:F0})");
            }

            // Группируем по хост-элементу (стена/плита)
            var hostGroups = placedHoles
                .Where(hole => hole.Host != null)
                .GroupBy(hole => hole.Host.Id.IntegerValue)
                .ToList();

            log.Add($"Хостов с отверстиями: {hostGroups.Count}");

            // Фильтруем только хосты с несколькими отверстиями
            var hostsWithMultipleHoles = hostGroups.Where(g => g.Count() > 1).ToList();
            log.Add($"Хостов с несколькими отверстиями: {hostsWithMultipleHoles.Count}");

            if (hostsWithMultipleHoles.Count == 0)
            {
                log.Add("ℹ️ Нет хостов с несколькими отверстиями для объединения");
                return 0;
            }

            int mergedClusters = 0;

            // Анализируем каждый хост отдельно
            foreach (var hostGroup in hostsWithMultipleHoles)
            {
                var hostElement = doc.GetElement(new ElementId(hostGroup.Key));
                if (hostElement == null) continue;

                log.Add($"Анализ хоста ID {hostGroup.Key}:");

                var holesOnHost = hostGroup.ToList();
                var clusters = FindIntersectingClusters(holesOnHost, mergeThresholdMm, log);

                foreach (var cluster in clusters)
                {
                    if (cluster.Count > 1)
                    {
                        try
                        {
                            var mergedHole = CreateMergedHole(doc, cluster, holeFamily, log);
                            if (mergedHole != null)
                            {
                                // Удаляем исходные отверстия
                                foreach (var originalHole in cluster)
                                {
                                    doc.Delete(originalHole.Id);
                                }
                                mergedClusters++;
                                log.Add($"✅ Создан объединенный кластер ({cluster.Count} отверстий)");
                            }
                        }
                        catch (Exception ex)
                        {
                            log.Add($"❌ Ошибка создания объединенного отверстия: {ex.Message}");
                        }
                    }
                }
            }

            return mergedClusters;
        }

        /// <summary>
        /// Находит кластеры пересекающихся отверстий
        /// </summary>
        private static List<List<FamilyInstance>> FindIntersectingClusters(List<FamilyInstance> holes, double mergeThresholdMm, HoleLogger log)
        {
            var clusters = new List<List<FamilyInstance>>();
            var processed = new HashSet<FamilyInstance>();

            for (int i = 0; i < holes.Count; i++)
            {
                var hole1 = holes[i];
                if (processed.Contains(hole1)) continue;

                var cluster = new List<FamilyInstance> { hole1 };
                processed.Add(hole1);

                // Ищем все отверстия, которые пересекаются с текущим кластером
                bool foundNew;
                do
                {
                    foundNew = false;
                    for (int j = 0; j < holes.Count; j++)
                    {
                        var hole2 = holes[j];
                        if (processed.Contains(hole2)) continue;

                        // Проверяем пересечение с любым отверстием в кластере
                        bool intersectsWithCluster = cluster.Any(clusterHole => 
                            HolesIntersect(clusterHole, hole2, mergeThresholdMm, log));

                        if (intersectsWithCluster)
                        {
                            cluster.Add(hole2);
                            processed.Add(hole2);
                            foundNew = true;
                        }
                    }
                } while (foundNew);

                clusters.Add(cluster);
            }

            return clusters;
        }

        /// <summary>
        /// Проверяет пересечение двух отверстий
        /// </summary>
        private static bool HolesIntersect(FamilyInstance hole1, FamilyInstance hole2, double mergeThresholdMm, HoleLogger log)
        {
            try
            {
                // Получаем размеры отверстий
                double w1 = HoleSizeCalculator.GetHoleWidth(hole1);
                double h1 = HoleSizeCalculator.GetHoleHeight(hole1);
                double w2 = HoleSizeCalculator.GetHoleWidth(hole2);
                double h2 = HoleSizeCalculator.GetHoleHeight(hole2);

                // Получаем позиции в локальных координатах стены
                XYZ pos1 = HoleSizeCalculator.GetLocalPosition(hole1);
                XYZ pos2 = HoleSizeCalculator.GetLocalPosition(hole2);

                log.Add($"    Проверка: {hole1.Id} vs {hole2.Id}");
                log.Add($"      Размеры: {w1:F0}×{h1:F0} vs {w2:F0}×{h2:F0}");
                log.Add($"      Позиции: ({pos1?.X * 304.8:F0}, {pos1?.Y * 304.8:F0}) vs ({pos2?.X * 304.8:F0}, {pos2?.Y * 304.8:F0})");

                if (pos1 == null || pos2 == null)
                {
                    log.Add($"      ❌ Не удалось получить позиции отверстий");
                    return false;
                }

                // Буфер для проверки пересечения (половина порога)
                double bufferFt = mergeThresholdMm / 2.0 / 304.8;
                log.Add($"      Буфер: {mergeThresholdMm / 2.0:F0}мм");

                // Расширенные границы отверстий (в футах)
                double minX1 = pos1.X - w1 / 2.0 / 304.8 - bufferFt;
                double maxX1 = pos1.X + w1 / 2.0 / 304.8 + bufferFt;
                double minY1 = pos1.Y - h1 / 2.0 / 304.8 - bufferFt;
                double maxY1 = pos1.Y + h1 / 2.0 / 304.8 + bufferFt;
                double minZ1 = pos1.Z - h1 / 2.0 / 304.8 - bufferFt;
                double maxZ1 = pos1.Z + h1 / 2.0 / 304.8 + bufferFt;

                double minX2 = pos2.X - w2 / 2.0 / 304.8 - bufferFt;
                double maxX2 = pos2.X + w2 / 2.0 / 304.8 + bufferFt;
                double minY2 = pos2.Y - h2 / 2.0 / 304.8 - bufferFt;
                double maxY2 = pos2.Y + h2 / 2.0 / 304.8 + bufferFt;
                double minZ2 = pos2.Z - h2 / 2.0 / 304.8 - bufferFt;
                double maxZ2 = pos2.Z + h2 / 2.0 / 304.8 + bufferFt;

                log.Add($"      Границы1: X[{minX1 * 304.8:F0}..{maxX1 * 304.8:F0}] Y[{minY1 * 304.8:F0}..{maxY1 * 304.8:F0}] Z[{minZ1 * 304.8:F0}..{maxZ1 * 304.8:F0}]");
                log.Add($"      Границы2: X[{minX2 * 304.8:F0}..{maxX2 * 304.8:F0}] Y[{minY2 * 304.8:F0}..{maxY2 * 304.8:F0}] Z[{minZ2 * 304.8:F0}..{maxZ2 * 304.8:F0}]");

                // Проверка пересечения расширенных AABB (включая Z координату)
                bool intersects = (maxX1 >= minX2 && minX1 <= maxX2) &&
                                (maxY1 >= minY2 && minY1 <= maxY2) &&
                                (maxZ1 >= minZ2 && minZ1 <= maxZ2);

                // Вычисляем 3D расстояние между центрами
                double distanceMm = Math.Sqrt(
                    Math.Pow((pos2.X - pos1.X) * 304.8, 2) +
                    Math.Pow((pos2.Y - pos1.Y) * 304.8, 2) +
                    Math.Pow((pos2.Z - pos1.Z) * 304.8, 2)
                );

                log.Add($"      3D расстояние между центрами: {distanceMm:F0}мм");

                bool shouldMerge = intersects && distanceMm < mergeThresholdMm;

                if (shouldMerge)
                {
                    log.Add($"      ✅ Объединение: {hole1.Id} ∩ {hole2.Id}, расстояние: {distanceMm:F0}мм < порог: {mergeThresholdMm:F0}мм");
                }
                else
                {
                    if (!intersects)
                    {
                        log.Add($"      ❌ Нет пересечения расширенных границ");
                    }
                    if (distanceMm >= mergeThresholdMm)
                    {
                        log.Add($"      ❌ Расстояние {distanceMm:F0}мм >= порога {mergeThresholdMm:F0}мм");
                    }
                }

                return shouldMerge;
            }
            catch (Exception ex)
            {
                log.Add($"      ❌ Ошибка проверки пересечения: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Создает объединенное отверстие для кластера
        /// </summary>
        private static FamilyInstance CreateMergedHole(Document doc, List<FamilyInstance> cluster, Family holeFamily, HoleLogger log)
        {
            if (cluster.Count == 0) return null;

            try
            {
                log.Add($"    ═══ АНАЛИЗ ГЕОМЕТРИИ СЕМЕЙСТВ ОТВЕРСТИЙ ═══");
                
                var familyBounds = new List<FamilyBounds>();
                
                foreach (var hole in cluster)
                {
                    var holeBounds = HoleGeometryAnalyzer.AnalyzeFamilyGeometry(hole, log);
                    if (holeBounds != null)
                    {
                        familyBounds.Add(holeBounds);
                    }
                    else
                    {
                        log.Add($"    ⚠️ Не удалось проанализировать геометрию {hole.Id}, используем fallback");
                        // Fallback к старому методу
                        var pos = HoleSizeCalculator.GetLocalPosition(hole);
                        if (pos != null)
                        {
                            double holeWidth = HoleSizeCalculator.GetHoleWidth(hole);
                            double holeHeight = HoleSizeCalculator.GetHoleHeight(hole);
                            
                            familyBounds.Add(new FamilyBounds
                            {
                                HoleId = hole.Id,
                                LeftMm = pos.X * 304.8 - holeWidth / 2.0,
                                RightMm = pos.X * 304.8 + holeWidth / 2.0,
                                BottomMm = pos.Y * 304.8 - holeHeight / 2.0,
                                TopMm = pos.Y * 304.8 + holeHeight / 2.0,
                                FrontMm = pos.Z * 304.8 - 50, // примерная глубина
                                BackMm = pos.Z * 304.8 + 50,
                                CenterPoint = pos,
                                Faces = new List<Face>()
                            });
                        }
                    }
                }
                
                if (familyBounds.Count == 0)
                {
                    log.Add($"    ❌ Не удалось получить границы ни одного отверстия");
                    return null;
                }
                
                // Сравниваем границы и определяем крайние грани
                HoleGeometryAnalyzer.CompareFamilyBounds(familyBounds, log);
                
                // Вычисляем общие границы всех отверстий
                double leftmostMm = familyBounds.Min(f => f.LeftMm);
                double rightmostMm = familyBounds.Max(f => f.RightMm);
                double bottomMm = familyBounds.Min(f => f.BottomMm);
                double topMm = familyBounds.Max(f => f.TopMm);
                
                // Получаем данные о позициях для расчета Z
                var positions = cluster.Select(HoleSizeCalculator.GetLocalPosition).Where(p => p != null).ToList();
                if (!positions.Any()) return null;

                double minZ = positions.Min(p => p.Z);
                double maxZ = positions.Max(p => p.Z);
                
                // Размеры объединенного отверстия = расстояние между крайними границами
                double mergedWidthMm = rightmostMm - leftmostMm;   // между левой и правой
                double mergedHeightMm = topMm - bottomMm;          // между низом и верхом  
                double mergedDepthMm = (maxZ - minZ) * 304.8 + 100; // глубина объединенного отверстия
                
                log.Add($"    ═══ ФИНАЛЬНЫЕ ГРАНИЦЫ НА ОСНОВЕ ГЕОМЕТРИИ ═══");
                log.Add($"    Левая граница:   {leftmostMm:F1}мм");
                log.Add($"    Правая граница:  {rightmostMm:F1}мм");
                log.Add($"    Нижняя граница:  {bottomMm:F1}мм"); 
                log.Add($"    Верхняя граница: {topMm:F1}мм");
                log.Add($"    ───────────────────────────────────");
                log.Add($"    РАЗМЕР ОХВАТЫВАЮЩЕГО ОТВЕРСТИЯ: {mergedWidthMm:F1}×{mergedHeightMm:F1}мм");
                
                // ФИНАЛЬНЫЕ РАЗМЕРЫ после всех проверок и коррекций
                log.Add($"    ═══ ФИНАЛЬНЫЕ РАЗМЕРЫ ═══");
                log.Add($"    Размер объединенного: {mergedWidthMm:F0}×{mergedHeightMm:F0}мм (глубина: {mergedDepthMm:F0}мм)");
                
                // Центр объединенного отверстия - центр охватывающего прямоугольника
                double centerXft = (leftmostMm + rightmostMm) / 2.0 / 304.8;  // центр между левой и правой границами
                double centerYft = (bottomMm + topMm) / 2.0 / 304.8;          // центр между нижней и верхней границами  
                double centerZft = (minZ + maxZ) / 2;                         // центр по глубине
                
                XYZ mergedCenter = new XYZ(centerXft, centerYft, centerZft);

                log.Add($"    Границы отверстий: X[{minZ * 304.8:F0}..{maxZ * 304.8:F0}] Y[{bottomMm:F0}..{topMm:F0}] Z[{minZ * 304.8:F0}..{maxZ * 304.8:F0}]");
                log.Add($"    Центр MBR отверстий: ({centerXft * 304.8:F0}, {centerYft * 304.8:F0}, {centerZft * 304.8:F0})");

                // Создаем типоразмер для объединенного отверстия ПОСЛЕ всех расчетов
                log.Add($"    ═══ СОЗДАНИЕ ТИПОРАЗМЕРА ═══");
                string typeName = $"Прям. {Math.Ceiling(mergedWidthMm)}×{Math.Ceiling(mergedHeightMm)}";
                log.Add($"    Имя типоразмера: '{typeName}'");
                var firstHole = cluster.First();
                var hostElement = firstHole.Host;

                // Находим или создаем типоразмер
                FamilySymbol mergedSymbol = holeFamily.GetFamilySymbolIds()
                    .Select(id => doc.GetElement(id) as FamilySymbol)
                    .FirstOrDefault(s => s.Name.Equals(typeName, StringComparison.OrdinalIgnoreCase));

                if (mergedSymbol == null)
                {
                    log.Add($"    Типоразмер '{typeName}' не найден, создаем новый");
                    mergedSymbol = firstHole.Symbol.Duplicate(typeName) as FamilySymbol;
                    HoleSizeCalculator.SetSize(mergedSymbol, mergedWidthMm, mergedHeightMm);
                    log.Add($"    ✅ Создан новый типоразмер с размерами {mergedWidthMm:F0}×{mergedHeightMm:F0}мм");
                }
                else
                {
                    log.Add($"    ✅ Найден существующий типоразмер '{typeName}'");
                }

                // Готовим базовый символ для размещения (используем первое отверстие)
                var baseSymbol = firstHole.Symbol;

                return FaceBasedPlacer.CreateMergedInstance(doc, hostElement, baseSymbol, mergedSymbol, mergedCenter, mergedDepthMm, log);
            }
            catch (Exception ex)
            {
                log.Add($"    Ошибка создания объединенного отверстия: {ex.Message}");
                return null;
            }
        }
    }
}
