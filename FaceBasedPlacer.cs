using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace RevitMEPHoleManager
{
    /// <summary>
    /// Класс для размещения семейств на гранях хостов
    /// </summary>
    public static class FaceBasedPlacer
    {
        /// <summary>
        /// Создает объединенный экземпляр семейства
        /// </summary>
        public static FamilyInstance CreateMergedInstance(Document doc, Element hostElement, FamilySymbol baseSymbol, 
            FamilySymbol mergedSymbol, XYZ mergedCenter, double mergedDepthMm, HoleLogger log)
        {
            try
            {
                // Получаем ПРАВИЛЬНОЕ направление для поиска грани
                XYZ pickDir;
                XYZ refDirection = XYZ.BasisX; // для размещения семейства

                if (hostElement is Wall wall)
                {
                    pickDir = wall.Orientation.Normalize();
                    log.Add($"    Анализ стены: pickDir.Z={pickDir.Z:F3}, pickDir.Y={pickDir.Y:F3}, pickDir.X={pickDir.X:F3}");
                    
                    // Для стен refDirection должен быть перпендикулярен к pickDir
                    if (Math.Abs(pickDir.DotProduct(XYZ.BasisX)) < 0.9)
                        refDirection = XYZ.BasisX;
                    else
                        refDirection = XYZ.BasisZ;
                    
                    log.Add($"    Выбрано refDirection для стены: ({refDirection.X:F3}, {refDirection.Y:F3}, {refDirection.Z:F3})");
                }
                else
                {
                    // Для плит
                    pickDir = XYZ.BasisZ;
                    refDirection = XYZ.BasisX;
                    log.Add($"    Анализ плиты: pickDir=BasisZ, refDirection=BasisX");
                }

                log.Add($"    Направление поиска: ({pickDir.X:F3}, {pickDir.Y:F3}, {pickDir.Z:F3})");
                log.Add($"    Направление размещения: ({refDirection.X:F3}, {refDirection.Y:F3}, {refDirection.Z:F3})");

                // Проверяем, что направления не параллельны
                double dot = Math.Abs(pickDir.DotProduct(refDirection));
                log.Add($"    Угол между направлениями (dot): {dot:F3} (должен быть < 0.9)");

                if (dot > 0.9)
                {
                    // Принудительно делаем refDirection перпендикулярным
                    if (Math.Abs(pickDir.DotProduct(XYZ.BasisY)) < 0.9)
                        refDirection = XYZ.BasisY;
                    else if (Math.Abs(pickDir.DotProduct(XYZ.BasisZ)) < 0.9)
                        refDirection = XYZ.BasisZ;
                    else
                        refDirection = XYZ.BasisX;
                    
                    log.Add($"    Скорректировано refDirection: ({refDirection.X:F3}, {refDirection.Y:F3}, {refDirection.Z:F3})");
                }

                // Ищем подходящую грань
                Reference faceRef = null;
                try
                {
                    faceRef = PickHostFace(doc, hostElement, mergedCenter, pickDir);
                    log.Add($"    Грань через PickHostFace: {faceRef != null}");
                }
                catch (Exception ex)
                {
                    log.Add($"    Ошибка поиска грани: {ex.Message}");
                }

                if (faceRef == null)
                {
                    log.Add($"    ⚠️ Грань не найдена, пытаемся создать host-based отверстие");
                    
                    // Последняя попытка - создаем отверстие как host-based
                    try
                    {
                        // ➊ Создаем host-based отверстие с базовым символом
                        log.Add($"    Создание host-based с базовым символом: {baseSymbol.Name}");
                        var hostBasedInstance = doc.Create.NewFamilyInstance(
                            mergedCenter, baseSymbol, hostElement, 
                            Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                        
                        // ➋ Переключаем на нужный типоразмер
                        log.Add($"    Переключение на типоразмер: {mergedSymbol.Name}");
                        if (!mergedSymbol.IsActive) mergedSymbol.Activate();
                        hostBasedInstance.ChangeTypeId(mergedSymbol.Id);
                        
                        // ➌ Устанавливаем глубину объединенного отверстия
                        HoleSizeCalculator.SetDepthParam(hostBasedInstance, mergedDepthMm);
                        
                        log.Add($"    ✅ Создано host-based отверстие");
                        log.Add($"    Установлена глубина: {mergedDepthMm:F0}мм");
                        return hostBasedInstance;
                    }
                    catch (Exception ex)
                    {
                        log.Add($"    ❌ Ошибка создания host-based отверстия: {ex.Message}");
                        return null;
                    }
                }

                // Проектируем центр на найденную грань
                XYZ projectedCenter = mergedCenter;
                try
                {
                    Face face = hostElement.GetGeometryObjectFromReference(faceRef) as Face;
                    if (face != null)
                    {
                        var projection = face.Project(mergedCenter);
                        if (projection != null)
                        {
                            projectedCenter = projection.XYZPoint;
                            log.Add($"    Центр спроектирован на грань: ({projectedCenter.X * 304.8:F0}, {projectedCenter.Y * 304.8:F0}, {projectedCenter.Z * 304.8:F0})");
                        }
                    }
                }
                catch (Exception ex)
                {
                    log.Add($"    Предупреждение проекции: {ex.Message}");
                }

                // ➊ Создаем face-based экземпляр с БАЗОВЫМ символом
                log.Add($"    Создание отверстия с базовым символом: {baseSymbol.Name}");
                var mergedInstance = doc.Create.NewFamilyInstance(faceRef, projectedCenter, refDirection, baseSymbol);
                
                // ➋ Переключаем на нужный типоразмер
                log.Add($"    Переключение на типоразмер: {mergedSymbol.Name}");
                if (!mergedSymbol.IsActive) mergedSymbol.Activate();
                mergedInstance.ChangeTypeId(mergedSymbol.Id);
                
                // ➌ Устанавливаем глубину объединенного отверстия
                HoleSizeCalculator.SetDepthParam(mergedInstance, mergedDepthMm);
                log.Add($"    Установлена глубина: {mergedDepthMm:F0}мм");

                return mergedInstance;
            }
            catch (Exception ex)
            {
                log.Add($"    Ошибка создания объединенного отверстия: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Находит подходящую грань хоста для размещения семейства
        /// </summary>
        public static Reference PickHostFace(Document doc, Element host, XYZ pt, XYZ pipeDir = null)
        {
            IEnumerable<Reference> refs;

            if (host is Floor floor)
            {
                // все верхние + нижние грани плиты
                refs = HostObjectUtils.GetTopFaces(floor)
                         .Concat(HostObjectUtils.GetBottomFaces(floor));
            }
            else if (host is Wall wall)
            {
                // внешние + внутренние грани стены
                refs = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior)
                         .Concat(HostObjectUtils.GetSideFaces(wall, ShellLayerType.Interior));
            }
            else
                throw new InvalidOperationException($"Неподдерживаемый тип хоста: {host.GetType().Name}");

            // проверяем, что у нас есть грани для работы
            if (!refs.Any())
            {
                throw new InvalidOperationException($"Не удалось получить грани для элемента {host.Id}");
            }

            var pipe = (pipeDir ?? XYZ.BasisZ).Normalize();

            var best = refs
                .Select<Reference, (Reference faceRef, double score, double dist)>(r =>
                {
                    Face face = null;
                    try
                    {
                        face = host.GetGeometryObjectFromReference(r) as Face;
                    }
                    catch
                    {
                    }

                    if (face == null) return (r, double.NegativeInfinity, double.PositiveInfinity);

                    XYZ n;
                    try
                    {
                        if (face is PlanarFace pf)
                        {
                            n = pf.FaceNormal.Normalize();
                        }
                        else if (face is CylindricalFace cf)
                        {
                            var projection = cf.Project(pt);
                            if (projection != null)
                            {
                                var uv = projection.UVPoint;
                                n = cf.ComputeNormal(uv).Normalize();
                            }
                            else
                            {
                                // Fallback: нормаль в центре грани
                                var box = cf.GetBoundingBox();
                                var centerUV = new UV((box.Min.U + box.Max.U) / 2, (box.Min.V + box.Max.V) / 2);
                                n = cf.ComputeNormal(centerUV).Normalize();
                            }
                        }
                        else
                        {
                            var box = face.GetBoundingBox();
                            var uv = new UV((box.Min.U + box.Max.U) * 0.5, (box.Min.V + box.Max.V) * 0.5);
                            n = face.ComputeNormal(uv).Normalize();
                        }
                    }
                    catch
                    {
                        n = XYZ.BasisZ;
                    }

                    var dot = Math.Abs(n.DotProduct(pipe));
                    var proj = face.Project(pt);
                    var dist = proj?.Distance ?? 1e9;

                    return (r, dot, dist);
                })
                .OrderByDescending(t => t.score) // сначала по «перпендикулярности»
                .ThenBy(t => t.dist)             // затем по близости к точке
                .FirstOrDefault();

            if (best.faceRef == null)
            {
                throw new InvalidOperationException($"Не удалось найти подходящую грань для элемента {host.Id}");
            }

            return best.faceRef;
        }

        /// <summary>
        /// Получает нормаль к грани в указанной точке (универсально для всех типов граней)
        /// </summary>
        public static XYZ GetFaceNormal(Face face, XYZ point)
        {
            if (face is PlanarFace planarFace)
            {
                return planarFace.FaceNormal.Normalize();
            }
            else if (face is CylindricalFace cylindricalFace)
            {
                var proj = cylindricalFace.Project(point);
                if (proj != null)
                {
                    var uv = proj.UVPoint;
                    return cylindricalFace.ComputeNormal(uv).Normalize();
                }
                else
                {
                    // Fallback: нормаль в центре грани
                    var bbox = cylindricalFace.GetBoundingBox();
                    var centerUV = new UV((bbox.Min.U + bbox.Max.U) / 2, (bbox.Min.V + bbox.Max.V) / 2);
                    return cylindricalFace.ComputeNormal(centerUV).Normalize();
                }
            }
            else
            {
                // Для других типов граней
                try
                {
                    var proj = face.Project(point);
                    if (proj != null)
                    {
                        var uv = proj.UVPoint;
                        return face.ComputeNormal(uv).Normalize();
                    }
                    else
                    {
                        var bbox = face.GetBoundingBox();
                        var centerUV = new UV((bbox.Min.U + bbox.Max.U) / 2, (bbox.Min.V + bbox.Max.V) / 2);
                        return face.ComputeNormal(centerUV).Normalize();
                    }
                }
                catch
                {
                    // Последний fallback
                    return XYZ.BasisZ;
                }
            }
        }

        /// <summary>
        /// Получает твердое тело хоста для точных геометрических расчетов
        /// </summary>
        public static Solid GetHostSolid(Element hostElement)
        {
            var options = new Options
            {
                ComputeReferences = false,
                IncludeNonVisibleObjects = true,
                DetailLevel = ViewDetailLevel.Fine
            };

            GeometryElement geoElem = hostElement.get_Geometry(options);
            if (geoElem == null) return null;

            foreach (GeometryObject geoObj in geoElem)
            {
                if (geoObj is Solid solid && solid.Volume > 0)
                {
                    return solid;
                }
            }

            return null;
        }

        /// <summary>
        /// Проверяет, попадает ли точка в проём окна/двери стены
        /// </summary>
        public static bool IsInDoorOrWindowOpening(Document doc, Element host, XYZ point)
        {
            if (!(host is Wall wall)) return false;

            try
            {
                var openings = wall.FindInserts(true, true, true, true);
                double tolFt = UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Millimeters); // толеранс 50мм

                foreach (var openingId in openings)
                {
                    var opening = doc.GetElement(openingId);
                    if (opening == null) continue;

                    var bb = opening.get_BoundingBox(null);
                    if (bb == null) continue;

                    // Проверяем попадание точки в bbox проёма с толерансом
                    if (point.X >= bb.Min.X - tolFt && point.X <= bb.Max.X + tolFt &&
                        point.Y >= bb.Min.Y - tolFt && point.Y <= bb.Max.Y + tolFt &&
                        point.Z >= bb.Min.Z - tolFt && point.Z <= bb.Max.Z + tolFt)
                    {
                        return true;
                    }
                }
            }
            catch
            {
                // При ошибке считаем что проёма нет
            }

            return false;
        }
    }
}
