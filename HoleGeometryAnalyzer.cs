using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace RevitMEPHoleManager
{
    /// <summary>
    /// Данные о границах семейства отверстия
    /// </summary>
    public class FamilyBounds
    {
        public ElementId HoleId { get; set; }
        public double LeftMm { get; set; }    // самая левая грань в мм
        public double RightMm { get; set; }   // самая правая грань в мм  
        public double BottomMm { get; set; }  // самая нижняя грань в мм
        public double TopMm { get; set; }     // самая верхняя грань в мм
        public double FrontMm { get; set; }   // передняя грань в мм
        public double BackMm { get; set; }    // задняя грань в мм
        public XYZ CenterPoint { get; set; }  // центр отверстия
        public List<Face> Faces { get; set; } // все грани отверстия
    }

    /// <summary>
    /// Класс для анализа геометрии семейств отверстий через Revit API
    /// </summary>
    public static class HoleGeometryAnalyzer
    {
        /// <summary>
        /// Анализирует геометрию семейства отверстия и находит крайние грани
        /// </summary>
        public static FamilyBounds AnalyzeFamilyGeometry(FamilyInstance hole, HoleLogger logger)
        {
            try
            {
                logger.Add($"    ═══ АНАЛИЗ ГЕОМЕТРИИ СЕМЕЙСТВА {hole.Id} ═══");
                
                var bounds = new FamilyBounds
                {
                    HoleId = hole.Id,
                    Faces = new List<Face>(),
                    LeftMm = double.MaxValue,
                    RightMm = double.MinValue,
                    BottomMm = double.MaxValue,
                    TopMm = double.MinValue,
                    FrontMm = double.MaxValue,
                    BackMm = double.MinValue
                };

                // Пробуем разные настройки для получения геометрии
                GeometryElement geomElem = null;
                
                // Первая попытка: с максимальными настройками
                Options opt1 = new Options();
                opt1.ComputeReferences = true;
                opt1.IncludeNonVisibleObjects = true;
                opt1.DetailLevel = ViewDetailLevel.Fine;
                geomElem = hole.get_Geometry(opt1);
                
                if (geomElem == null)
                {
                    logger.Add($"    Первая попытка неудачна, пробуем другие настройки");
                    
                    // Вторая попытка: минимальные настройки
                    Options opt2 = new Options();
                    opt2.DetailLevel = ViewDetailLevel.Coarse;
                    geomElem = hole.get_Geometry(opt2);
                }
                
                if (geomElem == null)
                {
                    logger.Add($"    ❌ Не удалось получить геометрию семейства с любыми настройками");
                    return null;
                }

                logger.Add($"    Анализ GeometryElement семейства {hole.Symbol.Name}");
                
                int faceCount = 0;
                foreach (GeometryObject geomObj in geomElem)
                {
                    logger.Add($"    GeometryObject тип: {geomObj.GetType().Name}");
                    
                    if (geomObj is Solid solid && solid.Faces.Size > 0)
                    {
                        logger.Add($"    Найден Solid с {solid.Faces.Size} гранями");
                        AnalyzeSolidFaces(solid, bounds, logger, ref faceCount);
                    }
                    else if (geomObj is GeometryInstance geomInst)
                    {
                        logger.Add($"    Найден GeometryInstance, анализируем вложенную геометрию");
                        
                        // Получаем трансформацию экземпляра
                        Transform instTransform = geomInst.Transform;
                        logger.Add($"    Transform: Origin=({instTransform.Origin.X * 304.8:F1}, {instTransform.Origin.Y * 304.8:F1}, {instTransform.Origin.Z * 304.8:F1})");
                        
                        GeometryElement instGeom = geomInst.GetInstanceGeometry();
                        
                        if (instGeom != null)
                        {
                            foreach (GeometryObject instObj in instGeom)
                            {
                                logger.Add($"    Вложенный объект тип: {instObj.GetType().Name}");
                                
                                if (instObj is Solid instSolid && instSolid.Faces.Size > 0)
                                {
                                    logger.Add($"    Вложенный Solid с {instSolid.Faces.Size} гранями");
                                    AnalyzeSolidFaces(instSolid, bounds, logger, ref faceCount);
                                }
                                else if (instObj is GeometryInstance nestedInst)
                                {
                                    logger.Add($"    Найден вложенный GeometryInstance второго уровня");
                                    // Рекурсивно анализируем еще более глубокую геометрию
                                    GeometryElement nestedGeom = nestedInst.GetInstanceGeometry();
                                    if (nestedGeom != null)
                                    {
                                        foreach (GeometryObject nestedObj in nestedGeom)
                                        {
                                            logger.Add($"    Объект 2-го уровня тип: {nestedObj.GetType().Name}");
                                            if (nestedObj is Solid nestedSolid && nestedSolid.Faces.Size > 0)
                                            {
                                                logger.Add($"    Solid 2-го уровня с {nestedSolid.Faces.Size} гранями");
                                                AnalyzeSolidFaces(nestedSolid, bounds, logger, ref faceCount);
                                            }
                                        }
                                    }
                                }
                                else if (instObj is Mesh mesh)
                                {
                                    logger.Add($"    Найден Mesh с {mesh.NumTriangles} треугольниками");
                                    // Для Mesh можем получить vertices и вычислить границы
                                    AnalyzeMeshGeometry(mesh, bounds, logger);
                                }
                                else
                                {
                                    logger.Add($"    Неизвестный тип геометрии: {instObj.GetType().Name}");
                                }
                            }
                        }
                        else
                        {
                            logger.Add($"    ❌ GetInstanceGeometry() вернул null");
                        }
                    }
                }

                if (faceCount == 0)
                {
                    logger.Add($"    ⚠️ Грани не найдены, используем BoundingBox");
                    var bbox = hole.get_BoundingBox(null);
                    if (bbox != null)
                    {
                        bounds.LeftMm = bbox.Min.X * 304.8;
                        bounds.RightMm = bbox.Max.X * 304.8;
                        bounds.BottomMm = bbox.Min.Y * 304.8;
                        bounds.TopMm = bbox.Max.Y * 304.8;
                        bounds.FrontMm = bbox.Min.Z * 304.8;
                        bounds.BackMm = bbox.Max.Z * 304.8;
                        bounds.CenterPoint = (bbox.Min + bbox.Max) / 2.0;
                    }
                }
                else
                {
                    // Вычисляем центр по границам
                    bounds.CenterPoint = new XYZ(
                        (bounds.LeftMm + bounds.RightMm) / 2.0 / 304.8,
                        (bounds.BottomMm + bounds.TopMm) / 2.0 / 304.8,
                        (bounds.FrontMm + bounds.BackMm) / 2.0 / 304.8
                    );
                }

                logger.Add($"    ═══ РЕЗУЛЬТАТ АНАЛИЗА ═══");
                logger.Add($"    Найдено граней: {faceCount}");
                logger.Add($"    Левая граница:   {bounds.LeftMm:F1}мм");
                logger.Add($"    Правая граница:  {bounds.RightMm:F1}мм");
                logger.Add($"    Нижняя граница:  {bounds.BottomMm:F1}мм");
                logger.Add($"    Верхняя граница: {bounds.TopMm:F1}мм");
                logger.Add($"    Передняя граница: {bounds.FrontMm:F1}мм");
                logger.Add($"    Задняя граница:  {bounds.BackMm:F1}мм");
                logger.Add($"    Центр: ({bounds.CenterPoint.X * 304.8:F1}, {bounds.CenterPoint.Y * 304.8:F1}, {bounds.CenterPoint.Z * 304.8:F1})");
                
                return bounds;
            }
            catch (Exception ex)
            {
                logger.Add($"    ❌ Ошибка анализа геометрии: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Анализирует грани солида и обновляет границы
        /// </summary>
        private static void AnalyzeSolidFaces(Solid solid, FamilyBounds bounds, HoleLogger logger, ref int faceCount)
        {
            foreach (Face face in solid.Faces)
            {
                faceCount++;
                bounds.Faces.Add(face);
                
                try
                {
                    // Получаем bounding box грани
                    BoundingBoxUV faceBBox = face.GetBoundingBox();
                    
                    logger.Add($"    Грань #{faceCount}: тип {face.GetType().Name}");
                    
                    // Анализируем углы грани в 3D пространстве
                    var corners = GetFaceCorners(face, faceBBox);
                    
                    foreach (var corner in corners)
                    {
                        double xMm = corner.X * 304.8;
                        double yMm = corner.Y * 304.8;
                        double zMm = corner.Z * 304.8;
                        
                        // Обновляем крайние значения
                        bounds.LeftMm = Math.Min(bounds.LeftMm, xMm);
                        bounds.RightMm = Math.Max(bounds.RightMm, xMm);
                        bounds.BottomMm = Math.Min(bounds.BottomMm, yMm);
                        bounds.TopMm = Math.Max(bounds.TopMm, yMm);
                        bounds.FrontMm = Math.Min(bounds.FrontMm, zMm);
                        bounds.BackMm = Math.Max(bounds.BackMm, zMm);
                    }
                    
                    // Дополнительно проверяем центр грани
                    var centerUV = new UV(
                        (faceBBox.Min.U + faceBBox.Max.U) / 2.0,
                        (faceBBox.Min.V + faceBBox.Max.V) / 2.0
                    );
                    
                    var centerXYZ = face.Evaluate(centerUV);
                    double centerXMm = centerXYZ.X * 304.8;
                    double centerYMm = centerXYZ.Y * 304.8;
                    double centerZMm = centerXYZ.Z * 304.8;
                    
                    logger.Add($"      Центр грани: ({centerXMm:F1}, {centerYMm:F1}, {centerZMm:F1})");
                    
                    // Получаем нормаль грани
                    var normal = face.ComputeNormal(centerUV);
                    logger.Add($"      Нормаль: ({normal.X:F3}, {normal.Y:F3}, {normal.Z:F3})");
                    
                    // Определяем ориентацию грани
                    string orientation = GetFaceOrientation(normal);
                    logger.Add($"      Ориентация: {orientation}");
                }
                catch (Exception ex)
                {
                    logger.Add($"      ❌ Ошибка анализа грани: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Получает углы грани в 3D пространстве
        /// </summary>
        private static List<XYZ> GetFaceCorners(Face face, BoundingBoxUV faceBBox)
        {
            var corners = new List<XYZ>();
            
            try
            {
                // Создаем сетку точек по UV параметрам грани
                int steps = 5; // количество шагов по каждой оси
                
                for (int u = 0; u <= steps; u++)
                {
                    for (int v = 0; v <= steps; v++)
                    {
                        double uParam = faceBBox.Min.U + (faceBBox.Max.U - faceBBox.Min.U) * u / steps;
                        double vParam = faceBBox.Min.V + (faceBBox.Max.V - faceBBox.Min.V) * v / steps;
                        
                        var uv = new UV(uParam, vParam);
                        var xyz = face.Evaluate(uv);
                        corners.Add(xyz);
                    }
                }
            }
            catch
            {
                // Fallback: только углы
                try
                {
                    corners.Add(face.Evaluate(faceBBox.Min));
                    corners.Add(face.Evaluate(faceBBox.Max));
                    corners.Add(face.Evaluate(new UV(faceBBox.Min.U, faceBBox.Max.V)));
                    corners.Add(face.Evaluate(new UV(faceBBox.Max.U, faceBBox.Min.V)));
                }
                catch
                {
                    // Если и это не работает, возвращаем пустой список
                }
            }
            
            return corners;
        }

        /// <summary>
        /// Анализирует геометрию Mesh и обновляет границы
        /// </summary>
        private static void AnalyzeMeshGeometry(Mesh mesh, FamilyBounds bounds, HoleLogger logger)
        {
            try
            {
                logger.Add($"    Анализ Mesh с {mesh.NumTriangles} треугольниками");
                
                int vertexCount = 0;
                for (int i = 0; i < mesh.NumTriangles; i++)
                {
                    MeshTriangle triangle = mesh.get_Triangle(i);
                    
                    // Анализируем вершины треугольника
                    for (int v = 0; v < 3; v++)
                    {
                        XYZ vertex = triangle.get_Vertex(v);
                        vertexCount++;
                        
                        double xMm = vertex.X * 304.8;
                        double yMm = vertex.Y * 304.8;
                        double zMm = vertex.Z * 304.8;
                        
                        // Обновляем крайние значения
                        bounds.LeftMm = Math.Min(bounds.LeftMm, xMm);
                        bounds.RightMm = Math.Max(bounds.RightMm, xMm);
                        bounds.BottomMm = Math.Min(bounds.BottomMm, yMm);
                        bounds.TopMm = Math.Max(bounds.TopMm, yMm);
                        bounds.FrontMm = Math.Min(bounds.FrontMm, zMm);
                        bounds.BackMm = Math.Max(bounds.BackMm, zMm);
                    }
                }
                
                logger.Add($"    Проанализировано вершин Mesh: {vertexCount}");
            }
            catch (Exception ex)
            {
                logger.Add($"    ❌ Ошибка анализа Mesh: {ex.Message}");
            }
        }

        /// <summary>
        /// Определяет ориентацию грани по нормали
        /// </summary>
        private static string GetFaceOrientation(XYZ normal)
        {
            var n = normal.Normalize();
            
            // Допуск для определения ориентации
            double tolerance = 0.7; // ~45 градусов
            
            if (Math.Abs(n.X) > tolerance)
                return n.X > 0 ? "Правая (+X)" : "Левая (-X)";
            else if (Math.Abs(n.Y) > tolerance)
                return n.Y > 0 ? "Верхняя (+Y)" : "Нижняя (-Y)";
            else if (Math.Abs(n.Z) > tolerance)
                return n.Z > 0 ? "Передняя (+Z)" : "Задняя (-Z)";
            else
                return "Наклонная";
        }

        /// <summary>
        /// Сравнивает границы отверстий и определяет общие крайние грани
        /// </summary>
        public static void CompareFamilyBounds(List<FamilyBounds> familyBounds, HoleLogger logger)
        {
            if (familyBounds.Count < 2) return;
            
            logger.Add($"    ═══ СРАВНЕНИЕ ГРАНИЦ СЕМЕЙСТВ ═══");
            
            // Находим общие крайние значения
            double globalLeft = familyBounds.Min(f => f.LeftMm);
            double globalRight = familyBounds.Max(f => f.RightMm);
            double globalBottom = familyBounds.Min(f => f.BottomMm);
            double globalTop = familyBounds.Max(f => f.TopMm);
            double globalFront = familyBounds.Min(f => f.FrontMm);
            double globalBack = familyBounds.Max(f => f.BackMm);
            
            logger.Add($"    Общие границы всех отверстий:");
            logger.Add($"      Левая:    {globalLeft:F1}мм");
            logger.Add($"      Правая:   {globalRight:F1}мм");
            logger.Add($"      Нижняя:   {globalBottom:F1}мм");
            logger.Add($"      Верхняя:  {globalTop:F1}мм");
            logger.Add($"      Передняя: {globalFront:F1}мм");
            logger.Add($"      Задняя:   {globalBack:F1}мм");
            
            // Находим какие отверстия образуют крайние границы
            logger.Add($"    ═══ ОПРЕДЕЛЕНИЕ КРАЙНИХ ОТВЕРСТИЙ ═══");
            
            var leftmost = familyBounds.Where(f => Math.Abs(f.LeftMm - globalLeft) < 1.0).ToList();
            var rightmost = familyBounds.Where(f => Math.Abs(f.RightMm - globalRight) < 1.0).ToList();
            var bottommost = familyBounds.Where(f => Math.Abs(f.BottomMm - globalBottom) < 1.0).ToList();
            var topmost = familyBounds.Where(f => Math.Abs(f.TopMm - globalTop) < 1.0).ToList();
            
            logger.Add($"    Самые левые отверстия: {string.Join(", ", leftmost.Select(f => f.HoleId.IntegerValue))}");
            logger.Add($"    Самые правые отверстия: {string.Join(", ", rightmost.Select(f => f.HoleId.IntegerValue))}");
            logger.Add($"    Самые нижние отверстия: {string.Join(", ", bottommost.Select(f => f.HoleId.IntegerValue))}");
            logger.Add($"    Самые верхние отверстия: {string.Join(", ", topmost.Select(f => f.HoleId.IntegerValue))}");
            
            // Размеры объединенного отверстия
            double mergedWidth = globalRight - globalLeft;
            double mergedHeight = globalTop - globalBottom;
            double mergedDepth = globalBack - globalFront;
            
            logger.Add($"    ═══ РАЗМЕРЫ ОБЪЕДИНЕННОГО ОТВЕРСТИЯ ═══");
            logger.Add($"    Ширина (правая - левая):   {mergedWidth:F1}мм");
            logger.Add($"    Высота (верхняя - нижняя): {mergedHeight:F1}мм");
            logger.Add($"    Глубина (задняя - передняя): {mergedDepth:F1}мм");
            
            // Центр объединенного отверстия
            double centerX = (globalLeft + globalRight) / 2.0;
            double centerY = (globalBottom + globalTop) / 2.0;
            double centerZ = (globalFront + globalBack) / 2.0;
            
            logger.Add($"    Центр объединенного: ({centerX:F1}, {centerY:F1}, {centerZ:F1})");
        }
    }
}
