using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Data;
using System.ComponentModel;
using System.Globalization;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Plumbing;   // для Pipe
using Autodesk.Revit.DB.Mechanical; // для Duct

namespace RevitMEPHoleManager
{
    internal sealed class HoleLogger
    {
        private readonly StringBuilder sb = new();
        public void Add(string line) => sb.AppendLine(line);
        public void HR() => sb.AppendLine(new string('─', 70));
        public override string ToString() => sb.ToString();
    }
    
    public partial class MainWindow : Window
    {
        private readonly UIApplication _uiApp;

        /// <summary>
        /// Универсальный поиск подходящей грани хоста для размещения face-based семейства
        /// </summary>
        private static Reference PickHostFace(Document doc, Element host, XYZ pt, XYZ pipeDir)
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

            // выбираем грань, чья нормаль максимально сонаправлена с осью трубы
            var best = refs
                .Where(r => r != null) // исключаем null references
                .Select(r =>
                {
                    Face face;
                    try
                    {
                        face = host.GetGeometryObjectFromReference(r) as Face;
                        if (face == null) return (r, 0.0);
                    }
                    catch
                    {
                        return (r, 0.0);
                    }
                    
                    XYZ faceNormal;
                    if (face is PlanarFace planarFace)
                    {
                        // Плоская грань - используем обычную нормаль
                        faceNormal = planarFace.FaceNormal.Normalize();
                    }
                    else if (face is CylindricalFace cylindricalFace)
                    {
                        // Цилиндрическая грань (дуговые стены) - нормаль в точке пересечения
                        var proj = cylindricalFace.Project(pt);
                        if (proj != null)
                        {
                            var uv = proj.UVPoint;
                            faceNormal = cylindricalFace.ComputeNormal(uv).Normalize();
                        }
                        else
                        {
                            // Fallback: нормаль в центре грани
                            var bbox = cylindricalFace.GetBoundingBox();
                            var centerUV = new UV((bbox.Min.U + bbox.Max.U) / 2, (bbox.Min.V + bbox.Max.V) / 2);
                            faceNormal = cylindricalFace.ComputeNormal(centerUV).Normalize();
                        }
                    }
                    else
                    {
                        // Другие типы граней - fallback
                        try
                        {
                            var bbox = face.GetBoundingBox();
                            var centerUV = new UV((bbox.Min.U + bbox.Max.U) / 2, (bbox.Min.V + bbox.Max.V) / 2);
                            faceNormal = face.ComputeNormal(centerUV).Normalize();
                        }
                        catch
                        {
                            return (r, 0.0);
                        }
                    }
                    
                    double dot = Math.Abs(faceNormal.DotProduct(pipeDir.Normalize()));
                    return (r, dot);
                })
                .Where(t => t.Item2 > 0.0) // исключаем грани с ошибками
                .OrderByDescending(t => t.Item2)
                .FirstOrDefault();

            if (best.r == null)
            {
                throw new InvalidOperationException($"Не удалось найти подходящую грань для элемента {host.Id}");
            }

            return best.r;
        }

        /// <summary>
        /// Проверяет, попадает ли точка в проём окна/двери стены
        /// </summary>
        private static bool IsInDoorOrWindowOpening(Document doc, Element host, XYZ point)
        {
            // Для простоты: учитываем только стены
            if (host is not Wall wall) return false;

            const double tolFt = 5.0 / 304.8; // 5 мм допуск

            // Все дверные/оконные экземпляры, у которых Host == данная стена
            var doors = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Doors)
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .Where(fi => fi.Host?.Id == wall.Id);

            var wins = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Windows)
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .Where(fi => fi.Host?.Id == wall.Id);

            foreach (var fi in doors.Concat(wins))
            {
                var bb = fi.get_BoundingBox(null);
                if (bb == null) continue;

                // Простой AABB‑тест с небольшим допуском
                if (point.X >= bb.Min.X - tolFt && point.X <= bb.Max.X + tolFt &&
                    point.Y >= bb.Min.Y - tolFt && point.Y <= bb.Max.Y + tolFt &&
                    point.Z >= bb.Min.Z - tolFt && point.Z <= bb.Max.Z + tolFt)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Строим проверочный бокс вокруг точки (в футах)
        /// </summary>
        private static BoundingBoxXYZ BuildCheckBox(XYZ center, double halfX, double halfY, double halfZ)
        {
            return new BoundingBoxXYZ
            {
                Min = new XYZ(center.X - halfX, center.Y - halfY, center.Z - halfZ),
                Max = new XYZ(center.X + halfX, center.Y + halfY, center.Z + halfZ)
            };
        }

        /// <summary>
        /// AABB-пересечение (как в IntersectionStats.Intersects, но локально)
        /// </summary>
        private static bool IntersectsAabb(BoundingBoxXYZ a, BoundingBoxXYZ b)
        {
            return (a.Min.X <= b.Max.X && a.Max.X >= b.Min.X) &&
                   (a.Min.Y <= b.Max.Y && a.Max.Y >= b.Min.Y) &&
                   (a.Min.Z <= b.Max.Z && a.Max.Z >= b.Min.Z);
        }

        /// <summary>
        /// Рассчитывает среднюю точку между входом и выходом трубы через стену
        /// </summary>
        private XYZ CalculateMidPointOnWall(IntersectRow row, Element host, Face face)
        {
            try
            {
                // Получаем геометрию трубы
                var mep = _uiApp.ActiveUIDocument.Document.GetElement(new ElementId(row.MepId));
                if (!(mep is MEPCurve mepCurve) || !(mepCurve.Location is LocationCurve locationCurve))
                {
                    // Fallback: простая проекция центра
                    var projCenter = face.Project(row.Center);
                    return projCenter?.XYZPoint ?? row.Center;
                }

                var line = locationCurve.Curve as Line;
                if (line == null)
                {
                    // Fallback для кривых труб
                    var projCenter = face.Project(row.Center);
                    return projCenter?.XYZPoint ?? row.Center;
                }

                // Получаем начальную и конечную точки трубы
                XYZ startPt = line.GetEndPoint(0);
                XYZ endPt = line.GetEndPoint(1);

                // Получаем солид стены для точного пересечения
                var hostSolid = GetHostSolid(host);
                if (hostSolid == null)
                {
                    // Fallback: средняя точка между проекциями концов
                    var projStart = face.Project(startPt);
                    var projEnd = face.Project(endPt);
                    if (projStart != null && projEnd != null)
                    {
                        return (projStart.XYZPoint + projEnd.XYZPoint) / 2.0;
                    }
                    var projCenter = face.Project(row.Center);
                    return projCenter?.XYZPoint ?? row.Center;
                }

                // Находим точки пересечения оси трубы с солидом стены
                var intersectionOptions = new SolidCurveIntersectionOptions();
                var intersectionResult = hostSolid.IntersectWithCurve(line, intersectionOptions);

                if (intersectionResult != null && intersectionResult.SegmentCount >= 1)
                {
                    // Берем первый сегмент пересечения (трубу через стену)
                    var curve = intersectionResult.GetCurveSegment(0);
                    XYZ entryPt = curve.GetEndPoint(0);   // точка входа
                    XYZ exitPt = curve.GetEndPoint(1);    // точка выхода

                    // Средняя точка между входом и выходом
                    XYZ midPoint = (entryPt + exitPt) / 2.0;

                    // Проектируем на грань для точного позиционирования
                    var projMid = face.Project(midPoint);
                    return projMid?.XYZPoint ?? midPoint;
                }
                else
                {
                    // Fallback: если пересечение не найдено
                    var projCenter = face.Project(row.Center);
                    return projCenter?.XYZPoint ?? row.Center;
                }
            }
            catch (Exception)
            {
                // При любой ошибке возвращаем безопасную позицию
                var projCenter = face.Project(row.Center);
                return projCenter?.XYZPoint ?? row.Center;
            }
        }

        /// <summary>
        /// Получает нормаль к грани в указанной точке (универсально для всех типов граней)
        /// </summary>
        private static XYZ GetFaceNormal(Face face, XYZ point)
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
        /// Получает солид хост-элемента для точных геометрических расчетов
        /// </summary>
        private static Solid GetHostSolid(Element host)
        {
            var options = new Options { ComputeReferences = true };
            var geomElem = host.get_Geometry(options);
            
            foreach (GeometryObject geomObj in geomElem)
            {
                if (geomObj is Solid solid && solid.Volume > 1e-6)
                {
                    return solid;
                }
                else if (geomObj is GeometryInstance instance)
                {
                    var instGeom = instance.GetInstanceGeometry();
                    foreach (GeometryObject instObj in instGeom)
                    {
                        if (instObj is Solid instSolid && instSolid.Volume > 1e-6)
                        {
                            return instSolid;
                        }
                    }
                }
            }
            return null;
        }

        public MainWindow(UIApplication uiApp)
        {
            InitializeComponent();
            _uiApp = uiApp;
            PopulateGenericModelFamilies();
            
            // Автоматический анализ проекта при запуске
            LoadProjectInformation();
        }

        //────────────────────────────────────────────────
        //  ComboBox: все семейства Generic Model (face‑based)
        //────────────────────────────────────────────────
        private void PopulateGenericModelFamilies()
        {
            Document doc = _uiApp.ActiveUIDocument.Document;

            var list = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .Where(f => f.FamilyCategory != null &&
                            f.FamilyCategory.Id.IntegerValue == (int)BuiltInCategory.OST_GenericModel &&
                            f.GetFamilySymbolIds().Count > 0)
                .Select(f => new { Name = f.Name, Id = f.Id })
                .OrderBy(i => i.Name)
                .ToList();

            FamilyCombo.ItemsSource = list;
            FamilyCombo.DisplayMemberPath = "Name";
            FamilyCombo.SelectedValuePath = "Id";
            if (list.Count > 0) FamilyCombo.SelectedIndex = 0;
            
            // Настройки по умолчанию
            ClearanceBox.Text = "50";
            MergeDistBox.Text = "200";
        }

        /// <summary>
        /// Автоматический анализ всех MEP элементов и стен в проекте при запуске
        /// </summary>
        private void LoadProjectInformation()
        {
            try
            {
                Document doc = _uiApp.ActiveUIDocument.Document;
                
                // Загружаем информацию о трубах
                LoadPipeInformation(doc);
                
                // Загружаем информацию о воздуховодах
                LoadDuctInformation(doc);
                
                // Загружаем информацию о лотках
                LoadTrayInformation(doc);
                
                // Загружаем информацию о стенах
                LoadWallInformation(doc);
                
                // Добавляем сводную информацию в лог
                LogBox.Text = $"=== АНАЛИЗ ПРОЕКТА ===\r\n" +
                             $"Дата анализа: {DateTime.Now:dd.MM.yyyy HH:mm}\r\n" +
                             $"Файл: {doc.Title}\r\n\r\n" +
                             "Загружена информация о всех элементах проекта.\r\n" +
                             "Для запуска расчета отверстий нажмите 'Старт'.\r\n\r\n";
            }
            catch (Exception ex)
            {
                LogBox.Text = $"Ошибка при загрузке информации о проекте:\r\n{ex.Message}";
            }
        }

        /// <summary>
        /// Загружает информацию о трубах
        /// </summary>
        private void LoadPipeInformation(Document doc)
        {
            var pipes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeCurves)
                .ToElements()
                .Where(e => e is Pipe)
                .Cast<Pipe>()
                .ToList();

            var pipeRows = new List<PipeRow>();

            foreach (var pipe in pipes)
            {
                try
                {
                    string systemName = "Не определена";
                    string levelName = "Не определен";
                    double lengthMm = 0;
                    double dnMm = 0;
                    string status = "✅ OK";

                    // Получаем систему
                    if (pipe.MEPSystem?.Name != null)
                        systemName = pipe.MEPSystem.Name;

                    // Получаем уровень
                    var level = doc.GetElement(pipe.LevelId) as Level;
                    if (level != null)
                        levelName = level.Name;

                    // Получаем длину
                    var lengthParam = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                    if (lengthParam?.HasValue == true)
                        lengthMm = UnitUtils.ConvertFromInternalUnits(lengthParam.AsDouble(), UnitTypeId.Millimeters);

                    // Получаем диаметр
                    if (SizeHelper.TryGetSizes(pipe, out double wMm, out double hMm, out var shape))
                    {
                        dnMm = wMm;
                    }
                    else
                    {
                        status = "⚠️ Размеры не определены";
                    }

                    pipeRows.Add(new PipeRow
                    {
                        Id = pipe.Id.IntegerValue,
                        System = systemName,
                        DN = dnMm,
                        Length = lengthMm,
                        Level = levelName,
                        Status = status
                    });
                }
                catch
                {
                    pipeRows.Add(new PipeRow
                    {
                        Id = pipe.Id.IntegerValue,
                        System = "Ошибка",
                        DN = 0,
                        Length = 0,
                        Level = "Ошибка",
                        Status = "❌ Ошибка анализа"
                    });
                }
            }

            PipeGrid.ItemsSource = pipeRows;
        }

        /// <summary>
        /// Загружает информацию о воздуховодах
        /// </summary>
        private void LoadDuctInformation(Document doc)
        {
            var ducts = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_DuctCurves)
                .ToElements()
                .Where(e => e is Duct)
                .Cast<Duct>()
                .ToList();

            var ductRows = new List<DuctRow>();

            foreach (var duct in ducts)
            {
                try
                {
                    string systemName = "Не определена";
                    string levelName = "Не определен";
                    string shapeStr = "Не определена";
                    string sizeStr = "Не определен";
                    double lengthMm = 0;
                    string status = "✅ OK";

                    // Получаем систему
                    if (duct.MEPSystem?.Name != null)
                        systemName = duct.MEPSystem.Name;

                    // Получаем уровень
                    var level = doc.GetElement(duct.LevelId) as Level;
                    if (level != null)
                        levelName = level.Name;

                    // Получаем длину
                    var lengthParam = duct.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                    if (lengthParam?.HasValue == true)
                        lengthMm = UnitUtils.ConvertFromInternalUnits(lengthParam.AsDouble(), UnitTypeId.Millimeters);

                    // Получаем размеры через улучшенный алгоритм
                    if (SizeHelper.TryGetSizes(duct, out double wMm, out double hMm, out var shape))
                    {
                        shapeStr = shape.ToString();
                        
                        if (shape == ShapeKind.Round)
                            sizeStr = $"Ø{wMm:F0}";
                        else
                            sizeStr = $"{wMm:F0}×{hMm:F0}";
                    }
                    else
                    {
                        status = "⚠️ Размеры не определены";
                    }

                    ductRows.Add(new DuctRow
                    {
                        Id = duct.Id.IntegerValue,
                        System = systemName,
                        Shape = shapeStr,
                        Size = sizeStr,
                        Length = lengthMm,
                        Level = levelName,
                        Status = status
                    });
                }
                catch
                {
                    ductRows.Add(new DuctRow
                    {
                        Id = duct.Id.IntegerValue,
                        System = "Ошибка",
                        Shape = "Ошибка",
                        Size = "Ошибка",
                        Length = 0,
                        Level = "Ошибка",
                        Status = "❌ Ошибка анализа"
                    });
                }
            }

            DuctGrid.ItemsSource = ductRows;
        }

        /// <summary>
        /// Загружает информацию о кабельных лотках
        /// </summary>
        private void LoadTrayInformation(Document doc)
        {
            var trays = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_CableTray)
                .ToElements()
                .ToList();

            var trayRows = new List<TrayRow>();

            foreach (var tray in trays)
            {
                try
                {
                    string systemName = "Не определена";
                    string levelName = "Не определен";
                    string sizeStr = "Не определен";
                    double lengthMm = 0;
                    string status = "✅ OK";

                    // Получаем уровень
                    var levelParam = tray.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM);
                    if (levelParam?.HasValue == true)
                    {
                        var level = doc.GetElement(levelParam.AsElementId()) as Level;
                        if (level != null)
                            levelName = level.Name;
                    }

                    // Получаем длину
                    var lengthParam = tray.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                    if (lengthParam?.HasValue == true)
                        lengthMm = UnitUtils.ConvertFromInternalUnits(lengthParam.AsDouble(), UnitTypeId.Millimeters);

                    // Получаем размеры
                    if (SizeHelper.TryGetSizes(tray, out double wMm, out double hMm, out var shape))
                    {
                        sizeStr = $"{wMm:F0}×{hMm:F0}";
                    }
                    else
                    {
                        status = "⚠️ Размеры не определены";
                    }

                    trayRows.Add(new TrayRow
                    {
                        Id = tray.Id.IntegerValue,
                        System = systemName,
                        Size = sizeStr,
                        Length = lengthMm,
                        Level = levelName,
                        Status = status
                    });
                }
                catch
                {
                    trayRows.Add(new TrayRow
                    {
                        Id = tray.Id.IntegerValue,
                        System = "Ошибка",
                        Size = "Ошибка",
                        Length = 0,
                        Level = "Ошибка",
                        Status = "❌ Ошибка анализа"
                    });
                }
            }

            TrayGrid.ItemsSource = trayRows;
        }

        /// <summary>
        /// Загружает информацию о стенах
        /// </summary>
        private void LoadWallInformation(Document doc)
        {
            var walls = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Walls)
                .WhereElementIsNotElementType()
                .ToElements()
                .Cast<Wall>()
                .ToList();

            var wallRows = new List<WallRow>();

            foreach (var wall in walls)
            {
                try
                {
                    string wallName = wall.Name ?? "Без имени";
                    string wallType = "Не определен";
                    string levelName = "Не определен";
                    double thicknessMm = 0;
                    double areaM2 = 0;

                    // Получаем тип стены
                    var wallTypeElement = doc.GetElement(wall.GetTypeId()) as WallType;
                    if (wallTypeElement != null)
                        wallType = wallTypeElement.Name;

                    // Получаем уровень
                    var level = doc.GetElement(wall.LevelId) as Level;
                    if (level != null)
                        levelName = level.Name;

                    // Получаем толщину
                    var thicknessParam = wall.get_Parameter(BuiltInParameter.GENERIC_THICKNESS);
                    if (thicknessParam?.HasValue == true)
                        thicknessMm = UnitUtils.ConvertFromInternalUnits(thicknessParam.AsDouble(), UnitTypeId.Millimeters);

                    // Получаем площадь
                    var areaParam = wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                    if (areaParam?.HasValue == true)
                        areaM2 = UnitUtils.ConvertFromInternalUnits(areaParam.AsDouble(), UnitTypeId.SquareMeters);

                    wallRows.Add(new WallRow
                    {
                        Id = wall.Id.IntegerValue,
                        Name = wallName,
                        Type = wallType,
                        ThicknessMm = thicknessMm,
                        AreaM2 = areaM2,
                        Level = levelName
                    });
                }
                catch
                {
                    wallRows.Add(new WallRow
                    {
                        Id = wall.Id.IntegerValue,
                        Name = "Ошибка",
                        Type = "Ошибка",
                        ThicknessMm = 0,
                        AreaM2 = 0,
                        Level = "Ошибка"
                    });
                }
            }

            WallGrid.ItemsSource = wallRows;
        }

        //────────────────────────────────────────────────
        //  Старт
        //────────────────────────────────────────────────
        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            Document doc = _uiApp.ActiveUIDocument.Document;

            // 0. проверяем выбранное семейство
            if (FamilyCombo.SelectedValue == null)
            {
                MessageBox.Show("Сначала выберите face‑based семейство Generic Model.");
                return;
            }

            IList<PipeRow> allPipes = new FilteredElementCollector(doc)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .Select(p =>
                {
                    double dnMm = 0;
                    Parameter pDN = p.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                    if (pDN != null && pDN.HasValue)
                    {
                        dnMm = UnitUtils.ConvertFromInternalUnits(pDN.AsDouble(), UnitTypeId.Millimeters);
                    }

                    double lenMm = 0;
                    Parameter pLen = p.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                    if (pLen != null && pLen.HasValue)
                    {
                        lenMm = UnitUtils.ConvertFromInternalUnits(pLen.AsDouble(), UnitTypeId.Millimeters);
                    }
                    
                    return new PipeRow
                    {
                        Id = p.Id.IntegerValue,
                        System = p.Name,
                        DN = Math.Round(dnMm, 0),
                        Length = Math.Round(lenMm, 0)
                    };
                })
                .OrderBy(r => r.DN)
                .ToList();

            PipeGrid.ItemsSource = allPipes;

            //── гарантируем тип «Копия1» ──
            const string NEW_TYPE = "Копия1";
            Family family = doc.GetElement((ElementId)FamilyCombo.SelectedValue) as Family;
            FamilySymbol baseSym = doc.GetElement(family.GetFamilySymbolIds().First()) as FamilySymbol;
            FamilySymbol holeSym = family.GetFamilySymbolIds()
                                         .Select(id => doc.GetElement(id) as FamilySymbol)
                                         .FirstOrDefault(s => s.Name.Equals(NEW_TYPE, StringComparison.OrdinalIgnoreCase));

            if (holeSym == null)
            {
                using (Transaction t = new Transaction(doc, "Duplicate type"))
                {
                    t.Start();
                    holeSym = baseSym.Duplicate(NEW_TYPE) as FamilySymbol;
                    t.Commit();
                }
            }

            //── 1. хост‑элементы: стены + перекрытия  ──
            IList<Element> hostElems = new FilteredElementCollector(doc)
                .WherePasses(new LogicalOrFilter(
                                new ElementClassFilter(typeof(Wall)),
                                new ElementClassFilter(typeof(Floor))))
                .ToElements();

            //── 2. собираем MEP из хоста и всех связей ──
            ElementFilter fPipe = new ElementCategoryFilter(BuiltInCategory.OST_PipeCurves);
            ElementFilter fDuct = new ElementCategoryFilter(BuiltInCategory.OST_DuctCurves);
            ElementFilter fTray = new ElementCategoryFilter(BuiltInCategory.OST_CableTray);
            ElementFilter mepFilter = new LogicalOrFilter(new LogicalOrFilter(fPipe, fDuct), fTray);

            IEnumerable<Element> CollectMEP(Document d) =>
                new FilteredElementCollector(d).WherePasses(mepFilter).ToElements();

            var mepList = new List<(Element elem, Transform tx)>();

            foreach (Element mep in CollectMEP(doc))                       // активная модель
                mepList.Add((mep, Transform.Identity));

            foreach (RevitLinkInstance link in
                     new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>())
            {
                Document lDoc = link.GetLinkDocument();
                if (lDoc == null) continue;        // ссылка не загружена

                Transform lTx = link.GetTransform();   // связь → координаты хоста
                foreach (Element mep in CollectMEP(lDoc))
                    mepList.Add((mep, lTx));
            }

            //── 2a. КОНСТРУКТИВ: колонны + балки (активная модель + связи) ──
            ElementFilter fCol = new ElementCategoryFilter(BuiltInCategory.OST_StructuralColumns);
            ElementFilter fBeam = new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming);
            ElementFilter structFilter = new LogicalOrFilter(fCol, fBeam);

            IEnumerable<Element> CollectStruct(Document d) =>
                new FilteredElementCollector(d).WherePasses(structFilter).WhereElementIsNotElementType().ToElements();

            var structList = new List<(Element elem, Transform tx)>();

            foreach (Element structElem in CollectStruct(doc))                          // активная модель
                structList.Add((structElem, Transform.Identity));

            foreach (RevitLinkInstance link in
                     new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>())
            {
                Document lDoc = link.GetLinkDocument();
                if (lDoc == null) continue;
                Transform lTx = link.GetTransform();                           // связь → координаты хоста
                foreach (Element structElem in CollectStruct(lDoc))
                    structList.Add((structElem, lTx));
            }

            //── 3. анализ пересечений ──
            double clearance = 50;
            if (double.TryParse(ClearanceBox.Text, out double cTmp) && cTmp > 0)
                clearance = cTmp;

            // Настройка смещения позиции отверстий (мм)
            double positionOffsetMm = 5.0; // значение по умолчанию
            // TODO: Можно добавить TextBox в UI для настройки этого параметра

            var (wRnd, wRec, fRnd, fRec, clashList, hostStats) =
                IntersectionStats.Analyze(hostElems, mepList, clearance);

            //─────────────────────────────────────────────
            // 3‑A. ОБЪЕДИНЯЕМ ОТВЕРСТИЯ (новый код)
            //─────────────────────────────────────────────
            bool mergeOn = EnableMergeChk.IsChecked == true;
            double mergeDist = 0;
            if (mergeOn && double.TryParse(MergeDistBox.Text, out double mTmp) && mTmp > 0)
                mergeDist = mTmp;       // мм

            var logger = new HoleLogger();
            if (mergeOn) logger.Add($"Порог объединения: {mergeDist} мм, clearance: {clearance} мм");

            // — вычисляем «чистый» зазор между наружными контурами —
            if (mergeDist > 0)
            {
                double gapLimitFt = UnitUtils.ConvertToInternalUnits(mergeDist, UnitTypeId.Millimeters);

                foreach (var grp in clashList.GroupBy(r => r.HostId))
                {
                    var rows = grp.ToList();

                    for (int i = 0; i < rows.Count; i++)
                    {
                        double minGapFt = double.MaxValue;

                        for (int j = 0; j < rows.Count; j++)
                        {
                            if (i == j) continue;

                            // ➊ расстояние между центрами
                            double centerDistFt = rows[i].Center.DistanceTo(rows[j].Center);

                            // ➋ радиусы (DN/2) в футах
                            double r1 = rows[i].WidthFt / 2;   // widthFt вы заполняете в Analyze
                            double r2 = rows[j].WidthFt / 2;

                            // ➌ «чистый» зазор
                            double gapFt = centerDistFt - (r1 + r2);

                            if (gapFt < minGapFt) minGapFt = gapFt;
                        }

                        // ➍ если зазор меньше порога — показываем в гриде
                        if (minGapFt < gapLimitFt)
                        {
                            rows[i].GapMm = Math.Round(
                                UnitUtils.ConvertFromInternalUnits(minGapFt, UnitTypeId.Millimeters));
                }
                else
                {
                            rows[i].GapMm = null;              // ячейка остаётся пустой
                        }
                    }
                }
            }

            // Убираем предварительное объединение - будем делать это после размещения
            // if (mergeOn)
            // {
            //     clashList = HoleGeometry.MergeByIntersection(clashList, clearance, logger).ToList();
            //     LogBox.Text = logger.ToString();
            //     LogBox.ScrollToEnd();
            // }
            //─────────────────────────────────────────────

            /* --- 3.1   DataGrid с детализацией пересечений
             *            + группировка по HostId            --- */
            var cvs = new CollectionViewSource { Source = clashList };
            cvs.GroupDescriptions.Add(new PropertyGroupDescription("HostId"));
            StatsGrid.ItemsSource = cvs.View;

            /* --- 3.2   (необязательно) вторая таблица
             *            со сводкой по хостам              --- */
            HostStatsGrid.ItemsSource = hostStats;

            // 3.3. всплывающее окно‑сводка
            //TaskDialog.Show("Статистика пересечений",
              //  $"В стенах:\n  • круглые   — {wRnd}\n  • квадратные — {wRec}\n\n" +
                //$"В перекрытиях:\n  • круглые   — {fRnd}\n  • квадратные — {fRec}");

            //── 4. вставка отверстий ──
            int placed = 0;
            Options opt = new Options { ComputeReferences = true };

            // ── 4.1 кэш уже готовых типоразмеров ──
            var typeCache = new Dictionary<string, FamilySymbol>(StringComparer.OrdinalIgnoreCase);

            using (Transaction tr = new Transaction(doc, "Place holes"))
            {
                tr.Start();

                foreach (var row in clashList)          // список уже после Merge
                {
                    // ── 4.2 получаем (или создаём) FamilySymbol с именем row.HoleTypeName ──
                    FamilySymbol targetSym;
                    if (!typeCache.TryGetValue(row.HoleTypeName, out targetSym))
                    {
                        targetSym = family.GetFamilySymbolIds()
                                          .Select(id => doc.GetElement(id) as FamilySymbol)
                                          .FirstOrDefault(s => s.Name.Equals(row.HoleTypeName,
                                                                             StringComparison.OrdinalIgnoreCase));

                        if (targetSym == null)                    // если нет — дублируем "Копия1"
                        {
                            targetSym = holeSym.Duplicate(row.HoleTypeName) as FamilySymbol;
                        }
                        
                        // Устанавливаем размеры
                        SetSize(targetSym, row.HoleWidthMm, row.HoleHeightMm);
                        
                        typeCache[row.HoleTypeName] = targetSym;
                    }

                    // ── 4.3 вставляем отверстие ──
                    Element host = doc.GetElement(new ElementId(row.HostId));
                    if (host == null) continue;

                    Level hostLvl = null;

                    // Стены
                    if (host is Wall)
                    {
                        var baseP = host.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT);
                        if (baseP != null) hostLvl = doc.GetElement(baseP.AsElementId()) as Level;
                    }
                    // Плиты
                    if (hostLvl == null && host is Floor)
                    {
                        var lvlP = host.get_Parameter(BuiltInParameter.LEVEL_PARAM);
                        if (lvlP != null) hostLvl = doc.GetElement(lvlP.AsElementId()) as Level;
                    }

                    // Фолбэк — ближайший уровень по Z
                    if (hostLvl == null)
                    {
                        double z = row.CenterZft;
                        hostLvl = new FilteredElementCollector(doc)
                                      .OfClass(typeof(Level))
                                      .Cast<Level>()
                                      .OrderBy(l => Math.Abs(l.Elevation - z))
                                      .FirstOrDefault();
                    }

                    XYZ clashPt = row.GroupCtr ?? row.Center;
                    
                    // ── 4.3.1 Выбираем подходящую грань хоста ──
                    Reference faceRef;
                    try
                    {
                        faceRef = PickHostFace(doc, host, clashPt, row.PipeDir);
                        if (faceRef == null)
                        {
                            continue; // пропускаем, если грань не найдена
                        }
                    }
                    catch (Exception ex)
                    {
                        continue; // пропускаем, если не удалось найти подходящую грань
                    }

                    // получаем геометрию грани для расчета точки размещения
                    Face face;
                    try
                    {
                        face = host.GetGeometryObjectFromReference(faceRef) as Face;
                        if (face == null)
                        {
                            continue; // пропускаем, если не удалось получить геометрию грани
                        }
                    }
                    catch (Exception ex)
                    {
                        continue; // пропускаем при ошибке получения геометрии
                    }
                    
                    // ────────────────────────────────────────────────────────
                    // УЛУЧШЕННОЕ ПОЗИЦИОНИРОВАНИЕ: учитываем геометрию пересечения
                    // ────────────────────────────────────────────────────────
                    
                    XYZ placePt;
                    
                    // Если это кластерное отверстие, используем GroupCtr
                    if (row.GroupCtr != null)
                    {
                        var projGroup = face?.Project(row.GroupCtr);
                        placePt = projGroup?.XYZPoint ?? face?.Project(clashPt)?.XYZPoint ?? clashPt;
                    }
                    else
                    {
                        // Для одиночного отверстия: находим среднюю точку между входом и выходом трубы через стену
                        placePt = CalculateMidPointOnWall(row, host, face);
                    }
                    
                    if (placePt == null) continue;
                    
                    // ⬇️ НОВОЕ: не вставлять в проёмы окон/дверей
                    if (IsInDoorOrWindowOpening(doc, host, placePt))
                    {
                        continue; // пропускаем это отверстие
                    }
                    
                    XYZ normal = GetFaceNormal(face, placePt);

                    // ⬇️⬇️⬇️ НАЧАЛО НОВОГО КОДА: ФИЛЬТР "СКОЛЬЗЯЩИХ" ПЕРЕСЕЧЕНИЙ ⬇️⬇️⬇️
                    // Вычисляем скалярное произведение между направлением трубы и нормалью грани.
                    // Если векторы перпендикулярны (труба идет вдоль стены), произведение -> 0.
                    // Если параллельны (труба протыкает стену), |произведение| -> 1.
                    double dotProduct = Math.Abs(row.PipeDir.Normalize().DotProduct(normal));

                    // Устанавливаем порог. Например, 0.5 соответствует углу в 60° между векторами,
                    // или 30° между трубой и плоскостью стены. Все, что "острее", считаем скользящим.
                    const double angleThreshold = 0.5; 

                    if (dotProduct < angleThreshold)
                    {
                        // Опционально: можно добавить запись в лог для отладки
                        // logger.Add($"Пропуск (скользящее пересечение): Host {row.HostId}, MEP {row.MepId}, dot={dotProduct:F2}");
                        continue; // Пропускаем это пересечение, т.к. оно не является "проколом"
                    }
                    // ⬆️⬆️⬆️ КОНЕЦ НОВОГО КОДА ⬆️⬆️⬆️
                    
                    // ────────────────────────────────────────────────────────
                    // УЛУЧШЕННАЯ ОРИЕНТАЦИЯ: более точное направление размещения
                    // ────────────────────────────────────────────────────────
                    
                    XYZ refDir;
                    var pipeDirNormalized = row.PipeDir.Normalize();
                    
                    // Проекция оси трубы на плоскость грани
                    XYZ projectedPipeDir = pipeDirNormalized - normal * pipeDirNormalized.DotProduct(normal);
                    
                    if (projectedPipeDir.GetLength() > 1e-6)
                    {
                        // Используем проекцию оси трубы для ориентации отверстия
                        refDir = projectedPipeDir.Normalize();
                    }
                    else
                    {
                        // Труба перпендикулярна грани - используем более стабильную ориентацию
                        // Ориентируем по преобладающему направлению в плане
                        if (Math.Abs(normal.X) > Math.Abs(normal.Y))
                        {
                            // Стена примерно параллельна оси Y - ориентируем вдоль Y
                            refDir = normal.CrossProduct(XYZ.BasisY).Normalize();
                        }
                        else
                        {
                            // Стена примерно параллельна оси X - ориентируем вдоль X
                            refDir = normal.CrossProduct(XYZ.BasisX).Normalize();
                        }
                        
                        // Если всё ещё проблемы с ориентацией, используем Z
                        if (refDir.GetLength() < 1e-6)
                        {
                            refDir = normal.CrossProduct(XYZ.BasisZ).Normalize();
                        }
                    }
                    
                    // Дополнительная корректировка позиции для лучшего выравнивания
                    // Сдвигаем точку размещения в сторону оси трубы для более точного позиционирования
                    if (row.GroupCtr == null) // только для одиночных отверстий
                    {
                        // Небольшой сдвиг вдоль проекции оси трубы для точности позиционирования
                        double offsetMm = positionOffsetMm; // настраиваемое смещение
                        double offsetFt = UnitUtils.ConvertToInternalUnits(offsetMm, UnitTypeId.Millimeters);
                        
                        if (projectedPipeDir.GetLength() > 1e-6)
                        {
                            var offsetVec = projectedPipeDir.Normalize() * offsetFt;
                            placePt = placePt + offsetVec;
                            
                            // Проверяем, что точка всё ещё на грани
                            var checkProj = face.Project(placePt);
                            if (checkProj != null && checkProj.Distance < 0.05) // 15мм допуск
                            {
                                placePt = checkProj.XYZPoint;
                            }
                            // Если точка ушла с грани, возвращаем исходную позицию
                            else
                            {
                                placePt = placePt - offsetVec;
                            }
                        }
                    }

                    // ── ФИЛЬТР: если рядом колонна/балка — не вставляем отверстие ──
                    double wFt = UnitUtils.ConvertToInternalUnits(row.HoleWidthMm > 0 ? row.HoleWidthMm : row.ElemWidthMm,
                                                                  UnitTypeId.Millimeters);
                    double hFt = UnitUtils.ConvertToInternalUnits(row.HoleHeightMm > 0 ? row.HoleHeightMm : row.ElemHeightMm,
                                                                  UnitTypeId.Millimeters);

                    // толщина проверочного бокса по нормали (±150 мм)
                    double halfZ = UnitUtils.ConvertToInternalUnits(150, UnitTypeId.Millimeters);

                    // проверочный бокс строим в координатах хоста (placePt уже проектирован на грань)
                    var checkBox = BuildCheckBox(placePt, wFt / 2, hFt / 2, halfZ);

                    // пробегаем по колоннам/балкам (их bbox трансформируем в координаты хоста)
                    bool hitStruct = false;
                    foreach (var (se, tx) in structList)
                    {
                        var bb = se.get_BoundingBox(null);
                        if (bb == null) continue;

                        var bbHost = new BoundingBoxXYZ
                        {
                            Min = tx.OfPoint(bb.Min),
                            Max = tx.OfPoint(bb.Max)
                        };

                        if (IntersectsAabb(checkBox, bbHost))
                        {
                            hitStruct = true;
                            break;
                        }
                    }

                    if (hitStruct)
                    {
                        // опционально: записать в лог
                        // logger.Add($"Пропуск: рядом колонна/балка — Host {row.HostId}, MEP {row.MepId}");
                        continue; // пропускаем вставку
                    }

                    // ➊ вставляем face-based экземпляр базовым символом (Копия1)
                    FamilyInstance inst = doc.Create.NewFamilyInstance(faceRef,
                                                                       placePt, refDir, holeSym);

                    // ➋ переключаем на нужный тип
                    if (!targetSym.IsActive) targetSym.Activate();   // важно до смены
                    inst.ChangeTypeId(targetSym.Id);
                    placed++;

                    // ── 4.4 Уровень спецификации ──
                    if (hostLvl == null) continue;

                    Parameter pLvl = null;

                    // ➊ пробуем «известные» BuiltInParameter
                    BuiltInParameter[] cand =
                    {
                        BuiltInParameter.SCHEDULE_LEVEL_PARAM,              // ≤ 2023
                        BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM,// 2024+
                        BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM
                    };
                    foreach (var bip in cand)
                    {
                        pLvl = inst.get_Parameter(bip);
                        if (pLvl != null) break;
                    }

                    // ➋ fallback: первый параметр-ElementId, где значение – Level
                    if (pLvl == null)
                    {
                        foreach (Parameter p in inst.Parameters)
                        {
                            if (p.StorageType != StorageType.ElementId || p.IsReadOnly) continue;
                            Element lvlElem = doc.GetElement(p.AsElementId());
                            if (lvlElem == null || lvlElem is Level) { pLvl = p; break; }
                        }
                    }

                    // ➌ запись
                    if (pLvl != null && !pLvl.IsReadOnly)
                        pLvl.Set(hostLvl.Id);

                    // ── 4.5 Отметка от уровня ──
                    Parameter pOff = inst.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM) ??
                                     inst.LookupParameter("Отметка от уровня") ??
                                     inst.LookupParameter("Offset");
                    if (pOff != null && hostLvl != null && !pOff.IsReadOnly)
                    {
                        double halfHft = UnitUtils.ConvertToInternalUnits(
                                             row.HoleHeightMm / 2.0, UnitTypeId.Millimeters);
                        double offset = row.CenterZft - hostLvl.Elevation - halfHft;
                        pOff.Set(offset);                               // футы
                    }
                }

                tr.Commit();
            }

            // ═══ НОВАЯ ЛОГИКА: ДВУХЭТАПНОЕ ОБЪЕДИНЕНИЕ ОТВЕРСТИЙ ═══
            if (placed > 0)
            {
                logger.Add("═══ ЭТАП 2: АНАЛИЗ РАЗМЕЩЕННЫХ ОТВЕРСТИЙ ═══");
                
                if (mergeOn)
                {
                    int merged = AnalyzeAndMergeHoles(doc, family, mergeDist, logger);
                    LogBox.Text = logger.ToString();
                    LogBox.ScrollToEnd();
                    
                    MessageBox.Show($"Размещено отверстий: {placed}\nОбъединено кластеров: {merged}",
                                    "Результат", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    // Анализируем пересечения БЕЗ объединения - только для диагностики
                    AnalyzeHoleIntersections(doc, family, mergeDist, logger);
                    LogBox.Text = logger.ToString();
                    LogBox.ScrollToEnd();
                    
                    MessageBox.Show($"Вставлено отверстий: {placed}",
                                    "Результат", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            else
            {
                MessageBox.Show($"Вставлено отверстий: {placed}",
                                "Результат", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        /// <summary>
        /// Универсальная установка размеров типоразмера семейства
        /// </summary>
        void SetSize(FamilySymbol fs, double wMm, double hMm)
        {
            // 1️⃣ Width
            Parameter pW = fs.get_Parameter(BuiltInParameter.GENERIC_WIDTH) ??
                           FindByNames(fs, "Width", "W", "B", "HoleWidth", "Ширина");
            if (pW != null && !pW.IsReadOnly)
                pW.Set(UnitUtils.ConvertToInternalUnits(wMm, UnitTypeId.Millimeters));

            // 2️⃣ Height
            Parameter pH = fs.get_Parameter(BuiltInParameter.GENERIC_HEIGHT) ??
                           FindByNames(fs, "Height", "H", "HoleHeight", "Высота");
            if (pH != null && !pH.IsReadOnly)
                pH.Set(UnitUtils.ConvertToInternalUnits(hMm, UnitTypeId.Millimeters));
        }

        /// <summary>
        /// Служебный поиск параметра по списку имён
        /// </summary>
        Parameter FindByNames(FamilySymbol fs, params string[] names)
        {
            foreach (string n in names)
            {
                Parameter p = fs.LookupParameter(n);
                if (p != null) return p;
            }
            return null;
        }

        /// <summary>
        /// Двухэтапный анализ: находит размещенные отверстия и объединяет пересекающиеся
        /// </summary>
        private int AnalyzeAndMergeHoles(Document doc, Family holeFamily, double mergeThresholdMm, HoleLogger log)
        {
            int mergedCount = 0;
            
            try
            {
                // Шаг 1: Найти все размещенные экземпляры семейства отверстий
                var placedHoles = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol.Family.Id == holeFamily.Id)
                    .ToList();

                log.Add($"Найдено размещенных отверстий: {placedHoles.Count}");
                log.Add($"Порог объединения: {mergeThresholdMm}мм");

                if (placedHoles.Count < 2) return 0; // Нет смысла объединять

                // Шаг 2: Группируем по хост-элементу (стена/плита)
                var hostGroups = placedHoles
                    .Where(hole => hole.Host != null)
                    .GroupBy(hole => hole.Host.Id.IntegerValue)
                    .Where(group => group.Count() > 1) // Только группы с несколькими отверстиями
                    .ToList();

                log.Add($"Хостов с несколькими отверстиями: {hostGroups.Count}");

                // ДИАГНОСТИКА: покажем все отверстия и их позиции
                foreach (var hole in placedHoles)
                {
                    var pos = GetLocalPosition(hole);
                    var hostId = hole.Host?.Id.IntegerValue ?? -1;
                    log.Add($"  Отверстие {hole.Id}: Host={hostId}, позиция=({pos?.X * 304.8:F0}, {pos?.Y * 304.8:F0}, {pos?.Z * 304.8:F0})");
                }

                using (Transaction tr = new Transaction(doc, "Merge intersecting holes"))
                {
                    tr.Start();

                    foreach (var hostGroup in hostGroups)
                    {
                        log.Add($"Анализ хоста ID {hostGroup.Key}:");
                        
                        var holeInstances = hostGroup.ToList();
                        var processed = new HashSet<FamilyInstance>();
                        
                        // Шаг 3: Находим пересекающиеся отверстия
                        foreach (var hole1 in holeInstances)
                        {
                            if (processed.Contains(hole1)) continue;

                            var cluster = new List<FamilyInstance> { hole1 };
                            processed.Add(hole1);

                            // Ищем все отверстия, пересекающиеся с текущим кластером
                            bool expanded = true;
                            while (expanded)
                            {
                                expanded = false;
                                foreach (var hole2 in holeInstances)
                                {
                                    if (processed.Contains(hole2)) continue;

                                    // Проверяем пересечение с любым отверстием в кластере
                                    if (cluster.Any(clusterHole => HolesIntersect(clusterHole, hole2, mergeThresholdMm, log)))
                                    {
                                        cluster.Add(hole2);
                                        processed.Add(hole2);
                                        expanded = true;
                                        log.Add($"    Добавлено отверстие {hole2.Id} в кластер");
                                    }
                                }
                            }

                            // Шаг 4: Если в кластере больше одного отверстия - объединяем
                            if (cluster.Count > 1)
                            {
                                var mergedHole = CreateMergedHole(doc, cluster, holeFamily, log);
                                if (mergedHole != null)
                                {
                                    // Удаляем исходные отверстия
                                    foreach (var oldHole in cluster)
                                    {
                                        doc.Delete(oldHole.Id);
                                    }
                                    mergedCount++;
                                    log.Add($"    ✅ Создан объединенный кластер ({cluster.Count} отверстий)");
                                }
                            }
                        }
                    }

                tr.Commit();
            }

                log.Add($"Итого объединено кластеров: {mergedCount}");
            }
            catch (Exception ex)
            {
                log.Add($"Ошибка при объединении отверстий: {ex.Message}");
            }

            return mergedCount;
        }

        /// <summary>
        /// Анализирует пересечения отверстий БЕЗ объединения - только для диагностики
        /// </summary>
        private void AnalyzeHoleIntersections(Document doc, Family holeFamily, double mergeThresholdMm, HoleLogger log)
        {
            try
            {
                // Найти все размещенные экземпляры семейства отверстий
                var placedHoles = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol.Family.Id == holeFamily.Id)
                    .ToList();

                log.Add($"Найдено размещенных отверстий: {placedHoles.Count}");
                log.Add($"Порог объединения: {mergeThresholdMm}мм (анализ без объединения)");

                if (placedHoles.Count < 2)
                {
                    log.Add("Менее 2 отверстий - анализ пересечений невозможен");
                    return;
                }

                // ДИАГНОСТИКА: покажем все отверстия и их позиции
                foreach (var hole in placedHoles)
                {
                    var pos = GetLocalPosition(hole);
                    var hostId = hole.Host?.Id.IntegerValue ?? -1;
                    var width = GetHoleWidth(hole);
                    var height = GetHoleHeight(hole);
                    log.Add($"  Отверстие {hole.Id}: Host={hostId}, размер={width:F0}×{height:F0}, позиция=({pos?.X * 304.8:F0}, {pos?.Y * 304.8:F0}, {pos?.Z * 304.8:F0})");
                }

                // Группируем по хост-элементу (стена/плита)
                var hostGroups = placedHoles
                    .Where(hole => hole.Host != null)
                    .GroupBy(hole => hole.Host.Id.IntegerValue)
                    .ToList();

                log.Add($"Хостов с отверстиями: {hostGroups.Count}");

                int totalPairs = 0;
                int intersectingPairs = 0;

                foreach (var hostGroup in hostGroups)
                {
                    log.Add($"═══ АНАЛИЗ ХОСТА ID {hostGroup.Key} ═══");
                    
                    var holeInstances = hostGroup.ToList();
                    if (holeInstances.Count < 2)
                    {
                        log.Add("  Только одно отверстие - пересечений нет");
                        continue;
                    }

                    // Проверяем все пары отверстий на этом хосте
                    for (int i = 0; i < holeInstances.Count; i++)
                    {
                        for (int j = i + 1; j < holeInstances.Count; j++)
                        {
                            totalPairs++;
                            bool intersects = HolesIntersect(holeInstances[i], holeInstances[j], mergeThresholdMm, log);
                            if (intersects) intersectingPairs++;
                        }
                    }
                }

                log.Add($"═══ ИТОГ АНАЛИЗА ПЕРЕСЕЧЕНИЙ ═══");
                log.Add($"Всего проверено пар: {totalPairs}");
                log.Add($"Пересекающихся пар: {intersectingPairs}");
                log.Add($"Не пересекающихся пар: {totalPairs - intersectingPairs}");
            }
            catch (Exception ex)
            {
                log.Add($"Ошибка при анализе пересечений: {ex.Message}");
            }
        }

        /// <summary>
        /// Проверяет пересечение двух отверстий в пространстве с учетом порога объединения
        /// </summary>
        private bool HolesIntersect(FamilyInstance hole1, FamilyInstance hole2, double mergeThresholdMm, HoleLogger log)
        {
            try
            {
                // Получаем размеры отверстий
                double w1 = GetHoleWidth(hole1);
                double h1 = GetHoleHeight(hole1);
                double w2 = GetHoleWidth(hole2);
                double h2 = GetHoleHeight(hole2);

                // Получаем позиции в локальных координатах стены
                XYZ pos1 = GetLocalPosition(hole1);
                XYZ pos2 = GetLocalPosition(hole2);

                log.Add($"    Проверка: {hole1.Id} vs {hole2.Id}");
                log.Add($"      Размеры: {w1:F0}×{h1:F0} vs {w2:F0}×{h2:F0}");
                log.Add($"      Позиции: ({pos1?.X * 304.8:F0}, {pos1?.Y * 304.8:F0}) vs ({pos2?.X * 304.8:F0}, {pos2?.Y * 304.8:F0})");

                if (pos1 == null || pos2 == null)
                {
                    log.Add($"      ❌ Позиция недоступна: pos1={pos1 != null}, pos2={pos2 != null}");
                    return false;
                }

                // Добавляем буферную зону (половина порога объединения)
                double bufferFt = mergeThresholdMm / 2.0 / 304.8; // мм → футы
                
                // Вычисляем расширенные границы отверстий (в футах)
                double minX1 = pos1.X - w1 / 2.0 / 304.8 - bufferFt;
                double maxX1 = pos1.X + w1 / 2.0 / 304.8 + bufferFt;
                double minY1 = pos1.Y - h1 / 2.0 / 304.8 - bufferFt;
                double maxY1 = pos1.Y + h1 / 2.0 / 304.8 + bufferFt;

                double minX2 = pos2.X - w2 / 2.0 / 304.8 - bufferFt;
                double maxX2 = pos2.X + w2 / 2.0 / 304.8 + bufferFt;
                double minY2 = pos2.Y - h2 / 2.0 / 304.8 - bufferFt;
                double maxY2 = pos2.Y + h2 / 2.0 / 304.8 + bufferFt;

                // Добавляем проверку по Z координате (высота в стене)
                double minZ1 = pos1.Z - h1 / 2.0 / 304.8 - bufferFt;
                double maxZ1 = pos1.Z + h1 / 2.0 / 304.8 + bufferFt;
                double minZ2 = pos2.Z - h2 / 2.0 / 304.8 - bufferFt;
                double maxZ2 = pos2.Z + h2 / 2.0 / 304.8 + bufferFt;

                // Проверяем пересечение расширенных прямоугольников в 3D
                bool intersects = !(maxX1 < minX2 || minX1 > maxX2 || 
                                  maxY1 < minY2 || minY1 > maxY2 ||
                                  maxZ1 < minZ2 || minZ1 > maxZ2);
                
                // Вычисляем реальное расстояние между центрами для логирования (3D расстояние)
                double distanceMm = Math.Sqrt(Math.Pow((pos2.X - pos1.X) * 304.8, 2) + 
                                            Math.Pow((pos2.Y - pos1.Y) * 304.8, 2) +
                                            Math.Pow((pos2.Z - pos1.Z) * 304.8, 2));

                log.Add($"      Буфер: {bufferFt * 304.8:F0}мм");
                log.Add($"      Границы1: X[{minX1 * 304.8:F0}..{maxX1 * 304.8:F0}] Y[{minY1 * 304.8:F0}..{maxY1 * 304.8:F0}] Z[{minZ1 * 304.8:F0}..{maxZ1 * 304.8:F0}]");
                log.Add($"      Границы2: X[{minX2 * 304.8:F0}..{maxX2 * 304.8:F0}] Y[{minY2 * 304.8:F0}..{maxY2 * 304.8:F0}] Z[{minZ2 * 304.8:F0}..{maxZ2 * 304.8:F0}]");
                log.Add($"      3D расстояние между центрами: {distanceMm:F0}мм");

                if (intersects)
                {
                    log.Add($"      ✅ Объединение: {hole1.Id} ∩ {hole2.Id}, расстояние: {distanceMm:F0}мм < порог: {mergeThresholdMm:F0}мм");
                }
                else
                {
                    log.Add($"      ❌ Пропуск: {hole1.Id} - {hole2.Id}, расстояние: {distanceMm:F0}мм, нет пересечения расширенных границ");
                }

                return intersects;
            }
            catch (Exception ex)
            {
                log.Add($"    Ошибка проверки пересечения {hole1.Id}-{hole2.Id}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Создает объединенное отверстие для кластера пересекающихся отверстий
        /// </summary>
        private FamilyInstance CreateMergedHole(Document doc, List<FamilyInstance> cluster, Family holeFamily, HoleLogger log)
        {
            try
            {
                // Вычисляем границы объединенного отверстия
                var positions = cluster.Select(GetLocalPosition).Where(p => p != null).ToList();
                var widths = cluster.Select(GetHoleWidth).ToList();
                var heights = cluster.Select(GetHoleHeight).ToList();

                if (!positions.Any()) return null;

                // Находим крайние точки всех отверстий
                var bounds = positions.Zip(widths.Zip(heights, (w, h) => new { Width = w, Height = h }), 
                    (pos, size) => new
                    {
                        MinX = pos.X - size.Width / 2.0 / 304.8,
                        MaxX = pos.X + size.Width / 2.0 / 304.8,
                        MinY = pos.Y - size.Height / 2.0 / 304.8,
                        MaxY = pos.Y + size.Height / 2.0 / 304.8,
                        Z = pos.Z
                    }).ToList();

                double minX = bounds.Min(b => b.MinX);
                double maxX = bounds.Max(b => b.MaxX);
                double minY = bounds.Min(b => b.MinY);
                double maxY = bounds.Max(b => b.MaxY);
                double avgZ = bounds.Average(b => b.Z);

                // Размеры объединенного отверстия + дополнительный зазор
                double mergedWidthMm = (maxX - minX) * 304.8 + 100; // +100мм запас
                double mergedHeightMm = (maxY - minY) * 304.8 + 100; // +100мм запас
                
                // Центр объединенного отверстия
                XYZ mergedCenter = new XYZ((minX + maxX) / 2, (minY + maxY) / 2, avgZ);

                log.Add($"    Размер объединенного: {mergedWidthMm:F0}×{mergedHeightMm:F0}мм");
                log.Add($"    Границы: X[{minX * 304.8:F0}..{maxX * 304.8:F0}] Y[{minY * 304.8:F0}..{maxY * 304.8:F0}]");

                // Создаем типоразмер для объединенного отверстия
                string typeName = $"Прям. {Math.Ceiling(mergedWidthMm)}×{Math.Ceiling(mergedHeightMm)}";
                var firstHole = cluster.First();
                var hostElement = firstHole.Host;

                // Находим или создаем типоразмер
                FamilySymbol mergedSymbol = holeFamily.GetFamilySymbolIds()
                    .Select(id => doc.GetElement(id) as FamilySymbol)
                    .FirstOrDefault(s => s.Name.Equals(typeName, StringComparison.OrdinalIgnoreCase));

                if (mergedSymbol == null)
                {
                    mergedSymbol = firstHole.Symbol.Duplicate(typeName) as FamilySymbol;
                    SetSize(mergedSymbol, mergedWidthMm, mergedHeightMm);
                }

                // Размещаем объединенное отверстие
                if (!mergedSymbol.IsActive) mergedSymbol.Activate();

                // Получаем направление первого отверстия для ориентации
                XYZ refDirection = XYZ.BasisX; // fallback
                var firstLocation = firstHole.Location as LocationPoint;
                if (firstLocation != null)
                {
                    // Используем ориентацию первого отверстия
                    var transform = firstHole.GetTransform();
                    refDirection = transform.BasisX;
                }

                // Используем хост-элемент как face-based
                var mergedInstance = doc.Create.NewFamilyInstance(mergedCenter, mergedSymbol, hostElement, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                return mergedInstance;
            }
            catch (Exception ex)
            {
                log.Add($"    Ошибка создания объединенного отверстия: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Получает ширину отверстия в мм
        /// </summary>
        private double GetHoleWidth(FamilyInstance hole)
        {
            var widthParam = hole.LookupParameter("Ширина") ?? 
                           hole.LookupParameter("Width") ??
                           hole.get_Parameter(BuiltInParameter.FAMILY_WIDTH_PARAM);
            
            if (widthParam != null && widthParam.HasValue)
            {
                return UnitUtils.ConvertFromInternalUnits(widthParam.AsDouble(), UnitTypeId.Millimeters);
            }
            return 200; // fallback
        }

        /// <summary>
        /// Получает высоту отверстия в мм
        /// </summary>
        private double GetHoleHeight(FamilyInstance hole)
        {
            var heightParam = hole.LookupParameter("Высота") ?? 
                            hole.LookupParameter("Height") ??
                            hole.get_Parameter(BuiltInParameter.FAMILY_HEIGHT_PARAM);
            
            if (heightParam != null && heightParam.HasValue)
            {
                return UnitUtils.ConvertFromInternalUnits(heightParam.AsDouble(), UnitTypeId.Millimeters);
            }
            return 200; // fallback
        }

        /// <summary>
        /// Получает локальную позицию отверстия относительно хоста
        /// </summary>
        private XYZ GetLocalPosition(FamilyInstance hole)
        {
            try
            {
                var location = hole.Location as LocationPoint;
                if (location?.Point != null)
                {
                    return location.Point;
                }
                
                // Fallback: попробуем получить из transform
                var transform = hole.GetTransform();
                if (transform != null)
                {
                    return transform.Origin;
                }
                
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка получения позиции отверстия {hole.Id}: {ex.Message}");
                return null;
            }
        }
    }

    /// <summary>
    /// Конвертер bool в "Да"/"Нет" для отображения в DataGrid
    /// </summary>
    public class BoolToYesNoConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
                return boolValue ? "Да" : "Нет";
            return "Нет";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value?.ToString() == "Да";
        }
    }
}
