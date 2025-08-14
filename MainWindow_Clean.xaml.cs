using System;
using System.Collections.Generic;
using System.Diagnostics;
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
    /// <summary>
    /// Главное окно плагина для управления отверстиями MEP
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly UIApplication _uiApp;
        private readonly Document _doc;

        public MainWindow(UIApplication uiApp)
        {
            InitializeComponent();
            _uiApp = uiApp ?? throw new ArgumentNullException(nameof(uiApp));
            _doc = uiApp.ActiveUIDocument.Document;

            try
            {
                PopulateGenericModelFamilies();
                PopulateDataOnStartup();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка инициализации: {ex.Message}");
            }
        }

        /// <summary>
        /// Заполняет ComboBox семействами Generic Model
        /// </summary>
        private void PopulateGenericModelFamilies()
        {
            try
            {
                var collector = new FilteredElementCollector(_doc);
                var genericModelFamilies = collector
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .Where(f => f.FamilyCategory?.Id.IntegerValue == (int)BuiltInCategory.OST_GenericModel)
                    .ToList();

                FamilyCombo.ItemsSource = genericModelFamilies;
                FamilyCombo.DisplayMemberPath = "Name";
                FamilyCombo.SelectedValuePath = "Id";

                if (genericModelFamilies.Any())
                {
                    // Ищем семейство со словом "отверстие" в названии
                    var holeFamily = genericModelFamilies
                        .FirstOrDefault(f => f.Name.ToLower().Contains("отверстие"));
                    
                    if (holeFamily != null)
                    {
                        FamilyCombo.SelectedItem = holeFamily;
                    }
                    else
                    {
                        FamilyCombo.SelectedIndex = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки семейств: {ex.Message}");
            }
        }

        /// <summary>
        /// Заполняет данными при запуске плагина
        /// </summary>
        private void PopulateDataOnStartup()
        {
            try
            {
                // Трубы
                var pipes = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Pipe))
                    .Cast<Pipe>()
                    .ToList();

                var pipeRows = pipes.Select(CreatePipeRow).Where(row => row != null).ToList();
                PipeDataGrid.ItemsSource = pipeRows;

                // Воздуховоды
                var ducts = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Duct))
                    .Cast<Duct>()
                    .ToList();

                var ductRows = ducts.Select(CreateDuctRow).Where(row => row != null).ToList();
                DuctDataGrid.ItemsSource = ductRows;

                // Лотки
                var trays = new FilteredElementCollector(_doc)
                    .OfClass(typeof(CableTray))
                    .Cast<CableTray>()
                    .ToList();

                var trayRows = trays.Select(CreateTrayRow).Where(row => row != null).ToList();
                TrayDataGrid.ItemsSource = trayRows;

                // Стены
                var walls = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Wall))
                    .Cast<Wall>()
                    .ToList();

                var wallRows = walls.Select(CreateWallRow).Where(row => row != null).ToList();
                WallDataGrid.ItemsSource = wallRows;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}");
            }
        }

        #region Создание строк данных для UI

        private PipeRow CreatePipeRow(Pipe pipe)
        {
            try
            {
                var connectors = pipe.ConnectorManager?.Connectors?.Cast<Connector>().ToList() ?? new List<Connector>();
                var firstConnector = connectors.FirstOrDefault();
                double dnMm = firstConnector != null 
                    ? UnitUtils.ConvertFromInternalUnits(firstConnector.Radius * 2, UnitTypeId.Millimeters)
                    : 0;

                var location = pipe.Location as LocationCurve;
                var curve = location?.Curve;
                
                return new PipeRow
                {
                    Id = pipe.Id.IntegerValue.ToString(),
                    DN = Math.Round(dnMm, 0).ToString(),
                    System = pipe.MEPSystem?.Name ?? "Нет системы",
                    IsDiagonal = IsPipeDiagonal(pipe),
                    StartX = curve?.GetEndPoint(0)?.X * 304.8 ?? 0,
                    StartY = curve?.GetEndPoint(0)?.Y * 304.8 ?? 0,
                    StartZ = curve?.GetEndPoint(0)?.Z * 304.8 ?? 0,
                    EndX = curve?.GetEndPoint(1)?.X * 304.8 ?? 0,
                    EndY = curve?.GetEndPoint(1)?.Y * 304.8 ?? 0,
                    EndZ = curve?.GetEndPoint(1)?.Z * 304.8 ?? 0,
                    Level = GetElementLevel(pipe)?.Name ?? "Нет уровня",
                    Status = "Готов"
                };
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка создания PipeRow для {pipe.Id}: {ex.Message}");
                return null;
            }
        }

        private DuctRow CreateDuctRow(Duct duct)
        {
            try
            {
                var connectors = duct.ConnectorManager?.Connectors?.Cast<Connector>().ToList() ?? new List<Connector>();
                var firstConnector = connectors.FirstOrDefault();
                
                double widthMm = firstConnector != null 
                    ? UnitUtils.ConvertFromInternalUnits(firstConnector.Width, UnitTypeId.Millimeters)
                    : 0;
                double heightMm = firstConnector != null 
                    ? UnitUtils.ConvertFromInternalUnits(firstConnector.Height, UnitTypeId.Millimeters)
                    : 0;

                var location = duct.Location as LocationCurve;
                var curve = location?.Curve;

                return new DuctRow
                {
                    Id = duct.Id.IntegerValue.ToString(),
                    Width = Math.Round(widthMm, 0).ToString(),
                    Height = Math.Round(heightMm, 0).ToString(),
                    System = duct.MEPSystem?.Name ?? "Нет системы",
                    IsDiagonal = IsDuctDiagonal(duct),
                    StartX = curve?.GetEndPoint(0)?.X * 304.8 ?? 0,
                    StartY = curve?.GetEndPoint(0)?.Y * 304.8 ?? 0,
                    StartZ = curve?.GetEndPoint(0)?.Z * 304.8 ?? 0,
                    EndX = curve?.GetEndPoint(1)?.X * 304.8 ?? 0,
                    EndY = curve?.GetEndPoint(1)?.Y * 304.8 ?? 0,
                    EndZ = curve?.GetEndPoint(1)?.Z * 304.8 ?? 0,
                    Level = GetElementLevel(duct)?.Name ?? "Нет уровня",
                    Status = "Готов"
                };
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка создания DuctRow для {duct.Id}: {ex.Message}");
                return null;
            }
        }

        private TrayRow CreateTrayRow(CableTray tray)
        {
            try
            {
                double widthMm = UnitUtils.ConvertFromInternalUnits(tray.Width, UnitTypeId.Millimeters);
                double heightMm = UnitUtils.ConvertFromInternalUnits(tray.Height, UnitTypeId.Millimeters);

                var location = tray.Location as LocationCurve;
                var curve = location?.Curve;

                return new TrayRow
                {
                    Id = tray.Id.IntegerValue.ToString(),
                    Width = Math.Round(widthMm, 0).ToString(),
                    Height = Math.Round(heightMm, 0).ToString(),
                    IsDiagonal = IsTrayDiagonal(tray),
                    StartX = curve?.GetEndPoint(0)?.X * 304.8 ?? 0,
                    StartY = curve?.GetEndPoint(0)?.Y * 304.8 ?? 0,
                    StartZ = curve?.GetEndPoint(0)?.Z * 304.8 ?? 0,
                    EndX = curve?.GetEndPoint(1)?.X * 304.8 ?? 0,
                    EndY = curve?.GetEndPoint(1)?.Y * 304.8 ?? 0,
                    EndZ = curve?.GetEndPoint(1)?.Z * 304.8 ?? 0,
                    Level = GetElementLevel(tray)?.Name ?? "Нет уровня",
                    Status = "Готов"
                };
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка создания TrayRow для {tray.Id}: {ex.Message}");
                return null;
            }
        }

        private WallRow CreateWallRow(Wall wall)
        {
            try
            {
                var location = wall.Location;
                XYZ start = null, end = null;

                if (location is LocationCurve locationCurve)
                {
                    var curve = locationCurve.Curve;
                    start = curve.GetEndPoint(0);
                    end = curve.GetEndPoint(1);
                }

                return new WallRow
                {
                    Id = wall.Id.IntegerValue.ToString(),
                    Type = wall.WallType?.Name ?? "Неизвестный тип",
                    StartX = start?.X * 304.8 ?? 0,
                    StartY = start?.Y * 304.8 ?? 0,
                    StartZ = start?.Z * 304.8 ?? 0,
                    EndX = end?.X * 304.8 ?? 0,
                    EndY = end?.Y * 304.8 ?? 0,
                    EndZ = end?.Z * 304.8 ?? 0,
                    Level = GetElementLevel(wall)?.Name ?? "Нет уровня",
                    Status = "Готов"
                };
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка создания WallRow для {wall.Id}: {ex.Message}");
                return null;
            }
        }

        #endregion

        #region Вспомогательные методы

        private bool IsPipeDiagonal(Pipe pipe)
        {
            try
            {
                var location = pipe.Location as LocationCurve;
                if (location?.Curve == null) return false;

                var start = location.Curve.GetEndPoint(0);
                var end = location.Curve.GetEndPoint(1);
                var direction = (end - start).Normalize();

                // Проверяем отклонение от основных осей
                double tolerance = 0.1;
                return Math.Abs(direction.X) > tolerance && 
                       Math.Abs(direction.Y) > tolerance &&
                       Math.Abs(direction.Z) < tolerance; // горизонтальная диагональ
            }
            catch
            {
                return false;
            }
        }

        private bool IsDuctDiagonal(Duct duct)
        {
            try
            {
                var location = duct.Location as LocationCurve;
                if (location?.Curve == null) return false;

                var start = location.Curve.GetEndPoint(0);
                var end = location.Curve.GetEndPoint(1);
                var direction = (end - start).Normalize();

                double tolerance = 0.1;
                return Math.Abs(direction.X) > tolerance && 
                       Math.Abs(direction.Y) > tolerance &&
                       Math.Abs(direction.Z) < tolerance;
            }
            catch
            {
                return false;
            }
        }

        private bool IsTrayDiagonal(CableTray tray)
        {
            try
            {
                var location = tray.Location as LocationCurve;
                if (location?.Curve == null) return false;

                var start = location.Curve.GetEndPoint(0);
                var end = location.Curve.GetEndPoint(1);
                var direction = (end - start).Normalize();

                double tolerance = 0.1;
                return Math.Abs(direction.X) > tolerance && 
                       Math.Abs(direction.Y) > tolerance &&
                       Math.Abs(direction.Z) < tolerance;
            }
            catch
            {
                return false;
            }
        }

        private Level GetElementLevel(Element element)
        {
            try
            {
                var levelParam = element.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM) ??
                               element.get_Parameter(BuiltInParameter.LEVEL_PARAM) ??
                               element.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT);

                if (levelParam?.AsElementId() != null)
                {
                    return _doc.GetElement(levelParam.AsElementId()) as Level;
                }
            }
            catch
            {
                // Игнорируем ошибки получения уровня
            }
            return null;
        }

        private static void SetDepthParam(Element e, double depthMm)
        {
            try
            {
                var depthParam = e.LookupParameter("Глубина") ?? e.LookupParameter("Depth") ??
                               e.LookupParameter("Толщина") ?? e.LookupParameter("Thickness");
                if (depthParam != null && !depthParam.IsReadOnly)
                {
                    double depthFt = UnitUtils.ConvertToInternalUnits(depthMm, UnitTypeId.Millimeters);
                    depthParam.Set(depthFt);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SetDepthParam ERROR: {ex.Message}");
            }
        }

        #endregion

        #region Обработчики событий кнопок

        /// <summary>
        /// Обработчик кнопки "Старт" - размещение одиночных отверстий
        /// </summary>
        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            var logger = new HoleLogger();
            logger.Add("═══ НАЧАЛО РАЗМЕЩЕНИЯ ОТВЕРСТИЙ ═══");

            try
            {
                var selectedFamily = _doc.GetElement((ElementId)FamilyCombo.SelectedValue) as Family;
                if (selectedFamily == null)
                {
                    MessageBox.Show("Выберите семейство отверстий");
                    return;
                }

                double clearanceMm = double.TryParse(ClearanceTextBox.Text, out double c) ? c : 50;
                logger.Add($"Clearance: {clearanceMm} мм");

                using var transaction = new Transaction(_doc, "Размещение отверстий");
                transaction.Start();

                try
                {
                    logger.Add("═══ ЭТАП РАЗМЕЩЕНИЯ ОДИНОЧНЫХ ОТВЕРСТИЙ ═══");
                    
                    var holeData = SizeHelper.CalculateAllHoles(_doc, clearanceMm);
                    logger.Add($"Количество отверстий к размещению: {holeData.Count}");
                    
                    int successCount = 0;
                    for (int i = 0; i < holeData.Count; i++)
                    {
                        var row = holeData[i];
                        logger.Add($"┌─ Отверстие {i + 1}/{holeData.Count} ─");
                        logger.Add($"│ MEP: {row.MepId}, Host: {row.HostId}");
                        logger.Add($"│ Размер: {row.WidthMm:F0}×{row.HeightMm:F0}мм");
                        logger.Add($"│ Тип: {row.HoleTypeName}");
                        logger.Add($"│ Позиция: ({row.CenterXft * 304.8:F0}, {row.CenterYft * 304.8:F0}, {row.CenterZft * 304.8:F0})");
                        
                        try
                        {
                            var placedHole = PlaceIndividualHole(_doc, selectedFamily, row, logger);
                            if (placedHole != null)
                            {
                                successCount++;
                                logger.Add($"│ ✅ Отверстие размещено! ID: {placedHole.Id}");
                            }
                            else
                            {
                                logger.Add($"│ ❌ Не удалось разместить отверстие");
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.Add($"│ ❌ Ошибка размещения: {ex.Message}");
                        }
                        
                        logger.Add($"└─ {(successCount > i ? "Успешно" : "Ошибка")} ({successCount}/{i + 1})");
                    }

                    logger.Add("═══ ЗАВЕРШЕНИЕ РАЗМЕЩЕНИЯ ═══");
                    logger.Add($"Успешно размещено: {successCount} из {holeData.Count} отверстий");

                    transaction.Commit();
                    
                    MessageBox.Show($"Размещено {successCount} из {holeData.Count} отверстий.\n" +
                                  "Для объединения пересекающихся отверстий нажмите кнопку 'Объединить'.");
                }
                catch (Exception ex)
                {
                    transaction.RollBack();
                    logger.Add($"❌ КРИТИЧЕСКАЯ ОШИБКА: {ex.Message}");
                    MessageBox.Show($"Ошибка размещения отверстий: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                logger.Add($"❌ ОШИБКА ИНИЦИАЛИЗАЦИИ: {ex.Message}");
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
            finally
            {
                LogTextBox.Text = logger.ToString();
            }
        }

        /// <summary>
        /// Обработчик кнопки "Объединить" - объединение пересекающихся отверстий
        /// </summary>
        private void MergeButton_Click(object sender, RoutedEventArgs e)
        {
            var logger = new HoleLogger();
            logger.Add("═══ АНАЛИЗ И ОБЪЕДИНЕНИЕ РАЗМЕЩЕННЫХ ОТВЕРСТИЙ ═══");
            logger.Add("🎯 Алгоритм: охватывающий прямоугольник между крайними границами отверстий");

            try
            {
                var selectedFamily = _doc.GetElement((ElementId)FamilyCombo.SelectedValue) as Family;
                if (selectedFamily == null)
                {
                    MessageBox.Show("Выберите семейство отверстий");
                    return;
                }

                double mergeThresholdMm = double.TryParse(MergeThresholdTextBox.Text, out double threshold) ? threshold : 250;
                logger.Add($"Порог объединения: {mergeThresholdMm:F0}мм");

                using var transaction = new Transaction(_doc, "Объединение отверстий");
                transaction.Start();

                try
                {
                    int mergedCount = HoleMergeManager.AnalyzeAndMergeHoles(_doc, selectedFamily, mergeThresholdMm, logger);
                    
                    transaction.Commit();
                    logger.Add($"Итого объединено кластеров: {mergedCount}");
                    
                    if (mergedCount > 0)
                    {
                        MessageBox.Show($"Успешно объединено {mergedCount} кластеров отверстий!");
                    }
                    else
                    {
                        MessageBox.Show("Отверстия для объединения не найдены или объединение невозможно.");
                    }
                }
                catch (Exception ex)
                {
                    transaction.RollBack();
                    logger.Add($"❌ Ошибка при объединении отверстий: {ex.Message}");
                    MessageBox.Show($"Ошибка объединения: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                logger.Add($"❌ Ошибка: {ex.Message}");
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
            finally
            {
                // ВСЕГДА показываем полный лог для диагностики
                LogTextBox.Text = logger.ToString();
            }
        }

        #endregion

        #region Размещение отдельных отверстий

        /// <summary>
        /// Размещает одиночное отверстие
        /// </summary>
        private FamilyInstance PlaceIndividualHole(Document doc, Family holeFamily, IntersectionStats row, HoleLogger logger)
        {
            try
            {
                // Поиск или создание типоразмера
                logger.Add($"│ Поиск типоразмера: {row.HoleTypeName}");
                FamilySymbol holeSymbol = holeFamily.GetFamilySymbolIds()
                    .Select(id => doc.GetElement(id) as FamilySymbol)
                    .FirstOrDefault(s => s.Name.Equals(row.HoleTypeName, StringComparison.OrdinalIgnoreCase));

                if (holeSymbol == null)
                {
                    logger.Add($"│ Создание нового типоразмера: {row.HoleTypeName}");
                    var baseSymbol = holeFamily.GetFamilySymbolIds()
                        .Select(id => doc.GetElement(id) as FamilySymbol)
                        .FirstOrDefault();
                        
                    if (baseSymbol != null)
                    {
                        holeSymbol = baseSymbol.Duplicate(row.HoleTypeName) as FamilySymbol;
                        HoleSizeCalculator.SetSize(holeSymbol, row.WidthMm, row.HeightMm);
                        logger.Add($"│ ✅ Создан типоразмер: {row.WidthMm:F0}×{row.HeightMm:F0}мм");
                    }
                    else
                    {
                        logger.Add($"│ ❌ Не найден базовый символ для семейства");
                        return null;
                    }
                }
                else
                {
                    logger.Add($"│ Найден существующий типоразмер: {row.HoleTypeName}");
                    HoleSizeCalculator.SetSize(holeSymbol, row.WidthMm, row.HeightMm);
                    logger.Add($"│ Установлены размеры: {row.WidthMm:F0}×{row.HeightMm:F0}мм");
                }

                if (!holeSymbol.IsActive)
                {
                    holeSymbol.Activate();
                }

                // Получаем хост и размещаем отверстие
                var host = doc.GetElement(new ElementId(row.HostId));
                if (host == null)
                {
                    logger.Add($"│ ❌ Хост элемент не найден: {row.HostId}");
                    return null;
                }

                logger.Add($"│ Размещение отверстия...");
                logger.Add($"│ Хост: {host.GetType().Name} (ID: {host.Id})");

                var centerPoint = new XYZ(row.CenterXft, row.CenterYft, row.CenterZft);
                logger.Add($"│ Создание FamilyInstance...");
                logger.Add($"│ Точка размещения: ({centerPoint.X * 304.8:F0}, {centerPoint.Y * 304.8:F0}, {centerPoint.Z * 304.8:F0})");

                FamilyInstance holeInstance;

                try
                {
                    // Пытаемся создать face-based
                    var faceRef = FaceBasedPlacer.PickHostFace(doc, host, centerPoint);
                    var refDirection = host is Wall ? XYZ.BasisX : XYZ.BasisX;
                    
                    holeInstance = doc.Create.NewFamilyInstance(faceRef, centerPoint, refDirection, holeSymbol);
                }
                catch
                {
                    // Fallback: host-based
                    holeInstance = doc.Create.NewFamilyInstance(
                        centerPoint, holeSymbol, host, 
                        Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                }

                // Устанавливаем глубину
                SetDepthParam(holeInstance, 300); // фиксированная глубина 300мм

                return holeInstance;
            }
            catch (Exception ex)
            {
                logger.Add($"│ ❌ Ошибка размещения: {ex.Message}");
                return null;
            }
        }

        #endregion
    }
}
