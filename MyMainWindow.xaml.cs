using System.Collections.Generic;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitMEPHoleManager
{
    public partial class MainWindow : Window
    {
        private readonly UIApplication _uiApp;   // ссылка на Revit

        public MainWindow(UIApplication uiApp)    // NEW
        {
            InitializeComponent();
            _uiApp = uiApp;
        }

        /// <summary>Обработчик кнопки «Старт».</summary>
        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            UIDocument uiDoc = _uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // 1. Собираем стены
            var walls = new FilteredElementCollector(doc)
                        .OfClass(typeof(Wall))
                        .ToElements();

            // 2. Собираем инженерные элементы
            var pipes = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_PipeCurves)
                        .ToElements();

            var ducts = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_DuctCurves)
                        .ToElements();

            var trays = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_CableTray)
                        .ToElements();

            bool intersectionFound = false;

            // 3. Быстрая проверка: пересечение BoundingBox'ов
            foreach (Element wall in walls)
            {
                BoundingBoxXYZ wallBox = wall.get_BoundingBox(null);
                if (wallBox == null) continue;

                // локальная функция для проверки списка
                bool Check(IEnumerable<Element> elems)
                {
                    foreach (Element el in elems)
                    {
                        BoundingBoxXYZ box = el.get_BoundingBox(null);
                        if (box == null) continue;

                        // пересекаются ли проекции по X,Y,Z
                        if (wallBox.Max.X >= box.Min.X && wallBox.Min.X <= box.Max.X &&
                            wallBox.Max.Y >= box.Min.Y && wallBox.Min.Y <= box.Max.Y &&
                            wallBox.Max.Z >= box.Min.Z && wallBox.Min.Z <= box.Max.Z)
                            return true;
                    }
                    return false;
                }

                if (Check(pipes) || Check(ducts) || Check(trays))
                {
                    intersectionFound = true;
                    break;
                }
            }

            // 4. Уведомление
            MessageBox.Show(intersectionFound ? "Есть пересечение"
                                              : "Пересечений не найдено",
                            "Проверка коллизий",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);
        }
    }
}
