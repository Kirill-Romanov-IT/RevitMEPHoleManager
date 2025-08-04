using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitMEPHoleManager
{
    public partial class MainWindow : Window
    {
        private readonly UIApplication _uiApp;

        public MainWindow(UIApplication uiApp)
        {
            InitializeComponent();
            _uiApp = uiApp;
            PopulateGenericModelFamilies();      // заполняем ComboBox при открытии
        }

        // ────────────────────────────────────────────────────────────────
        //  Заполняем ComboBox всеми семействами категории Generic Model
        // ────────────────────────────────────────────────────────────────
        private void PopulateGenericModelFamilies()
        {
            Document doc = _uiApp.ActiveUIDocument.Document;

            var families = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .Where(f => f.FamilyCategory != null &&
                            f.FamilyCategory.Id.IntegerValue == (int)BuiltInCategory.OST_GenericModel &&
                            f.GetFamilySymbolIds().Count > 0)
                .Select(f => new { Name = f.Name, Id = f.Id })
                .OrderBy(i => i.Name)
                .ToList();

            FamilyCombo.ItemsSource = families;
            FamilyCombo.DisplayMemberPath = "Name";
            FamilyCombo.SelectedValuePath = "Id";
            if (families.Count > 0) FamilyCombo.SelectedIndex = 0;
        }

        // ────────────────────────────────────────────────────────────────
        //  Обработчик кнопки «Старт»
        // ────────────────────────────────────────────────────────────────
        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            Document doc = _uiApp.ActiveUIDocument.Document;

            // 1. Проверяем выбор семейства
            if (FamilyCombo.SelectedValue == null)
            {
                MessageBox.Show("Сначала выберите face-based семейство Generic Model.",
                                "Предупреждение",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            ElementId familyId = (ElementId)FamilyCombo.SelectedValue;
            Family family = doc.GetElement(familyId) as Family;
            FamilySymbol symbol = doc.GetElement(family.GetFamilySymbolIds().First()) as FamilySymbol;

            // 2. Собираем коллекции элементов
            var walls = new FilteredElementCollector(doc).OfClass(typeof(Wall)).ToElements();
            var pipes = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).ToElements();
            var ducts = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctCurves).ToElements();
            var trays = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTray).ToElements();

            // 3. Ищем первое пересечение BoundingBox'ов
            XYZ clashPt = null;
            Wall hostWall = null;

            bool BoxesIntersect(BoundingBoxXYZ a, BoundingBoxXYZ b, out XYZ center)
            {
                double minX = Math.Max(a.Min.X, b.Min.X);
                double minY = Math.Max(a.Min.Y, b.Min.Y);
                double minZ = Math.Max(a.Min.Z, b.Min.Z);

                double maxX = Math.Min(a.Max.X, b.Max.X);
                double maxY = Math.Min(a.Max.Y, b.Max.Y);
                double maxZ = Math.Min(a.Max.Z, b.Max.Z);

                bool hit = (minX <= maxX) && (minY <= maxY) && (minZ <= maxZ);
                center = hit ? new XYZ((minX + maxX) / 2,
                                       (minY + maxY) / 2,
                                       (minZ + maxZ) / 2) : null;
                return hit;
            }

            IEnumerable<Element>[] groups = { pipes, ducts, trays };

            foreach (Wall w in walls)
            {
                var wBox = w.get_BoundingBox(null);
                if (wBox == null) continue;

                foreach (var g in groups)
                {
                    foreach (Element el in g)
                    {
                        var b = el.get_BoundingBox(null);
                        if (b == null) continue;

                        if (BoxesIntersect(wBox, b, out XYZ center))
                        {
                            hostWall = w;
                            clashPt = center;
                            goto FOUND;
                        }
                    }
                }
            }
        FOUND:

            if (clashPt == null || hostWall == null)
            {
                MessageBox.Show("Пересечений не найдено.",
                                "Результат",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // 4. Находим грань стены, ближайшую к точке
            Options opts = new Options { ComputeReferences = true };
            Face targetFace = null;
            XYZ pointOnFace = null;
            UV uvOnFace = null;

            foreach (GeometryObject go in hostWall.get_Geometry(opts))
            {
                if (go is Solid solid)
                {
                    foreach (Face f in solid.Faces)
                    {
                        var proj = f.Project(clashPt);
                        if (proj != null)
                        {
                            targetFace = f;
                            pointOnFace = proj.XYZPoint;
                            uvOnFace = proj.UVPoint;
                            break;
                        }
                    }
                }
                if (targetFace != null) break;
            }

            if (targetFace == null)
            {
                MessageBox.Show("Не удалось найти плоскость стены для размещения.",
                                "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // 5. Вычисляем нормаль и refDir
            XYZ normal = targetFace.ComputeNormal(uvOnFace).Normalize();

            // сдвиг на 1 мм наружу, чтобы точка гарантированно была «над» гранью
            XYZ placePt = pointOnFace + normal * (1.0 / 304.8); // 1 мм в футах

            // referenceDirection должен лежать в плоскости грани
            XYZ refDir = normal.CrossProduct(XYZ.BasisZ);
            if (refDir.GetLength() < 1e-9)     // если нормаль почти || Z
                refDir = normal.CrossProduct(XYZ.BasisX);
            refDir = refDir.Normalize();

            // 6. Вставляем экземпляр семейства
            using (Transaction t = new Transaction(doc, "Place Face-based Family"))
            {
                t.Start();
                if (!symbol.IsActive) symbol.Activate();

                doc.Create.NewFamilyInstance(
                    targetFace.Reference,   // Reference грани-хоста
                    placePt,                // точка на грани
                    refDir,                 // направление в плоскости
                    symbol);

                t.Commit();
            }

            MessageBox.Show("Семейство размещено на грани стены.",
                            "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
