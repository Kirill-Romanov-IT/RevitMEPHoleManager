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
            PopulateGenericModelFamilies();
        }

        // ───────────────── ComboBox ─────────────────
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
        }

        // ───────────────── Старт ─────────────────
        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            Document doc = _uiApp.ActiveUIDocument.Document;

            // 0. Проверяем семейство
            if (FamilyCombo.SelectedValue == null)
            {
                MessageBox.Show("Сначала выберите face-based семейство Generic Model.");
                return;
            }

            ElementId familyId = (ElementId)FamilyCombo.SelectedValue;
            Family family = doc.GetElement(familyId) as Family;

            // берём первый существующий тип как «образец»
            FamilySymbol baseSymbol =
                doc.GetElement(family.GetFamilySymbolIds().First()) as FamilySymbol;

            // ❶ ───── Дублируем/получаем тип «Копия1» ─────
            const string NEW_TYPE_NAME = "Копия1";
            FamilySymbol symbolCopy = null;

            foreach (ElementId id in family.GetFamilySymbolIds())
            {
                var sym = doc.GetElement(id) as FamilySymbol;
                if (sym.Name.Equals(NEW_TYPE_NAME, StringComparison.OrdinalIgnoreCase))
                {
                    symbolCopy = sym;     // тип уже есть
                    break;
                }
            }

            if (symbolCopy == null)           // тип «Копия1» ещё не существует
            {
                using (Transaction t = new Transaction(doc, "Duplicate type"))
                {
                    t.Start();

                    // Duplicate теперь возвращает сам объект нового типа
                    FamilySymbol newSym = baseSymbol.Duplicate(NEW_TYPE_NAME) as FamilySymbol;

                    symbolCopy = newSym;              // используем его дальше (Id доступен как newSym.Id)

                    t.Commit();
                }
            }
            // ───────────────────────────────────────────

            // 1. Собираем стены и инженерные элементы
            var walls = new FilteredElementCollector(doc).OfClass(typeof(Wall)).ToElements();
            var pipes = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).ToElements();
            var ducts = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctCurves).ToElements();
            var trays = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTray).ToElements();

            // 2. Ищем пересечение
            XYZ clashPt = null; Wall hostWall = null;
            bool BoxesIntersect(BoundingBoxXYZ a, BoundingBoxXYZ b, out XYZ c)
            {
                double minX = Math.Max(a.Min.X, b.Min.X);
                double minY = Math.Max(a.Min.Y, b.Min.Y);
                double minZ = Math.Max(a.Min.Z, b.Min.Z);
                double maxX = Math.Min(a.Max.X, b.Max.X);
                double maxY = Math.Min(a.Max.Y, b.Max.Y);
                double maxZ = Math.Min(a.Max.Z, b.Max.Z);
                bool hit = (minX <= maxX) && (minY <= maxY) && (minZ <= maxZ);
                c = hit ? new XYZ((minX + maxX) / 2, (minY + maxY) / 2, (minZ + maxZ) / 2) : null;
                return hit;
            }

            IEnumerable<Element>[] groups = { pipes, ducts, trays };
            foreach (Wall w in walls)
            {
                var wBox = w.get_BoundingBox(null); if (wBox == null) continue;
                foreach (var g in groups)
                    foreach (Element el in g)
                    {
                        var b = el.get_BoundingBox(null); if (b == null) continue;
                        if (BoxesIntersect(wBox, b, out XYZ cpt))
                        { hostWall = w; clashPt = cpt; goto FOUND; }
                    }
            }
        FOUND:
            if (clashPt == null)
            {
                MessageBox.Show("Пересечений не найдено.");
                return;
            }

            // 3. Грань стены → Reference + точка
            Options opt = new Options { ComputeReferences = true };
            Face face = null; XYZ pFace = null; UV uv = null;
            foreach (GeometryObject go in hostWall.get_Geometry(opt))
            {
                if (go is Solid s)
                    foreach (Face f in s.Faces)
                    {
                        var pr = f.Project(clashPt);
                        if (pr != null) { face = f; pFace = pr.XYZPoint; uv = pr.UVPoint; break; }
                    }
                if (face != null) break;
            }
            if (face == null)
            {
                MessageBox.Show("Не удалось найти плоскость стены.");
                return;
            }

            XYZ normal = face.ComputeNormal(uv).Normalize();
            XYZ placePt = pFace + normal * (1.0 / 304.8);      // +1 мм
            XYZ refDir = normal.CrossProduct(XYZ.BasisZ);
            if (refDir.GetLength() < 1e-9)
                refDir = normal.CrossProduct(XYZ.BasisX);
            refDir = refDir.Normalize();

            // 4. Вставка экземпляра типа «Копия1»
            using (Transaction t = new Transaction(doc, "Place Copy1 instance"))
            {
                t.Start();
                if (!symbolCopy.IsActive) symbolCopy.Activate();
                doc.Create.NewFamilyInstance(face.Reference, placePt, refDir, symbolCopy);
                t.Commit();
            }

            MessageBox.Show($"Создан экземпляр типа «{NEW_TYPE_NAME}».",
                            "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
