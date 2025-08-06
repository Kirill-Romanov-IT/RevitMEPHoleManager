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

            // ── 0. Проверяем выбор семейства ──
            if (FamilyCombo.SelectedValue == null)
            {
                MessageBox.Show("Сначала выберите face-based семейство Generic Model.");
                return;
            }

            ElementId familyId = (ElementId)FamilyCombo.SelectedValue;
            Family family = doc.GetElement(familyId) as Family;
            FamilySymbol baseSymbol =
                doc.GetElement(family.GetFamilySymbolIds().First()) as FamilySymbol;

            // ── 1. Получаем/создаём тип «Копия1» ──
            const string NEW_TYPE = "Копия1";
            FamilySymbol holeSym = family
                .GetFamilySymbolIds()
                .Select(id => doc.GetElement(id) as FamilySymbol)
                .FirstOrDefault(s => s.Name.Equals(NEW_TYPE, StringComparison.OrdinalIgnoreCase));

            if (holeSym == null)
            {
                using (Transaction t = new Transaction(doc, "Duplicate type"))
                {
                    t.Start();
                    holeSym = baseSymbol.Duplicate(NEW_TYPE) as FamilySymbol; // Duplicate возвращает объект-тип
                    t.Commit();
                }
            }

            // ── 2. Коллекции элементов ──
            var walls = new FilteredElementCollector(doc).OfClass(typeof(Wall)).ToElements();
            var pipes = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).ToElements();
            var ducts = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctCurves).ToElements();
            var trays = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTray).ToElements();
            IEnumerable<Element>[] mepGroups = { pipes, ducts, trays };

            // — вспомогательная проверка BoundingBox —
            bool Hit(BoundingBoxXYZ a, BoundingBoxXYZ b, out XYZ c)
            {
                double minX = Math.Max(a.Min.X, b.Min.X);
                double minY = Math.Max(a.Min.Y, b.Min.Y);
                double minZ = Math.Max(a.Min.Z, b.Min.Z);
                double maxX = Math.Min(a.Max.X, b.Max.X);
                double maxY = Math.Min(a.Max.Y, b.Max.Y);
                double maxZ = Math.Min(a.Max.Z, b.Max.Z);
                bool ok = (minX <= maxX) && (minY <= maxY) && (minZ <= maxZ);
                c = ok ? new XYZ((minX + maxX) / 2, (minY + maxY) / 2, (minZ + maxZ) / 2) : null;
                return ok;
            }

            int placed = 0;
            Options opt = new Options { ComputeReferences = true };

            using (Transaction tr = new Transaction(doc, "Place holes"))
            {
                tr.Start();
                if (!holeSym.IsActive) holeSym.Activate();

                foreach (Wall wall in walls)
                {
                    var wBox = wall.get_BoundingBox(null); if (wBox == null) continue;

                    foreach (var group in mepGroups)
                        foreach (Element mep in group)
                        {
                            var mBox = mep.get_BoundingBox(null); if (mBox == null) continue;
                            if (!Hit(wBox, mBox, out XYZ clashPt)) continue;

                            // ── ищем грань стены ──
                            Face face = null; XYZ pFace = null; UV uv = null;
                            foreach (GeometryObject go in wall.get_Geometry(opt))
                            {
                                if (go is Solid s)
                                    foreach (Face f in s.Faces)
                                    {
                                        var pr = f.Project(clashPt);
                                        if (pr != null) { face = f; pFace = pr.XYZPoint; uv = pr.UVPoint; break; }
                                    }
                                if (face != null) break;
                            }
                            if (face == null) continue;      // безопасность

                            // ── точка, нормаль, refDir ──
                            XYZ normal = face.ComputeNormal(uv).Normalize();
                            XYZ pt = pFace + normal * (1.0 / 304.8);              // +1 мм
                            XYZ refDir = normal.CrossProduct(XYZ.BasisZ);
                            if (refDir.GetLength() < 1e-9) refDir = normal.CrossProduct(XYZ.BasisX);
                            refDir = refDir.Normalize();

                            // ── вставка ──
                            doc.Create.NewFamilyInstance(
                                face.Reference,
                                pt,
                                refDir,
                                holeSym);

                            placed++;
                        }
                }
                tr.Commit();
            }

            MessageBox.Show(
                placed > 0 ? $"Вставлено отверстий: {placed}" : "Пересечений не найдено.",
                "Результат", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
