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
        // ───────────────── Старт ─────────────────
        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            Document doc = _uiApp.ActiveUIDocument.Document;

            // 0. Проверяем выбранное семейство
            if (FamilyCombo.SelectedValue == null)
            {
                MessageBox.Show("Сначала выберите face-based семейство Generic Model.");
                return;
            }

            ElementId famId = (ElementId)FamilyCombo.SelectedValue;
            Family family = doc.GetElement(famId) as Family;
            FamilySymbol baseSym =
                doc.GetElement(family.GetFamilySymbolIds().First()) as FamilySymbol;

            // ── гарантируем тип «Копия1» ──
            const string NEW_TYPE = "Копия1";
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

            // ── 1. собираем стены хоста ──
            IList<Element> wallsHost = new FilteredElementCollector(doc)
                                       .OfClass(typeof(Wall))
                                       .ToElements();

            // ── 2. строим фильтр MEP (Pipe ∨ Duct ∨ Tray) для старого API ──
            ElementFilter fPipe = new ElementCategoryFilter(BuiltInCategory.OST_PipeCurves);
            ElementFilter fDuct = new ElementCategoryFilter(BuiltInCategory.OST_DuctCurves);
            ElementFilter fTray = new ElementCategoryFilter(BuiltInCategory.OST_CableTray);
            ElementFilter mepFilter = new LogicalOrFilter(new LogicalOrFilter(fPipe, fDuct), fTray);

            // локальная функция: все MEP-элементы в документе d
            IEnumerable<Element> CollectMEP(Document d) =>
                new FilteredElementCollector(d).WherePasses(mepFilter).ToElements();

            // ── 3. список (Element, Transform) из хоста и связей ──
            var mepList = new List<(Element elem, Transform tx)>();

            // 3.1 MEP из активной модели
            foreach (Element mep in CollectMEP(doc))
                mepList.Add((mep, Transform.Identity));

            // 3.2 MEP из Revit-связей
            foreach (RevitLinkInstance link in
                     new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>())
            {
                Document linkDoc = link.GetLinkDocument();
                if (linkDoc == null) continue;           // связь не загружена

                Transform lTx = link.GetTransform();     // связь → хост
                foreach (Element mepElem in CollectMEP(linkDoc))
                    mepList.Add((mepElem, lTx));
            }

            // ── 4. вставляем отверстия ──
            int placed = 0;
            Options opt = new Options { ComputeReferences = true };

            // простая проверка пересечения BoundingBox'ов
            bool Intersects(BoundingBoxXYZ a, BoundingBoxXYZ b, out XYZ center)
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

            using (Transaction tr = new Transaction(doc, "Place holes"))
            {
                tr.Start();
                if (!holeSym.IsActive) holeSym.Activate();

                foreach (Wall wall in wallsHost)
                {
                    BoundingBoxXYZ wBox = wall.get_BoundingBox(null);
                    if (wBox == null) continue;

                    foreach ((Element mepElem, Transform tx) in mepList)
                    {
                        BoundingBoxXYZ bb = mepElem.get_BoundingBox(null);
                        if (bb == null) continue;

                        // трансформируем BoundingBox MEP-элемента в координаты хоста
                        BoundingBoxXYZ bbHost = new BoundingBoxXYZ
                        {
                            Min = tx.OfPoint(bb.Min),
                            Max = tx.OfPoint(bb.Max)
                        };

                        if (!Intersects(wBox, bbHost, out XYZ clashPt)) continue;

                        // ищем грань стены
                        Face hostFace = null; XYZ pOnFace = null; UV uv = null;
                        foreach (GeometryObject go in wall.get_Geometry(opt))
                        {
                            if (go is Solid s)
                                foreach (Face f in s.Faces)
                                {
                                    var proj = f.Project(clashPt);
                                    if (proj != null) { hostFace = f; pOnFace = proj.XYZPoint; uv = proj.UVPoint; break; }
                                }
                            if (hostFace != null) break;
                        }
                        if (hostFace == null) continue;

                        // направление и точка вставки
                        XYZ normal = hostFace.ComputeNormal(uv).Normalize();
                        XYZ refDir = normal.CrossProduct(XYZ.BasisZ).GetLength() < 1e-9
                                     ? normal.CrossProduct(XYZ.BasisX)
                                     : normal.CrossProduct(XYZ.BasisZ);
                        refDir = refDir.Normalize();
                        XYZ placePt = pOnFace + normal * (1.0 / 304.8); // +1 мм наружу

                        doc.Create.NewFamilyInstance(hostFace.Reference, placePt, refDir, holeSym);
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
