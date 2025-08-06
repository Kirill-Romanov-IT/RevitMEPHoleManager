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

            // 0. Проверяем выбранное семейство
            if (FamilyCombo.SelectedValue == null)
            {
                MessageBox.Show("Сначала выберите face-based семейство Generic Model.");
                return;
            }

            // ── гарантируем тип «Копия1» ──
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

            // ── 1. Хост-элементы: стены + перекрытия ──
            IList<Element> hostElems = new FilteredElementCollector(doc)
                .WherePasses(new LogicalOrFilter(
                                new ElementClassFilter(typeof(Wall)),
                                new ElementClassFilter(typeof(Floor))))   // добавили Floor
                .ToElements();

            // ── 2. Собираем MEP из хоста и всех связей ──
            ElementFilter fPipe = new ElementCategoryFilter(BuiltInCategory.OST_PipeCurves);
            ElementFilter fDuct = new ElementCategoryFilter(BuiltInCategory.OST_DuctCurves);
            ElementFilter fTray = new ElementCategoryFilter(BuiltInCategory.OST_CableTray);
            ElementFilter mepFilter = new LogicalOrFilter(new LogicalOrFilter(fPipe, fDuct), fTray);

            IEnumerable<Element> CollectMEP(Document d) =>
                new FilteredElementCollector(d).WherePasses(mepFilter).ToElements();

            var mepList = new List<(Element elem, Transform tx)>();
            foreach (Element mep in CollectMEP(doc))                                           // из активной модели
                mepList.Add((mep, Transform.Identity));

            foreach (RevitLinkInstance link in
                     new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>())
            {
                Document lDoc = link.GetLinkDocument();
                if (lDoc == null) continue;                    // ссылка не загружена

                Transform lTx = link.GetTransform();           // связь → хост
                foreach (Element mep in CollectMEP(lDoc))
                    mepList.Add((mep, lTx));
            }

            // ── 3. Пересечения и вставка отверстий ──
            bool Intersects(BoundingBoxXYZ a, BoundingBoxXYZ b, out XYZ ctr)
            {
                double minX = Math.Max(a.Min.X, b.Min.X);
                double minY = Math.Max(a.Min.Y, b.Min.Y);
                double minZ = Math.Max(a.Min.Z, b.Min.Z);
                double maxX = Math.Min(a.Max.X, b.Max.X);
                double maxY = Math.Min(a.Max.Y, b.Max.Y);
                double maxZ = Math.Min(a.Max.Z, b.Max.Z);
                bool hit = (minX <= maxX) && (minY <= maxY) && (minZ <= maxZ);
                ctr = hit ? new XYZ((minX + maxX) / 2, (minY + maxY) / 2, (minZ + maxZ) / 2) : null;
                return hit;
            }

            int placed = 0;
            Options opt = new Options { ComputeReferences = true };

            using (Transaction tr = new Transaction(doc, "Place holes"))
            {
                tr.Start();
                if (!holeSym.IsActive) holeSym.Activate();

                foreach (Element host in hostElems)
                {
                    BoundingBoxXYZ hBox = host.get_BoundingBox(null); if (hBox == null) continue;

                    foreach ((Element mepElem, Transform tx) in mepList)
                    {
                        BoundingBoxXYZ bb = mepElem.get_BoundingBox(null); if (bb == null) continue;

                        BoundingBoxXYZ bbHost = new BoundingBoxXYZ          // bb → координаты хоста
                        {
                            Min = tx.OfPoint(bb.Min),
                            Max = tx.OfPoint(bb.Max)
                        };

                        if (!Intersects(hBox, bbHost, out XYZ clashPt)) continue;

                        // ищем грань (работает и для стены, и для плиты)
                        Face face = null; XYZ pOnFace = null; UV uv = null;
                        foreach (GeometryObject go in host.get_Geometry(opt))
                        {
                            if (go is Solid s)
                                foreach (Face f in s.Faces)
                                {
                                    var pr = f.Project(clashPt);
                                    if (pr != null) { face = f; pOnFace = pr.XYZPoint; uv = pr.UVPoint; break; }
                                }
                            if (face != null) break;
                        }
                        if (face == null) continue;

                        XYZ normal = face.ComputeNormal(uv).Normalize();
                        XYZ refDir = normal.CrossProduct(XYZ.BasisZ).GetLength() < 1e-9
                                     ? normal.CrossProduct(XYZ.BasisX)
                                     : normal.CrossProduct(XYZ.BasisZ);
                        refDir = refDir.Normalize();

                        XYZ placePt = pOnFace + normal * (1.0 / 304.8);   // +1 мм наружу

                        doc.Create.NewFamilyInstance(face.Reference, placePt, refDir, holeSym);
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
