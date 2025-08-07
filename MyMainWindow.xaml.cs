using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Data;
using System.ComponentModel;
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

            //── 1. хост‑элементы: стены + перекрытия ──
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

            //── 3. анализ пересечений ──
            double clearance = 50;
            if (double.TryParse(ClearanceBox.Text, out double cTmp) && cTmp > 0)
                clearance = cTmp;

            var (wRnd, wRec, fRnd, fRec, clashList, hostStats) =
                IntersectionStats.Analyze(hostElems, mepList, clearance);

            //─────────────────────────────────────────────
            // 3‑A. ОБЪЕДИНЯЕМ ОТВЕРСТИЯ (новый код)
            //─────────────────────────────────────────────
            bool mergeOn = EnableMergeChk.IsChecked == true;
            double mergeDist = 0;
            if (mergeOn && double.TryParse(MergeDistBox.Text, out double mTmp) && mTmp > 0)
                mergeDist = mTmp;       // мм

            if (mergeOn && mergeDist > 0)
                clashList = MergeService.Merge(clashList, mergeDist, clearance).ToList();
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
            TaskDialog.Show("Статистика пересечений",
                $"В стенах:\n  • круглые   — {wRnd}\n  • квадратные — {wRec}\n\n" +
                $"В перекрытиях:\n  • круглые   — {fRnd}\n  • квадратные — {fRec}");

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

                    XYZ clashPt = row.Center;
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

                    XYZ placePt = pOnFace;                             // прямо на грань

                    // ➊ вставляем face-based экземпляр базовым символом (Копия1)
                    FamilyInstance inst = doc.Create.NewFamilyInstance(face.Reference,
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

            MessageBox.Show($"Вставлено отверстий: {placed}",
                            "Результат", MessageBoxButton.OK, MessageBoxImage.Information);
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
    }
}
