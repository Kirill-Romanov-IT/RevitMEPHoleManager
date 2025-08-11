using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Data;
using System.ComponentModel;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Plumbing;   // ← новый using

namespace RevitMEPHoleManager
{
    internal sealed class HoleLogger
    {
        private readonly StringBuilder sb = new();
        public void Add(string line) => sb.AppendLine(line);
        public void HR() => sb.AppendLine(new string('─', 70));
        public override string ToString() => sb.ToString();
    }
    internal class PipeRow
    {
        public int Id { get; set; }
        public string System { get; set; }
        public double DN { get; set; }    // мм
        public double Length { get; set; }    // мм
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

            // выбираем грань, чья нормаль максимально сонаправлена с осью трубы
            var best = refs
                .Select(r =>
                {
                    var face = (PlanarFace)host.GetGeometryObjectFromReference(r);
                    double dot = Math.Abs(face.FaceNormal.Normalize()
                                           .DotProduct(pipeDir.Normalize()));
                    return (r, dot);
                })
                .OrderByDescending(t => t.dot)
                .First().r;

            return best;
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

            if (mergeOn && mergeDist > 0)
            {
                clashList = MergeService.Merge(clashList, mergeDist, clearance, logger).ToList();
                LogBox.Text = logger.ToString();
                LogBox.ScrollToEnd();
            }
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
                    }
                    catch (Exception ex)
                    {
                        continue; // пропускаем, если не удалось найти подходящую грань
                    }

                    // получаем геометрию грани для расчета точки размещения
                    var face = (PlanarFace)host.GetGeometryObjectFromReference(faceRef);
                    var projection = face.Project(clashPt);
                    if (projection == null) continue;
                    
                    XYZ placePt = projection.XYZPoint;
                    
                    // ⬇️ НОВОЕ: не вставлять в проёмы окон/дверей
                    if (IsInDoorOrWindowOpening(doc, host, placePt))
                    {
                        continue; // пропускаем это отверстие
                    }
                    
                    XYZ normal = face.FaceNormal.Normalize();

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
                    
                    // направление размещения: проекция оси трубы на плоскость грани
                    XYZ refDir = row.PipeDir - normal * row.PipeDir.DotProduct(normal);
                    if (refDir.GetLength() < 1e-6)  // труба перпендикулярна грани
                        refDir = normal.CrossProduct(XYZ.BasisZ).GetLength() > 1e-6 
                               ? normal.CrossProduct(XYZ.BasisZ) 
                               : normal.CrossProduct(XYZ.BasisX);
                    refDir = refDir.Normalize();

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
