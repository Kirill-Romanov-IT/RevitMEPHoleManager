using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace RevitMEPHoleManager
{
    /// <summary>
    /// Одна строка для вывода в <see cref="MainWindow.StatsGrid"/>.
    /// </summary>
    internal sealed class IntersectRow
    {
        public string Host { get; set; }  // «Стена» / «Перекрытие»
        public int HostId { get; set; }
        public int MepId { get; set; }
        public string Shape { get; set; }  // «круглая» / «квадратная»
        public double Size1 { get; set; }  // Ø  или  ширина, мм
        public double Size2 { get; set; }  //       высота,  мм (для прямоугольных)
    }

    /// <summary>
    /// Подсчитывает количество круглых / прямоугольных
    /// пересечений отдельно для стен и перекрытий,
    /// а также формирует подробную таблицу для UI.
    /// </summary>
    internal static class IntersectionStats
    {
        /// <returns>
        /// (wallRound, wallRect, floorRound, floorRect, rows)
        /// </returns>
        public static (int wRnd, int wRec, int fRnd, int fRec, List<IntersectRow> rows) Analyze(
            IEnumerable<Element> hosts,                       // стены + плиты
            IEnumerable<(Element elem, Transform tx)> mepList // MEP + их трансформы
        )
        {
            int wRnd = 0, wRec = 0, fRnd = 0, fRec = 0;
            var rows = new List<IntersectRow>();

            bool Intersects(BoundingBoxXYZ a, BoundingBoxXYZ b)
            {
                return !(a.Max.X < b.Min.X || a.Min.X > b.Max.X ||
                         a.Max.Y < b.Min.Y || a.Min.Y > b.Max.Y ||
                         a.Max.Z < b.Min.Z || a.Min.Z > b.Max.Z);
            }

            foreach (Element host in hosts)
            {
                BoundingBoxXYZ hBox = host.get_BoundingBox(null);
                if (hBox == null) continue;

                bool isWall = host is Wall;
                bool isFloor = host is Floor;
                string hostLabel = isWall ? "Стена" :
                                   isFloor ? "Перекрытие" : host.Category?.Name ?? "Host";

                foreach ((Element mep, Transform tx) in mepList)
                {
                    BoundingBoxXYZ bb = mep.get_BoundingBox(null);
                    if (bb == null) continue;

                    // BoundingBox MEP → координаты хоста
                    var bbHost = new BoundingBoxXYZ
                    {
                        Min = tx.OfPoint(bb.Min),
                        Max = tx.OfPoint(bb.Max)
                    };

                    if (!Intersects(hBox, bbHost)) continue;

                    bool isRound = false; // форма
                    double s1 = 0, s2 = 0;

                    switch ((BuiltInCategory)mep.Category.Id.IntegerValue)
                    {
                        case BuiltInCategory.OST_PipeCurves:
                            // трубы – круглые
                            if (SizeHelper.TryGetRoundDiameter(mep, out double d))
                            {
                                isRound = true;
                                s1 = d;
                            }
                            break;

                        case BuiltInCategory.OST_DuctCurves:
                            {
                                var type = mep.Document.GetElement(mep.GetTypeId()) as MEPCurveType;
                                if (type?.Shape == ConnectorProfileType.Round)
                                {
                                    isRound = true;
                                    if (SizeHelper.TryGetRoundDiameter(mep, out double dDuct))
                                        s1 = dDuct;
                                }
                                else
                                {
                                    if (SizeHelper.TryGetRectSize(mep, out double w, out double h))
                                    {
                                        s1 = w; s2 = h;
                                    }
                                }
                                break;
                            }

                        case BuiltInCategory.OST_CableTray:
                            // лотки – «квадратные»
                            if (SizeHelper.TryGetRectSize(mep, out double wT, out double hT))
                            {
                                s1 = wT; s2 = hT;
                            }
                            break;
                    }

                    // суммируем статистику
                    if (isWall)
                        if (isRound) wRnd++; else wRec++;
                    if (isFloor)
                        if (isRound) fRnd++; else fRec++;

                    // добавляем строку для DataGrid
                    rows.Add(new IntersectRow
                    {
                        Host = hostLabel,
                        HostId = host.Id.IntegerValue,
                        MepId = mep.Id.IntegerValue,
                        Shape = isRound ? "круглая" : "квадратная",
                        Size1 = s1,
                        Size2 = isRound ? 0 : s2
                    });
                }
            }

            return (wRnd, wRec, fRnd, fRec, rows);
        }
    }
}
