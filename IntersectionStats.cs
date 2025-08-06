using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace RevitMEPHoleManager
{
    /// <summary>
    /// Подсчитывает количество круглых / прямоугольных пересечений
    /// отдельно для стен и перекрытий.
    /// </summary>
    internal static class IntersectionStats
    {
        /// <returns>
        /// (wallRound, wallRect, floorRound, floorRect)
        /// </returns>
        public static (int wRnd, int wRec, int fRnd, int fRec) Analyze(
            IEnumerable<Element> hosts,                       // стены + плиты
            IEnumerable<(Element elem, Transform tx)> mepList // MEP + их трансформы
        )
        {
            int wRnd = 0, wRec = 0, fRnd = 0, fRec = 0;

            bool Intersects(BoundingBoxXYZ a, BoundingBoxXYZ b)
            {
                return !(a.Max.X < b.Min.X || a.Min.X > b.Max.X ||
                         a.Max.Y < b.Min.Y || a.Min.Y > b.Max.Y ||
                         a.Max.Z < b.Min.Z || a.Min.Z > b.Max.Z);
            }

            foreach (Element host in hosts)
            {
                var hBox = host.get_BoundingBox(null);
                if (hBox == null) continue;

                bool isWall = host is Wall;
                bool isFloor = host is Floor;

                foreach ((Element mep, Transform tx) in mepList)
                {
                    var bb = mep.get_BoundingBox(null);
                    if (bb == null) continue;

                    // BoundingBox MEP → координаты хоста
                    var bbHost = new BoundingBoxXYZ
                    {
                        Min = tx.OfPoint(bb.Min),
                        Max = tx.OfPoint(bb.Max)
                    };

                    if (!Intersects(hBox, bbHost)) continue;

                    bool isRound = false; // классификация формы

                    switch ((BuiltInCategory)mep.Category.Id.IntegerValue)
                    {
                        case BuiltInCategory.OST_PipeCurves:
                            isRound = true;               // трубы – круглые
                            break;

                        case BuiltInCategory.OST_DuctCurves:
                            {
                                var type = mep.Document.GetElement(mep.GetTypeId()) as MEPCurveType;
                                isRound = type?.Shape == ConnectorProfileType.Round;
                                break;
                            }

                        case BuiltInCategory.OST_CableTray:
                            isRound = false;              // лотки – «квадратные»
                            break;
                    }

                    if (isWall)
                        if (isRound) wRnd++; else wRec++;
                    if (isFloor)
                        if (isRound) fRnd++; else fRec++;
                }
            }
            return (wRnd, wRec, fRnd, fRec);
        }
    }
}
