using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace RevitMEPHoleManager
{
    /// <summary>Подсчитывает, сколько коллизий имеют круглую и прямоугольную геометрию.</summary>
    internal static class IntersectionStats
    {
        /// <summary>
        /// Анализирует пересечения и возвращает количество круглых и прямоугольных.
        /// </summary>
        /// <param name="hosts">Стены и/или перекрытия, в которых ищем отверстия.</param>
        /// <param name="mepList">MEP-элементы и их трансформа к координатам хоста.</param>
        /// <returns>(roundCount, rectangularCount)</returns>
        public static (int roundCnt, int rectCnt) Analyze(
            IEnumerable<Element> hosts,
            IEnumerable<(Element mep, Transform tx)> mepList)
        {
            int roundCnt = 0, rectCnt = 0;

            bool BoxesIntersect(BoundingBoxXYZ a, BoundingBoxXYZ b)
            {
                return !(a.Max.X < b.Min.X || a.Min.X > b.Max.X ||
                         a.Max.Y < b.Min.Y || a.Min.Y > b.Max.Y ||
                         a.Max.Z < b.Min.Z || a.Min.Z > b.Max.Z);
            }

            foreach (Element host in hosts)
            {
                BoundingBoxXYZ hBox = host.get_BoundingBox(null);
                if (hBox == null) continue;

                foreach ((Element mep, Transform tx) in mepList)
                {
                    BoundingBoxXYZ bb = mep.get_BoundingBox(null);
                    if (bb == null) continue;

                    // преобразуем в координаты хоста
                    BoundingBoxXYZ bbHost = new BoundingBoxXYZ
                    {
                        Min = tx.OfPoint(bb.Min),
                        Max = tx.OfPoint(bb.Max)
                    };

                    if (!BoxesIntersect(hBox, bbHost)) continue;

                    // классификация формы
                    switch (mep.Category.Id.IntegerValue)
                    {
                        // трубы всегда считаем круглыми
                        case (int)BuiltInCategory.OST_PipeCurves:
                            roundCnt++;
                            break;

                        // воздуховоды: смотрим тип
                        case (int)BuiltInCategory.OST_DuctCurves:
                            {
                                Element typeElem = mep.Document.GetElement(mep.GetTypeId());
                                if (typeElem is MEPCurveType mType)
                                {
                                    if (mType.Shape == ConnectorProfileType.Round)
                                        roundCnt++;
                                    else
                                        rectCnt++;            // Rectangular, Oval, Oblong → считаем «квадратными»
                                }
                                else rectCnt++;
                                break;
                            }

                        // кабель-лотки
                        case (int)BuiltInCategory.OST_CableTray:
                            rectCnt++;
                            break;
                    }

                }
            }
            return (roundCnt, rectCnt);
        }
    }
}
