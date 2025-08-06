using System;

namespace RevitMEPHoleManager
{
    /// <summary>
    /// Центральное место всех расчётов отверстий.
    /// ───────────────────────────────────────────
    /// • Расчёт окончательных габаритов отверстия
    /// • Формирование человека-читабельного имени типоразмера
    /// </summary>
    internal static class Calculaters
    {
        /// <param name="isRound">true — круглая трасса (труба / круглый воздуховод)</param>
        /// <param name="elemW">Ширина (или Ø) инженерной трассы, мм</param>
        /// <param name="elemH">Высота трассы, мм (для круглой = elemW)</param>
        /// <param name="clearanceMm">Зазор вокруг трассы, мм (в обе стороны!)</param>
        /// <param name="holeW">Выход: ширина отверстия, мм</param>
        /// <param name="holeH">Выход: высота отверстия, мм</param>
        /// <param name="holeTypeName">Выход: строка-имя типоразмера</param>
        public static void GetHoleSize(
            bool isRound,
            double elemW,
            double elemH,
            double clearanceMm,
            out double holeW,
            out double holeH,
            out string holeTypeName)
        {
            // две стороны зазора (+ по 1 с каждой стороны)
            double add = clearanceMm * 2;

            // для читаемости будем «поднимать» до кратности 5 мм
            static double RoundUp5(double v) => Math.Ceiling(v / 5.0) * 5.0;

            if (isRound)
            {
                // политика компании → квадратное отверстие под трубу
                holeW = holeH = RoundUp5(elemW + add);
                holeTypeName = $"Квадр. {holeW}×{holeH}";
            }
            else
            {
                holeW = RoundUp5(elemW + add);
                holeH = RoundUp5(elemH + add);
                holeTypeName = $"Прям. {holeW}×{holeH}";
            }
        }
    }
}
