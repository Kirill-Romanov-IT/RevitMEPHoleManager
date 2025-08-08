using System;
using Autodesk.Revit.DB;

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

        /// <summary>
        /// Габариты отверстия для круглой/прямоугольной секции с учётом уклона.
        /// axisLocal — ось MEP в локальных осях стены (Right/Up/Normal).
        /// </summary>
        /// <param name="isRound">true — круглая трасса (труба / круглый воздуховод)</param>
        /// <param name="elemWmm">Ширина (или Ø) инженерной трассы, мм</param>
        /// <param name="elemHmm">Высота трассы, мм (для круглой = elemWmm)</param>
        /// <param name="clearanceMm">Зазор вокруг трассы, мм</param>
        /// <param name="axisLocal">Единичный вектор оси в локальных координатах (Right-Up-Normal)</param>
        /// <param name="holeWmm">Выход: ширина отверстия, мм</param>
        /// <param name="holeHmm">Выход: высота отверстия, мм</param>
        public static void GetHoleSizeIncline(
            bool isRound,
            double elemWmm,
            double elemHmm,
            double clearanceMm,
            XYZ axisLocal,
            out double holeWmm,
            out double holeHmm)
        {
            // cos θ между осью трассы и нормалью стены (BasisZ)
            double cosTheta = Math.Abs(axisLocal.Z);
            cosTheta = Math.Max(1e-3, cosTheta); // защита деления на 0

            if (isRound)
            {
                // эллипс: мал. ось = D, бол. = D / cosθ
                holeWmm = elemWmm + 2 * clearanceMm;          // по Right
                holeHmm = elemWmm / cosTheta + 2 * clearanceMm; // по Up
            }
            else   // прямоугольная
            {
                // проекция прямоуг. (консервативно): умножаем большую сторону на 1/cosθ
                double w = Math.Max(elemWmm, elemHmm);
                double h = Math.Min(elemWmm, elemHmm);
                holeWmm = w + 2 * clearanceMm;
                holeHmm = h / cosTheta + 2 * clearanceMm;
            }
        }
    }
}
