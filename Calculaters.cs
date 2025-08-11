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
            // Вектор оси трубы в локальных координатах хоста (Right-Up-Normal)
            // axisLocal.X = проекция на ось Right (вдоль стены)
            // axisLocal.Y = проекция на ось Up (вертикаль)
            // axisLocal.Z = проекция на ось Normal (через стену)

            // Защита от деления на ноль
            double axX = Math.Abs(axisLocal.X);
            double axY = Math.Abs(axisLocal.Y);
            double axZ = Math.Max(1e-3, Math.Abs(axisLocal.Z));

            if (isRound)
            {
                // Для круглой трубы: проекция круга на плоскость стены даёт эллипс
                // Размеры эллипса зависят от углов наклона в обеих плоскостях
                
                // По оси Right (X): учитываем наклон в плоскости XZ
                double cosAlphaX = axZ / Math.Sqrt(axX * axX + axZ * axZ);
                cosAlphaX = Math.Max(1e-3, cosAlphaX);
                
                // По оси Up (Y): учитываем наклон в плоскости YZ  
                double cosAlphaY = axZ / Math.Sqrt(axY * axY + axZ * axZ);
                cosAlphaY = Math.Max(1e-3, cosAlphaY);
                
                holeWmm = elemWmm / cosAlphaX + 2 * clearanceMm;  // ширина эллипса
                holeHmm = elemWmm / cosAlphaY + 2 * clearanceMm;  // высота эллипса
            }
            else   // прямоугольная/квадратная секция
            {
                // Для прямоугольного сечения нужно учесть поворот прямоугольника в плоскости стены
                // Консервативный подход: рассчитываем габариты повёрнутого прямоугольника
                
                // Проекция осей прямоугольного сечения на плоскость стены
                // Ось сечения (вдоль трубы) проектируется как axisLocal
                // Поперечные оси сечения нужно спроектировать на плоскость XY
                
                // Нормализованная проекция оси трубы на плоскость стены (XY)
                double projLength = Math.Sqrt(axX * axX + axY * axY);
                projLength = Math.Max(1e-6, projLength);
                
                double unitX = axX / projLength;  // единичный вектор проекции по X
                double unitY = axY / projLength;  // единичный вектор проекции по Y
                
                // Поперечный вектор к проекции (повёрнутый на 90°)
                double perpX = -unitY;
                double perpY = unitX;
                
                // Габариты повёрнутого прямоугольника в плоскости стены
                // Для каждой оси (X,Y) находим максимальную проекцию углов прямоугольника
                double hw = elemWmm / 2;  // полуширина сечения
                double hh = elemHmm / 2;  // полувысота сечения
                
                // Углы прямоугольного сечения при проекции на плоскость стены:
                double projW = Math.Abs(hw * unitX) + Math.Abs(hh * perpX);
                double projH = Math.Abs(hw * unitY) + Math.Abs(hh * perpY);
                
                // Учитываем также увеличение из-за наклона к плоскости стены
                double cosTheta = axZ;  // cos угла между осью трубы и нормалью стены
                cosTheta = Math.Max(1e-3, cosTheta);
                
                holeWmm = 2 * projW / cosTheta + 2 * clearanceMm;
                holeHmm = 2 * projH / cosTheta + 2 * clearanceMm;
            }
        }
    }
}
