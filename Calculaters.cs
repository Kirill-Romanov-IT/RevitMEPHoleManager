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
            // ДИАГНОСТИКА: логируем входные данные
            System.Diagnostics.Debug.WriteLine($"GetHoleSizeIncline: isRound={isRound}, elemW={elemWmm:F0}mm, elemH={elemHmm:F0}mm, clearance={clearanceMm:F0}mm");
            System.Diagnostics.Debug.WriteLine($"  axisLocal=({axisLocal.X:F3}, {axisLocal.Y:F3}, {axisLocal.Z:F3})");
            
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
                // УПРОЩЕННЫЙ подход для прямоугольных воздуховодов
                // Для большинства случаев воздуховоды идут горизонтально и перпендикулярно стенам
                
                // Учитываем наклон только если труба сильно наклонена
                double cosTheta = Math.Abs(axZ);  // cos угла между осью трубы и нормалью стены
                cosTheta = Math.Max(0.5, cosTheta);  // ограничиваем минимальное значение
                
                // Простой расчет с учетом наклона
                holeWmm = elemWmm / cosTheta + 2 * clearanceMm;
                holeHmm = elemHmm / cosTheta + 2 * clearanceMm;
                
                System.Diagnostics.Debug.WriteLine($"  Прямоугольный: cosTheta={cosTheta:F3}, элемент {elemWmm:F0}×{elemHmm:F0}");
            }
            
            // ДИАГНОСТИКА: логируем результат
            System.Diagnostics.Debug.WriteLine($"  Результат: holeW={holeWmm:F0}mm, holeH={holeHmm:F0}mm");
        }
    }
}
