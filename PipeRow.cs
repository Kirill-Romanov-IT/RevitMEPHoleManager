namespace RevitMEPHoleManager
{
    /// <summary>
    /// Информация о трубе для отображения в таблице
    /// </summary>
    internal class PipeRow
    {
        public int Id { get; set; }          // ElementId
        public string System { get; set; }   // система/тип
        public double DN { get; set; }       // номинальный Ø, мм
        public double Length { get; set; }   // длина в мм
        public double LengthM => Length / 1000.0;  // длина в метрах
        public string Level { get; set; }    // уровень
        public string Status { get; set; }   // статус размеров
    }

    /// <summary>
    /// Информация о воздуховоде для отображения в таблице
    /// </summary>
    internal class DuctRow
    {
        public int Id { get; set; }
        public string System { get; set; }
        public string Shape { get; set; }    // Round/Rect/Square
        public string Size { get; set; }     // размеры (Ø200 или 400×300)
        public double Length { get; set; }   // длина в мм
        public double LengthM => Length / 1000.0;  // длина в метрах
        public string Level { get; set; }
        public string Status { get; set; }   // статус получения размеров
    }

    /// <summary>
    /// Информация о кабельном лотке для отображения в таблице
    /// </summary>
    internal class TrayRow
    {
        public int Id { get; set; }
        public string System { get; set; }
        public string Size { get; set; }     // размеры ШxВ
        public double Length { get; set; }   // длина в мм
        public double LengthM => Length / 1000.0;  // длина в метрах
        public string Level { get; set; }
        public string Status { get; set; }
    }

    /// <summary>
    /// Информация о стене для отображения в таблице
    /// </summary>
    internal class WallRow
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public double ThicknessMm { get; set; }  // толщина в мм
        public double AreaM2 { get; set; }       // площадь в м²
        public string Level { get; set; }
    }
}
