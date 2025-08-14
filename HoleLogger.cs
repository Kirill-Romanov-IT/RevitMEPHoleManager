using System.Text;

namespace RevitMEPHoleManager
{
    /// <summary>
    /// Класс для логирования операций с отверстиями
    /// </summary>
    public sealed class HoleLogger
    {
        private readonly StringBuilder sb = new();
        
        /// <summary>
        /// Добавляет строку в лог
        /// </summary>
        public void Add(string line) => sb.AppendLine(line);
        
        /// <summary>
        /// Добавляет горизонтальную линию разделитель
        /// </summary>
        public void HR() => sb.AppendLine(new string('─', 70));
        
        /// <summary>
        /// Возвращает весь лог как строку
        /// </summary>
        public override string ToString() => sb.ToString();
        
        /// <summary>
        /// Очищает лог
        /// </summary>
        public void Clear() => sb.Clear();
        
        /// <summary>
        /// Проверяет, пуст ли лог
        /// </summary>
        public bool IsEmpty => sb.Length == 0;
        
        /// <summary>
        /// Получает количество строк в логе
        /// </summary>
        public int LineCount => sb.ToString().Split('\n').Length;
    }
}
