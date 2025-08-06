using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitMEPHoleManager
{
    /// <summary>
    /// Сводная строка: сколько круглых / квадратных пересечений
    /// для одного хост-элемента (стены или перекрытия).
    /// </summary>
    internal class HostStatRow
    {
        public int HostId { get; set; }   // Id стены / плиты
        public string HostName { get; set; }   // «Стена» / «Перекрытие»
        public int Round { get; set; }   // кол-во круглых
        public int Rect { get; set; }   // кол-во квадратных
    }
}

