namespace RevitMEPHoleManager
{
    internal class PipeRow
    {
        public int Id { get; set; }          // ElementId
        public string System { get; set; }          // система/тип
        public double DN { get; set; }          // номинальный Ø, мм
        public double Length { get; set; }          // опционально
    }
}
