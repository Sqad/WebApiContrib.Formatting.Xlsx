namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Models
{
    public class RowDefinition
    {
        public int OriginalRowIndex { get; set; }

        public int StartExcelRowIndex { get; set; }
        public int EndExcelRowIndex { get; set; }
        public int PrimaryExcelRowIndex { get; set; }

        public int AboveCount { get; set; }
        public int BelowCount { get; set; }
    }
}
