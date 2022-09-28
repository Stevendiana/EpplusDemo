using System.Data;

namespace RetrieveExcelData.Models
{
    public class RangeTables
    {
        public string? SheetName { get; set; }
        public string? CellAddress { get; set; }
        public DataTable? ExtractedList { get; set; }
    }
}