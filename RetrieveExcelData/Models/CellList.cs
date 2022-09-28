using System.ComponentModel.DataAnnotations;
using System.Data;

namespace RetrieveExcelData.Models
{
    [Serializable]
    public class CellList
    {
        public string ImageMessage { get; set; } = "Drop File Here";
        public string FileName { get; set; }
        public FileInfo File { get; set; }
        public List<RangeTables> ExtractedTables { get; set; } = new();
        public List<SingleCellsTable> ExtractedCells { get; set; } = new();
        public List<Cells> InputData { get; set; } = new();
        public List<OutputCells> OutputData { get; set; } = new();

        public CellList()
        {
            ExtractedTables = new List<RangeTables>();
            ExtractedCells = new List<SingleCellsTable>();
            InputData = new List<Cells>();
            OutputData = new List<OutputCells>();
        }
    }
}
