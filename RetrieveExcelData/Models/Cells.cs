using RetrieveExcelData.Validation;
using System.ComponentModel.DataAnnotations;

namespace RetrieveExcelData.Models
{
    [Serializable]
    public class Cells
    {
        [Required(ErrorMessage = "Cell Address is required.")]
        public string? CellAddress { get; set; }
        
        [Required(ErrorMessage = "Sheet Name is required.")]
        [Display(Name = "Sheet Name")]
        public string? SheetName { get; set; }
        public string? DataType { get; set; }

        [NumberOrText(nameof(DataType), nameof(CellAddress))]
        [Required(ErrorMessage = "Cell value is required.")]
        public string? CellValue { get; set; }
    }
}