using System.ComponentModel.DataAnnotations;

namespace RetrieveExcelData.Models
{
    [Serializable]
    public class OutputCells
    {

        [Required(ErrorMessage = "Range is required.")]
        [RegularExpression(@"^([a-zA-Z]|[a-zA-Z][a-zA-Z]+)(\d+)$:(\1\d+|[a-zA-Z]|[a-zA-Z][a-zA-Z]+\2)|([a-zA-Z]|[a-zA-Z][a-zA-Z]+)(\d+)$", ErrorMessage = "Invaid range format.")]
        //[RegularExpression(@"^([a-zA-Z]{2}\d{1,3}:[a-zA-Z]{2}\d{1,3})|([a-zA-Z]{2}\d{1,3})$", ErrorMessage = "Invaid range format.")]
        public string? Range { get; set; }
        
        [Required(ErrorMessage = "Sheet Name is required.")]
        [Display(Name = "Sheet Name")]
        public string? OutputSheetName { get; set; }
    }
}
