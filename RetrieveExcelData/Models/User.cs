using System.ComponentModel.DataAnnotations;

namespace RetrieveExcelData.Models
{
    [Serializable]
    public class User
    {
        [Display(Name = "M.I.")]
        public string MiddleInitial { get; set; }

        [Display(Name = "Suffix")]
        public string NameSuffix { get; set; }

        [Required(AllowEmptyStrings = false)]
        [Display(Name = "First Name")]
        public string FirstName { get; set; }

        [Required(AllowEmptyStrings = false)]
        [Display(Name = "Last Name")]
        public string LastName { get; set; }
    }
}
