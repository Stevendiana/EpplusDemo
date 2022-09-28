using System.ComponentModel.DataAnnotations;

namespace RetrieveExcelData.Models
{
    [Serializable]
    public class MultiAddUser
    {
        [Required]
        [Display(Name = "Start Date")]
        public DateTime StartDate { get; set; }
        public List<User> Users { get; set; }

        public MultiAddUser()
        {
            Users = new List<User>();
        }
    }
}
