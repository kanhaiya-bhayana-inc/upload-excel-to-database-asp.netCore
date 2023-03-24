using System.ComponentModel.DataAnnotations;

namespace CheckExcel.Models
{
    public class Customer
    {
        [Key]

        public int id { get; set; }
        public string firstName { get; set; }
        public string lastName { get; set; }
        public string job { get; set; }
        public float amount { get; set; }
        public DateTime tdate { get; set; }
    }
}
