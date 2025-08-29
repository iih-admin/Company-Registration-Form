using System.ComponentModel.DataAnnotations;

namespace CustomerInfo.Model
{
    public class FormSubmission
    {
        [Required]
        public string FullName { get; set; } = "";
        [Required]
        public string PhoneNumber { get; set; } = "";
        [MaxLength(200)]
        public string? Position { get; set; } = "";
        [MaxLength(200)]
        public string? Company { get; set; } = "";
    }
}
