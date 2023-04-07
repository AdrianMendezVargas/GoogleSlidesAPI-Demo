using System.ComponentModel.DataAnnotations;

namespace GoogleSlides.Api.Models.Domain
{
    public class AuthConfig
    {
        [Key]
        public int Id { get; set; }
        public string Token { get; set; }
        public string RefreshToken { get; set; }
        public DateTime IssueTime { get; set; }
        public int DurationInSeconds { get; set; }
    }
}
