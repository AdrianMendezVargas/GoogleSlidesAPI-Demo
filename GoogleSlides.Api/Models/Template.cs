using System.ComponentModel.DataAnnotations;

namespace GoogleSlides.Api.Models
{
    public class TemplateMetadata
    {

        public TemplateMetadata()
        {
            Slides = new List<SlideMetadata>();
        }

        [Key]
        public string Id { get; set; }
        public string Name { get; set; }
        public List<SlideMetadata> Slides { get; set; }
    }
}
