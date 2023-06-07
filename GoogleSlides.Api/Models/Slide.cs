using Newtonsoft.Json;

namespace GoogleSlides.Api.Models
{

    public class SlideMetadata
    {
        public SlideMetadata()
        {
            Placeholders = new List<PlaceholderMetadata>();
        }

        public int Id { get; set; }
        public int Index { get; set; }
        public bool Removable { get; set; }
        [JsonIgnore]
        [System.Text.Json.Serialization.JsonIgnore]
        public string TemplateId { get; set; } // Foreign key to Template
        [JsonIgnore]
        [System.Text.Json.Serialization.JsonIgnore]
        public TemplateMetadata Template { get; set; } // Navigation property to Template
        public List<PlaceholderMetadata> Placeholders { get; set; }
    }
}
