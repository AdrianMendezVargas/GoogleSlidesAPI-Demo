using Newtonsoft.Json;

namespace GoogleSlides.Api.Models
{
    public class PlaceholderMetadata
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public int MaxLength { get; set; }
        public bool Editable { get; set; }
        public bool Removable { get; set; }
        public string BindedTo { get; set; }
        [JsonIgnore]
        [System.Text.Json.Serialization.JsonIgnore]
        public string SlideId { get; set; } // Foreign key to Slide
        [JsonIgnore]
        [System.Text.Json.Serialization.JsonIgnore]
        public SlideMetadata Slide { get; set; } // Navigation property to Slide
    }
}
