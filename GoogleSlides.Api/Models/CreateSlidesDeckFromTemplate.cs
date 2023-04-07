using ClassLibrary1;

namespace GoogleSlides.Api.Models
{
    public class CreateSlidesDeckFromTemplate
    {

        public string TemplateId { get; set; }
        public string PresentationName { get; set; }
        public IDictionary<string, string> TextPlaceholders { get; set; }
        public IDictionary<string, string> ImagePlaceholders { get; set; }
        public IDictionary<string, ChartInfo> ChartPlaceholders { get; set; }

        public int[] SlidesToRemove { get; set; }
        public string ReciverEmail { get; set; }


    }

    public class ChartInfo
    {

        public int Id { get; set; } = 0;
        public string Type { get; set; } = "COLUMN";
        public string Title { get; set; } = string.Empty;
        public string Subtitle { get; set; } = string.Empty;
        public string LeftAxisName { get; set; }
        public string BottomAxisName { get; set; }
        public string LegendPosition { get; set; } = "BOTTOM_LEGEND";
        public IDictionary<string, string[]> Domains { get; set; }
        public IDictionary<string, string[]> Series { get; set; }

    }


}
