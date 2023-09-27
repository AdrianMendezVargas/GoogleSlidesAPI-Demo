namespace GoogleSlides.Api.Models
{
    public class FolderMetadata
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public List<FolderMetadata> Subfolders { get; set; }
        public List<SlideItem> SlideItems { get; set; }


        public FolderMetadata()
        {
            Subfolders = new List<FolderMetadata>();
            SlideItems = new List<SlideItem>();
        }

        public class SlideItem
        {
            public string Id => "m" + Guid.NewGuid().ToString();
            public string SlideId { get; set; }
            public string PresentationId { get; set; }
            public string PresentationName { get; set; }
            public string Name { get; set; }
            public string ThumbnailUrl { get; set; }
        }
    }
}
