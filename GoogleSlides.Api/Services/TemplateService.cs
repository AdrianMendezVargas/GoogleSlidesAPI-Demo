using GoogleSlides.Api.Data;
using GoogleSlides.Api.Models;
using Microsoft.EntityFrameworkCore;

namespace GoogleSlides.Api.Services
{
    public class TemplateService
    {
        private readonly ApplicationDbContext _dbContext;

        public TemplateService(ApplicationDbContext dbContext)
        {
            _dbContext = dbContext;
        }

        public void SaveTemplate(TemplateMetadata template)
        {
            var existingTemplate = _dbContext.Templates.Include(t => t.Slides)
                                                       .ThenInclude(s => s.Placeholders)
                                                       .FirstOrDefault(t => t.Id == template.Id);

            if (existingTemplate != null)
            {
                // Update existing template and related entities
                _dbContext.Entry(existingTemplate).CurrentValues.SetValues(template);

                // Remove deleted slides
                foreach (var existingSlide in existingTemplate.Slides.ToList())
                {
                    if (!template.Slides.Any(s => s.Id == existingSlide.Id))
                    {
                        _dbContext.Remove(existingSlide);
                    }
                }

                // Update or add new slides and placeholders
                foreach (var slide in template.Slides)
                {
                    var existingSlide = existingTemplate.Slides.FirstOrDefault(s => s.Id == slide.Id);
                    if (existingSlide != null)
                    {
                        _dbContext.Entry(existingSlide).CurrentValues.SetValues(slide);
                        existingSlide.Placeholders = slide.Placeholders;
                    }
                    else
                    {
                        existingTemplate.Slides.Add(slide);
                    }
                }
            }
            else
            {
                _dbContext.Templates.Add(template);
            }

            _dbContext.SaveChanges();
        }

        public TemplateMetadata? GetTemplateById(string id)
        {
            return _dbContext.Templates.Include(t => t.Slides.OrderBy(s => s.Index))
                                      .ThenInclude(s => s.Placeholders)
                                      .FirstOrDefault(t => t.Id == id);
        }

    }
}
