using GoogleSlides.Api.Models;
using GoogleSlides.Api.Models.Domain;
using Microsoft.EntityFrameworkCore;

namespace GoogleSlides.Api.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext() {}

        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options) {}

        public DbSet<AuthConfig> AuthConfig { get; set; }

        public DbSet<TemplateMetadata> Templates { get; set; }
        public DbSet<SlideMetadata> Slides { get; set; }
        public DbSet<PlaceholderMetadata> Placeholders { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            // Configure the relationships between entities
            modelBuilder.Entity<TemplateMetadata>()
                .HasMany(t => t.Slides)
                .WithOne(s => s.Template)
                .HasForeignKey(s => s.TemplateId)
                .OnDelete(DeleteBehavior.Cascade);

            modelBuilder.Entity<SlideMetadata>()
                .HasMany(s => s.Placeholders)
                .WithOne(p => p.Slide)
                .HasForeignKey(p => p.SlideId)
                .OnDelete(DeleteBehavior.Cascade);
        }


    }
}
