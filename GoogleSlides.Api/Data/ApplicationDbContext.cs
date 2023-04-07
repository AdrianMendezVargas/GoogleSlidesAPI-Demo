using GoogleSlides.Api.Models.Domain;
using Microsoft.EntityFrameworkCore;

namespace GoogleSlides.Api.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext() {}

        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options) {}

        public DbSet<AuthConfig> AuthConfig { get; set; }

    }
}
