using GoogleSlides.Api.Data;
using GoogleSlides.Api.Models.Domain;
using Microsoft.EntityFrameworkCore;

namespace GoogleSlides.Api.Services
{
    public class AuthConfigService
    {
        private readonly ApplicationDbContext _context;

        public AuthConfigService(ApplicationDbContext context)
        {
            _context = context;
        }

        public async Task SaveAsync(AuthConfig authConfig)
        {
            if (authConfig.Id > 0)
            {
                _context.Update(authConfig);
            }
            else
            {
                await _context.AddAsync(authConfig);
            }
            await _context.SaveChangesAsync();
        }
        public async Task<bool> ExistAsync(int id)
        {
            return await _context.AuthConfig.AnyAsync(a => a.Id == id);
        }

        public void Save(AuthConfig authConfig)
        {
            if (authConfig.Id > 0)
            {
                _context.Update(authConfig);
            }
            else
            {
                _context.Add(authConfig);
            }
            _context.SaveChanges();
        }
    }
}
