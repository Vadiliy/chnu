using Microsoft.EntityFrameworkCore;

namespace chnu.Models
{
    public class UniversityContext : DbContext
    {
        public DbSet<Year> Years { get; set; }

        public UniversityContext(DbContextOptions<UniversityContext> options)
            : base(options)
        {
            Database.EnsureCreated();
        }
    }
}
