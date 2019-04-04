using Microsoft.EntityFrameworkCore;

namespace chnu.Models
{
    public class UniversityContext : DbContext
    {
        public DbSet<Year> Years { get; set; }

        public DbSet<Student> Students { get; set; }

        public DbSet<Subject> Subjects { get; set; }

        public DbSet<Group> Groups { get; set; }

        public UniversityContext(DbContextOptions<UniversityContext> options)
            : base(options)
        {
            Database.EnsureCreated();
        }

        public DbSet<User> Admins { get; set; }
    }
}
