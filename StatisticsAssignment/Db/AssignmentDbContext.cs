using Microsoft.EntityFrameworkCore;
using StatisticsAssignment.Db.Entities;

namespace StatisticsAssignment.Db
{
    public sealed class AssignmentDbContext : DbContext
    {
        public AssignmentDbContext(DbContextOptions options)
            :base(options)
        {
        }

        public DbSet<CountryData> CountryData { get; set; }
    }
}
