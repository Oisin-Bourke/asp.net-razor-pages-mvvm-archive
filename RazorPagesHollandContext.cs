//the DB context for Holland 
using Microsoft.EntityFrameworkCore;

namespace RazorPagesHolland.Models

{

    public class RazorPagesHollandContext : DbContext
    {
        public RazorPagesHollandContext(DbContextOptions<RazorPagesHollandContext> options)
            : base(options)
        {
        }

        //creates the 'enitity set' or table
        public DbSet<RazorPagesHolland.Models.Holland> Holland { get; set; } 
        public DbSet<RazorPagesHolland.Models.Survey> Survey { get; set; }
    }

}
