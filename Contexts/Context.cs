using EmailReseiver.Models;
using Microsoft.EntityFrameworkCore;
namespace EmailReseiver.Contexts
{
    public class Context: DbContext
    {
        public Context(DbContextOptions<Context> options) : base(options) { }
        public DbSet<Xmls> Xmls { get; set; }
        
    }
}