using BoschACDC.Models;
using Microsoft.EntityFrameworkCore;

namespace BoschACDC.Data
{
    public class SysFreightDbContext : DbContext
    {
        public SysFreightDbContext(DbContextOptions<SysFreightDbContext> options) : base(options){} 
        //public virtual DbSet<BoschModel> Boschs { get; set; }
        //protected override void OnModelCreating(ModelBuilder modelBuilder)
        //{
        //    modelBuilder.Entity<BoschModel>()
        //        .HasKey(c => new { c.DeclarationNum, c.LineNum });
        //}
    }
}
