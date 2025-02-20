using System;
using System.Data.Entity;

public class AppDbContext : DbContext
{
    public DbSet<RentalRecord> Rentals { get; set; }

    public AppDbContext() : base("name=AppDbContext")
    { }

    protected override void OnModelCreating(DbModelBuilder modelBuilder)
    {
        modelBuilder.Entity<RentalRecord>().ToTable("Rentals");
        base.OnModelCreating(modelBuilder);
    }
}
