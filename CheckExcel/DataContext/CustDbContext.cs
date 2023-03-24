using CheckExcel.Models;
using Microsoft.EntityFrameworkCore;

namespace CheckExcel.DataContext
{
    public class CustDbContext : DbContext
    {
        public CustDbContext(DbContextOptions<CustDbContext> options) : base(options)
        {

        }

        public virtual DbSet<Customer> Customers { get; set; }
    }
}
