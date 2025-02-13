using DataBaseToExcelAPI.Models;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;

namespace DataBaseToExcelAPI.Data
{
    public class ApplicationDbContext: DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options)
        {
        }
        public DbSet<Student> Students { get; set; }
    }
}
