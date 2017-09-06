using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace UM.Models
{
    public class MyDbContext : DbContext
    {
        public MyDbContext() : base("name=db")
        {

        }
        public virtual DbSet<NhanVien> NhanViens { get; set; }
    }
}