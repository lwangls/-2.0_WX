using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 我的工具箱
{
    class MyDbContext : DbContext
    {
        //       public MyDbContext() : base("DefaultConnection") { }
        public MyDbContext(DbConnection conn) : base(conn,true) { }

        public DbSet<Lang> Langs { get; set; }



        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            //            modelBuilder.Entity<Person>()
            //                .HasMany(_ => _.OwnedCars).WithOptional(_ => _.Owner);
        }

        //使用自定义连接串
    }
}
