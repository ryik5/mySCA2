using ASTA.Classes.People;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;
using System.Data.Entity.ModelConfiguration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASTA.Classes
{
    public class DBConnectSQL : DbContext
    {
        public DbSet<Protocol> ProtocolObjects { get; set; }

        public DBConnectSQL(string connection) : base(connection) //connection;//"Server=(localdb)\\mssqllocaldb;Database=helloappdb;Trusted_Connection=True;"
        {
            Database.SetInitializer<DBConnectSQL>(null);
        }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Configurations.Add(new LastVisitors());
        }
    }

    public class Protocol
    {
        public string objtype { get; set; }
        public string objid { get; set; }
        public string action { get; set; }
        public string region_id { get; set; }
        public string FIO { get; set; }
        public string idCard { get; set; }
        public string param2 { get; set; }
        public string param3 { get; set; }
        public float user_param_double { get; set; }
        public DateTime date { get; set; }
        public DateTime time { get; set; }
        public DateTime time2 { get; set; }
        public string owner { get; set; }

        //    [System.ComponentModel.DataAnnotations.Schema.DatabaseGenerated(System.ComponentModel.DataAnnotations.Schema.DatabaseGeneratedOption.Identity)]
        //    [System.ComponentModel.DataAnnotations.Key]
        public Guid pk { get; set; }
    }

    public class LastVisitors : EntityTypeConfiguration<Protocol>
    {
        public LastVisitors()
        {         
            // Primary key         
            this.HasKey(m => m.pk);
            this.Property(m => m.pk)
                .HasColumnType("uniqueidentifier")
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);
            // Properties         
            this.Property(m => m.FIO)
                 .HasColumnType("nvarchar");
            this.Property(m => m.idCard)
                 .HasColumnType("nvarchar");
            this.Property(m => m.action)
                 .HasColumnType("nvarchar");
            this.Property(m => m.objid)
                .HasColumnType("nvarchar");
            this.Property(m => m.objtype)
                .HasColumnType("nvarchar");
            this.Property(m => m.date)
                .HasColumnType("datetime");
            this.Property(m => m.time)
                .HasColumnType("datetime");

            this.Ignore(m => m.time2);
            this.Ignore(m => m.param2);
            this.Ignore(m => m.param3);
            this.Ignore(m => m.owner);
            this.Ignore(m => m.region_id);
            this.Ignore(m => m.user_param_double);

            // Table & column mappings       
            this.ToTable("Protocol", "dbo");
            this.Property(m => m.pk).HasColumnName("pk");
            this.Property(m => m.FIO).HasColumnName("param0");
            this.Property(m => m.idCard).HasColumnName("param1");
            this.Property(m => m.objid).HasColumnName("objid");
            this.Property(m => m.objtype).HasColumnName("objtype");
            this.Property(m => m.action).HasColumnName("action");
            this.Property(m => m.date).HasColumnName("date");
            this.Property(m => m.time).HasColumnName("time");
        }
    }
}
