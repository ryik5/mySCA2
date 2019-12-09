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
    /// <summary>
    /// DB Connector to Table 'Protocol' of DB 'Intellect'
    /// </summary>
    public class ProtocolConnector : DbContext
    {
        public DbSet<Protocol> ProtocolObjects { get; set; }

        public ProtocolConnector(string connection) : base(connection) //connection;//"Server=(localdb)\\mssqllocaldb;Database=helloappdb;Trusted_Connection=True;"
        {
            Database.SetInitializer<ProtocolConnector>(null);
        }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Configurations.Add(new ProtocolMap());
        }
    }

   /// <summary>
   /// Model 'Protocol'
   /// </summary>
    public class Protocol
    {
        public string ActionType { get; set; }
        public string PointName { get; set; }
        public string ActionDescr { get; set; }
        public string region_id { get; set; }
        public string FIO { get; set; }
        public string IdCard { get; set; }
        public string param2 { get; set; }
        public string param3 { get; set; }
        public float user_param_double { get; set; }
        public DateTime ActionDate { get; set; }
        public DateTime ActionTime { get; set; }
        public DateTime ActionTime2 { get; set; }
        public string ActionRegistrator { get; set; }
        public Guid UniqueKey { get; set; }
    }

    /// <summary>
    /// Map model 'Protocol' at table 'Protocol' of DB 'Intellect'
    /// </summary>
    public class ProtocolMap : EntityTypeConfiguration<Protocol>
    {
        public ProtocolMap()
        {         
            // Primary key         
            this.HasKey(m => m.UniqueKey);
            this.Property(m => m.UniqueKey)
                .HasColumnType("uniqueidentifier")
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);
            // Properties         
            this.Property(m => m.FIO)
                 .HasColumnType("nvarchar");
            this.Property(m => m.IdCard)
                 .HasColumnType("nvarchar");
            this.Property(m => m.ActionDescr)
                 .HasColumnType("nvarchar");
            this.Property(m => m.PointName)
                .HasColumnType("nvarchar");
            this.Property(m => m.ActionType)
                .HasColumnType("nvarchar");
            this.Property(m => m.ActionDate)
                .HasColumnType("datetime");
            this.Property(m => m.ActionTime)
                .HasColumnType("datetime");
            //Ignored properties
            this.Ignore(m => m.ActionTime2);
            this.Ignore(m => m.param2);
            this.Ignore(m => m.param3);
            this.Ignore(m => m.ActionRegistrator);
            this.Ignore(m => m.region_id);
            this.Ignore(m => m.user_param_double);
            // Table & column mappings       
            this.ToTable("Protocol", "dbo");
            this.Property(m => m.UniqueKey).HasColumnName("pk");
            this.Property(m => m.FIO).HasColumnName("param0");
            this.Property(m => m.IdCard).HasColumnName("param1");
            this.Property(m => m.PointName).HasColumnName("objid");
            this.Property(m => m.ActionType).HasColumnName("objtype");
            this.Property(m => m.ActionDescr).HasColumnName("action");
            this.Property(m => m.ActionDate).HasColumnName("date");
            this.Property(m => m.ActionTime).HasColumnName("time");
        }
    }
}
