using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace ASTA.Classes
{

#pragma warning disable IDE1006 // Naming Styles
    /// <summary>
    /// Model 'Person'
    /// </summary>
    [Table("OBJ_PERSON")]
    public class Person
    {
        public string all_param { get; set; }//] [nvarchar] (32) NULL,
        public string area_id { get; set; }//] [nvarchar] (40) NULL,
        public string auto_brand { get; set; }//] [nvarchar] (32) NULL,
        public string auto_number { get; set; }//] [nvarchar] (32) NULL,
        public int auto_pass_type { get; set; }//] [int] NULL,
        public string begin { get; set; }//] [nvarchar] (25) NULL,
        public string begin_temp_level { get; set; }//] [nvarchar] (25) NULL,
        public string card { get; set; }//] [nvarchar] (255) NULL,
        public string card_date { get; set; }//] [nvarchar] (25) NULL,
        public int card_loss { get; set; }//] [int] NULL,
        public int card_type { get; set; }//] [int] NULL,
        public string comment { get; set; }//] [ntext] NULL,
        public string department { get; set; }//] [nvarchar] (255) NULL,
        public string drivers_licence { get; set; }//] [nvarchar] (120) NULL,
        public string email { get; set; }//] [nvarchar] (60) NULL,
        public string end_temp_level { get; set; }//] [nvarchar] (25) NULL,
        public string expired { get; set; }//] [nvarchar] (25) NULL,
        public string external_id { get; set; }//] [nvarchar] (40) NULL,
        public string facility_code { get; set; }//] [nvarchar] (255) NULL,
        public DateTime? finished_at { get; set; }//] [datetime] NULL,
        public int flags { get; set; }//] [int] NULL,

        [Key]
        public Guid guid { get; set; }//] [uniqueidentifier] NULL,

        public string id { get; set; }//] [nvarchar] (40) NULL,
        public int is_active_temp_level { get; set; }//] [smallint] NULL,
        public int is_apb { get; set; }//] [smallint] NULL,
        public int is_locked { get; set; }//] [smallint] NULL,
        public string level2_id { get; set; }//] [nvarchar] (40) NULL,
        public string level_id { get; set; }//] [ntext] NULL,
        public string levels_times { get; set; }//] [ntext] NULL,
        public string location { get; set; }//] [nvarchar] (16) NULL,
        public string marketing_info { get; set; }//] [ntext] NULL,

        [StringLength(255)]
        public string name { get; set; }//] [nvarchar] (255) NULL,

        public string owner_person_id { get; set; }//] [nvarchar] (40) NULL,
        public string parent_id { get; set; }//] [nvarchar] (40) NULL,
        public string passport { get; set; }//] [nvarchar] (120) NULL,

        [StringLength(255)]
        public string patronymic { get; set; }//] [nvarchar] (255) NULL,

        public string person { get; set; }//] [nvarchar] (50) NULL,
        public string phone { get; set; }//] [nvarchar] (60) NULL,
        public string pin { get; set; }//] [nvarchar] (255) NULL,
        public string post { get; set; }//] [nvarchar] (255) NULL,
        public int pur { get; set; }//] [int] NULL,
        public string rubeg8_zone_id { get; set; }//] [nvarchar] (40) NULL,
        public string schedule_id { get; set; }//] [nvarchar] (40) NULL,
        public DateTime? started_at { get; set; }//] [datetime] NULL,

        [StringLength(255)]
        public string surname { get; set; }//] [nvarchar] (255) NULL,

        public string tabnum { get; set; }//] [nvarchar] (60) NULL,
        public string teleph_work { get; set; }//] [nvarchar] (16) NULL,
        public string temp_card { get; set; }//] [nvarchar] (15) NULL,
        public string temp_level_id { get; set; }//] [ntext] NULL,
        public string temp_levels_times { get; set; }//] [ntext] NULL,
        public string visit_birthplace { get; set; }//] [nvarchar] (120) NULL,
        public string visit_document { get; set; }//] [nvarchar] (120) NULL,
        public string visit_purpose { get; set; }//] [nvarchar] (120) NULL,
        public string visit_reg { get; set; }//] [nvarchar] (120) NULL,
        public DateTime? when_area_id_changed { get; set; }//] [datetime] NULL,
        public string whence { get; set; }//] [nvarchar] (60) NULL,
        public string who_card { get; set; }//] [nvarchar] (60) NULL,
        public string who_level { get; set; }//] [nvarchar] (60) NULL
    }
#pragma warning restore IDE1006 // Naming Styles

    /// <summary>
    /// Map model 'Person' at table 'OBJ_PERSON' of DB 'Intellect'
    /// </summary>
    public class PersonMap : EntityTypeConfiguration<Person>
    {
        public PersonMap()
        {
            // Primary key
            this.HasKey(m => m.guid);
            this.Property(m => m.guid)
                .HasColumnType("uniqueidentifier")
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);
            // Properties
            this.Property(m => m.facility_code)
                 .HasColumnType("nvarchar");
            this.Property(m => m.tabnum)
                 .HasColumnType("nvarchar");
            this.Property(m => m.card)
                 .HasColumnType("nvarchar");
            this.Property(m => m.id)
                .HasColumnType("nvarchar");
            this.Property(m => m.name)
                .HasColumnType("nvarchar");
            this.Property(m => m.surname)
                .HasColumnType("nvarchar");
            this.Property(m => m.patronymic)
                .HasColumnType("nvarchar");
            this.Property(m => m.post)
                .HasColumnType("nvarchar");
            this.Property(m => m.parent_id)
                .HasColumnType("nvarchar");
            //Ignored properties
            this.Ignore(m => m.all_param);
            this.Ignore(m => m.area_id);
            this.Ignore(m => m.auto_brand);
            this.Ignore(m => m.auto_number);
            this.Ignore(m => m.auto_pass_type);
            this.Ignore(m => m.begin);
            this.Ignore(m => m.begin_temp_level);
            this.Ignore(m => m.card_date);
            this.Ignore(m => m.card_loss);
            this.Ignore(m => m.card_type);
            this.Ignore(m => m.comment);
            this.Ignore(m => m.department);
            this.Ignore(m => m.drivers_licence);
            this.Ignore(m => m.email);
            this.Ignore(m => m.end_temp_level);
            this.Ignore(m => m.expired);
            this.Ignore(m => m.external_id);
            this.Ignore(m => m.finished_at);
            this.Ignore(m => m.flags);
            this.Ignore(m => m.is_active_temp_level);
            this.Ignore(m => m.is_apb);
            this.Ignore(m => m.is_locked);
            this.Ignore(m => m.level2_id);
            this.Ignore(m => m.level_id);
            this.Ignore(m => m.levels_times);
            this.Ignore(m => m.location);
            this.Ignore(m => m.marketing_info);
            this.Ignore(m => m.owner_person_id);
            this.Ignore(m => m.passport);
            this.Ignore(m => m.person);
            this.Ignore(m => m.phone);
            this.Ignore(m => m.pin);
            this.Ignore(m => m.pur);
            this.Ignore(m => m.rubeg8_zone_id);
            this.Ignore(m => m.schedule_id);
            this.Ignore(m => m.started_at);
            this.Ignore(m => m.teleph_work);
            this.Ignore(m => m.temp_card);
            this.Ignore(m => m.temp_level_id);
            this.Ignore(m => m.temp_levels_times);
            this.Ignore(m => m.visit_birthplace);
            this.Ignore(m => m.visit_document);
            this.Ignore(m => m.visit_purpose);
            this.Ignore(m => m.visit_reg);
            this.Ignore(m => m.when_area_id_changed);
            this.Ignore(m => m.whence);
            this.Ignore(m => m.who_card);
            this.Ignore(m => m.who_level);
            // Table & column mappings
            this.ToTable("OBJ_PERSON", "dbo");
            this.Property(m => m.guid).HasColumnName("guid");
            this.Property(m => m.facility_code).HasColumnName("facility_code");
            this.Property(m => m.tabnum).HasColumnName("tabnum");
            this.Property(m => m.card).HasColumnName("card");
            this.Property(m => m.id).HasColumnName("id");
            this.Property(m => m.name).HasColumnName("name");
            this.Property(m => m.surname).HasColumnName("surname");
            this.Property(m => m.patronymic).HasColumnName("patronymic");
            this.Property(m => m.post).HasColumnName("post");
            this.Property(m => m.parent_id).HasColumnName("parent_id");
        }
    }
}