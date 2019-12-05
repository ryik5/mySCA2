using ASTA.Classes.People;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASTA.Classes
{
    public class DBConnectSQL : DbContext
    {
        private string _connection;//"Server=(localdb)\\mssqllocaldb;Database=helloappdb;Trusted_Connection=True;"
        
        public DbSet<Protocol> ProtocolObjects { get; set; }

        public DBConnectSQL(string connection)
        {
            _connection = connection;
            Database.EnsureCreated();
        }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer(_connection);
        }
    }

    public class Protocol
    {
     public string objtype { get; set; }
        public string objid { get; set; }
        public string action { get; set; }
        public string region_id { get; set; }
        public string param0 { get; set; }
        public string param1 { get; set; }
        public string param2 { get; set; }
        public string param3 { get; set; }
        public string user_param_double { get; set; }
        public DateTime date { get; set; }
        public DateTime time { get; set; }
        public DateTime time2 { get; set; }
        public string owner { get; set; }
      //  public string pk { get; set; }
    }
}
