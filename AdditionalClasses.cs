using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Security.Cryptography;
using System.Windows.Forms;

namespace ASTA
{

    /*

     using in a method:
     {
     DataTable dt = DbSqlJob.GetTable();

     var category_Data = dt.AsEnumerable()
         .GroupBy(row => row.Field<string>("Категория"))
         .Select(cat => new {
             Category_Name = cat.Key,
             Category_Count = cat.Count()
         });
     }

     public static class DbSqlJob
 {
     private static string DB_PATH = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "testbase.mdf");
     private static string CONNECT_STRING = string.Format(
                           @"Data Source=.\SQLEXPRESS;AttachDbFilename={0};" +
                           "Integrated Security=True;Connect Timeout=30;User Instance=True", DB_PATH);
     public static DataTable GetTable()
     {
         DataTable dt = new DataTable();
         using (System.Data.SqlClient.SqlDataAdapter adapter = new System.Data.SqlClient.SqlDataAdapter("SELECT * FROM Table_1", CONNECT_STRING))
         {
             adapter.Fill(dt);
         }
         return dt;
     }
 }*/


    //todo
    /*
     const string connStr = "server=localhost;user=root; database=test;password=;";
using (MySqlConnection conn = new MySqlConnection(connStr))
{
string sql = "SELECT Text FROM Kody WHERE Kod=@Kod";
MySqlCommand comand = new MySqlCommand(sql, conn);
command.Parameters.AddWithValue("@Kod", textBox1.Text);
conn.Open();
string name = comand.ExecuteScalar().ToString();
label1.Text = name;
}
*/

    interface IDBParameter
    {
        string DBName { get; set; }
        string TableName { get; set; }
        string DBConnectionString { get; set; }
        string DBQuery { get; set; }
        void Execute();
    }

    class ParameterConfig
    {
        public string parameterName;
        public string parameterDescription;
        public string parameterValue;
        public string dateCreated;
        public bool isPassword;
        public string isExample;

        public override string ToString()
        {
            return parameterName + "\t" + parameterDescription + "\t" + parameterValue + "\t" + dateCreated + "\t" +
                isPassword.ToString() + "\t" + isExample;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            ParameterConfig df = obj as ParameterConfig;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }
        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    //todo
    /*
     const string connStr = "server=localhost;user=root; database=test;password=;";
using (MySqlConnection conn = new MySqlConnection(connStr))
{
string sql = "SELECT Text FROM Kody WHERE Kod=@Kod";
MySqlCommand comand = new MySqlCommand(sql, conn);
command.Parameters.AddWithValue("@Kod", textBox1.Text);
conn.Open();
string name = comand.ExecuteScalar().ToString();
label1.Text = name;
}
*/

    class ParameterOfConfiguration
    {
        public string ParameterName { get; set; }
        public string ParameterValue { get; set; }
        public string ParameterDescription { get; set; }
        public string isExample { get; set; }
        public bool isPassword { get; set; }

        public static ParameterOfConfigurationBuilder CreateParameter()
        {
            return new ParameterOfConfigurationBuilder();
        }        
    }

    class ParameterOfConfigurationBuilder
    {
        private ParameterOfConfiguration parameterOfConfiguration;

        public ParameterOfConfigurationBuilder()
        {
            parameterOfConfiguration = new ParameterOfConfiguration();
        }
        public ParameterOfConfigurationBuilder SetParameterName(string name)
        {
            parameterOfConfiguration.ParameterName = name;
            return this;
        }
        public ParameterOfConfigurationBuilder SetParameterValue(string name)
        {
            parameterOfConfiguration.ParameterValue = name;
            return this;
        }
        public ParameterOfConfigurationBuilder SetParameterDescription(string name)
        {
            parameterOfConfiguration.ParameterDescription = name;
            return this;
        }
        public ParameterOfConfigurationBuilder SetIsExample(string name)
        {
            parameterOfConfiguration.isExample = name;
            return this;
        }
        public ParameterOfConfigurationBuilder IsPassword(bool state)
        {
                parameterOfConfiguration.isPassword = state;
                return this;
        }

        public static implicit operator ParameterOfConfiguration(ParameterOfConfigurationBuilder parameter)
        {
            return parameter.parameterOfConfiguration;
        }
    }


    class ParameterOfConfigurationInSQLiteDB 
    {
        ParameterOfConfiguration parameterOfConfiguration;
         System.IO.FileInfo databasePerson;

        static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        readonly byte[] btsMess1 = Convert.FromBase64String(@"OCvesunvXXsxtt381jr7vp3+UCwDbE4ebdiL1uinGi0="); //Key Encrypt
        readonly byte[] btsMess2 = Convert.FromBase64String(@"NO6GC6Zjl934Eh8MAJWuKQ=="); //Key Decrypt

        public ParameterOfConfigurationInSQLiteDB(System.IO.FileInfo _databasePerson)
        {
            databasePerson = _databasePerson;
        }

        public string SaveParameter(ParameterOfConfiguration _parameterOfConfiguration)
        {
            parameterOfConfiguration = _parameterOfConfiguration;

            if (parameterOfConfiguration.ParameterName == null || parameterOfConfiguration.ParameterName.Length == 0)
                throw new ArgumentNullException("a ParameterName should have a name");

            if (parameterOfConfiguration.isPassword && (parameterOfConfiguration.ParameterValue == null || parameterOfConfiguration.ParameterValue.Length == 0))
                throw new ArgumentNullException("this ParameterValue can't be empty!");
            else
            {
                if (databasePerson.Exists)
                {
                    using (SQLiteConnection sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                    {
                        sqlConnection.Open();

                        SQLiteCommand sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                        sqlCommand1.ExecuteNonQuery();
                        string ParameterValueSave = "";
                        if (parameterOfConfiguration.isPassword)
                        { ParameterValueSave = EncryptionDecryptionCriticalData.EncryptStringToBase64Text(parameterOfConfiguration.ParameterValue, btsMess1, btsMess2); }
                        else
                        { ParameterValueSave = parameterOfConfiguration.ParameterValue; }

                        using (SQLiteCommand sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'ConfigDB' (ParameterName, Value, Description, DateCreated, IsPassword, IsExample)" +
                                   " VALUES (@ParameterName, @Value, @Description, @DateCreated, @IsPassword, @IsExample)", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@ParameterName", DbType.String).Value = parameterOfConfiguration.ParameterName;
                            sqlCommand.Parameters.Add("@Value", DbType.String).Value = ParameterValueSave;
                            sqlCommand.Parameters.Add("@Description", DbType.String).Value = parameterOfConfiguration.ParameterDescription;
                            sqlCommand.Parameters.Add("@DateCreated", DbType.String).Value = DateTime.Now.ToYYYYMMDDHHMM();
                            sqlCommand.Parameters.Add("@IsPassword", DbType.Boolean).Value = parameterOfConfiguration.isPassword;
                            sqlCommand.Parameters.Add("@IsExample", DbType.String).Value = parameterOfConfiguration.isExample;
                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }
                        sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                        sqlCommand1.ExecuteNonQuery();
                    }
                    return parameterOfConfiguration.ParameterName + " (" + parameterOfConfiguration.ParameterValue + ")" + " - was saved!";
                }
                else
                {
                    return parameterOfConfiguration.ParameterName + " - wasn't saved!\n"+ databasePerson.ToString()+" is not exist!";
                }
            }
        }

        public  List<ParameterConfig> GetParameters(string parameter)
        {
            List<ParameterConfig> parametersConfig = new List<ParameterConfig>(); ;

            ParameterConfig parameterConfig = new ParameterConfig() { parameterName = "", parameterDescription = "", parameterValue = "", isPassword = false, isExample = "no" };

            string value = ""; string valueTmp = ""; string name = ""; string decrypt = "";
            if (databasePerson.Exists)
            {
                using (SQLiteConnection sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();

                    try
                    {
                        using (SQLiteCommand sqlCommand = sqlConnection.CreateCommand())
                        {
                            if (parameter.Length > 0 && parameter != "%%")
                            {
                                sqlCommand.CommandText = @"Select ParameterName, Value, Description, DateCreated, IsPassword, IsExample from ConfigDB where ParameterName=@parameter";
                                sqlCommand.Parameters.Add(new SQLiteParameter("@parameter") { Value = parameter });
                            }
                            else
                            {
                                sqlCommand.CommandText = @"Select ParameterName, Value, Description, DateCreated, IsPassword, IsExample from ConfigDB";
                            }
                            sqlCommand.CommandType = System.Data.CommandType.Text;
                            logger.Trace("ParameterOfConfigurationInSQLiteDB: " + sqlCommand.CommandText.ToString());
                            using (SQLiteDataReader reader = sqlCommand.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    name = reader["ParameterName"]?.ToString();
                                    if (name?.Length > 0)
                                    {
                                        parameterConfig = new ParameterConfig() { parameterName = "", parameterDescription = "", parameterValue = "", isPassword = false, isExample = "no" };
                                        decrypt = reader["IsPassword"]?.ToString();

                                        valueTmp = reader["Value"]?.ToString();
                                        if (decrypt == "1" && valueTmp?.Length > 0)
                                        {
                                            value = EncryptionDecryptionCriticalData.DecryptBase64ToString(valueTmp, btsMess1, btsMess2);
                                            parameterConfig.isPassword = true;
                                        }
                                        else
                                        {
                                            value = valueTmp;
                                            parameterConfig.isPassword = false;
                                        }
                                        parameterConfig.parameterName = name;
                                        parameterConfig.parameterValue = value;
                                        parameterConfig.parameterDescription = reader["Description"]?.ToString();
                                        parameterConfig.dateCreated = reader["DateCreated"]?.ToString();
                                        parameterConfig.isExample = reader["IsExample"]?.ToString();

                                        parametersConfig.Add(parameterConfig);
                                        logger.Trace("ParameterOfConfigurationInSQLiteDB, add: " + name);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception exc)
                    {
                        logger.Trace("ParameterOfConfigurationInSQLiteDB, error: " + exc.ToString());
                    }
                }
            }
            return parametersConfig;
        }
    }

    static class EncryptionDecryptionCriticalData
    {
        public static string EncryptStringToBase64Text(string plainText, byte[] Key, byte[] IV) //Encrypt variables PlainText Data
        {
            string sBase64Test;
            sBase64Test = Convert.ToBase64String(EncryptStringToBytes(plainText, Key, IV));
            return sBase64Test;
        }

        public static byte[] EncryptStringToBytes(string plainText, byte[] Key, byte[] IV) //Encrypt variables PlainText Data
        {
            // Check arguments. 
            if (plainText == null || plainText.Length <= 0)
                throw new ArgumentNullException("plainText");
            if (Key == null || Key.Length <= 0)
                throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0)
                throw new ArgumentNullException("IV");
            byte[] encrypted;
            // Create an RijndaelManaged object with the specified key and IV. 
            using (RijndaelManaged rijAlg = new RijndaelManaged())
            {
                rijAlg.Key = Key;
                rijAlg.IV = IV;

                // Create a decrytor to perform the stream transform.
                ICryptoTransform encryptor = rijAlg.CreateEncryptor(rijAlg.Key, rijAlg.IV);

                // Create the streams used for encryption. 
                using (System.IO.MemoryStream msEncrypt = new System.IO.MemoryStream())
                {
                    using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                    {
                        using (System.IO.StreamWriter swEncrypt = new System.IO.StreamWriter(csEncrypt))
                        {

                            //Write all data to the stream.
                            swEncrypt.Write(plainText);
                        }
                        encrypted = msEncrypt.ToArray();
                    }
                }
            }
            // Return the encrypted bytes from the memory stream. 
            return encrypted;
        }

        public static string DecryptBase64ToString(string sBase64Text, byte[] Key, byte[] IV) //Encrypt variables PlainText Data
        {
            byte[] bBase64Test;
            bBase64Test = Convert.FromBase64String(sBase64Text);
            return DecryptStringFromBytes(bBase64Test, Key, IV);
        }

        public static string DecryptStringFromBytes(byte[] cipherText, byte[] Key, byte[] IV) //Decrypt PlainText Data to variables
        {
            // Check arguments. 
            if (cipherText == null || cipherText.Length <= 0)
                throw new ArgumentNullException("cipherText");
            if (Key == null || Key.Length <= 0)
                throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0)
                throw new ArgumentNullException("IV");

            // Declare the string used to hold the decrypted text.
            string plaintext = null;

            // Create an RijndaelManaged object  with the specified key and IV.
            using (RijndaelManaged rijAlg = new RijndaelManaged())
            {
                rijAlg.Key = Key;
                rijAlg.IV = IV;

                // Create a decrytor to perform the stream transform.
                ICryptoTransform decryptor = rijAlg.CreateDecryptor(rijAlg.Key, rijAlg.IV);

                // Create the streams used for decryption. 
                using (System.IO.MemoryStream msDecrypt = new System.IO.MemoryStream(cipherText))
                {
                    using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                    {
                        using (System.IO.StreamReader srDecrypt = new System.IO.StreamReader(csDecrypt))
                        {
                            // Read the decrypted bytes from the decrypting stream and place them in a string. 
                            plaintext = srDecrypt.ReadToEnd();
                        }
                    }
                }
            }

            return plaintext;
        }
    }

    class MailingStructure
    {
        public string _sender;
        public string _recipient;
        public string _groupsReport;
        public string _nameReport;
        public string _descriptionReport;
        public string _period;
        public string _status;
        public string _typeReport;
        public string _dayReport;

        public override string ToString()
        {
            return _sender + "\t" + _recipient + "\t" +
                _groupsReport + "\t" + _nameReport + "\t" + _descriptionReport + "\t" + _period + "\t" +
                _status + "\t" + _typeReport + "\t" + _dayReport;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            MailingStructure df = obj as MailingStructure;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    interface IPerson
    {
        string FIO { get; set; }
        string NAV { get; set; }
    }

    class Person : IPerson
    {
        public string FIO { get; set; }
        public string NAV { get; set; }

        public override string ToString()
        {
            return FIO + "\t" + NAV;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            Person df = obj as Person;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    class PersonFull : IPerson
    {
        public int idCard;//= 0
        public string FIO { get; set; }//= ""
        public string NAV { get; set; }//= ""
        public string Department;//= ""
        public string DepartmentId;//= ""
        public string DepartmentBossCode;//= ""
        public string PositionInDepartment;//= ""
        public string GroupPerson;//= ""
        public string City;//= ""
        public int ControlInSeconds;//= 32460
        public int ControlOutSeconds;// =64800
        public string ControlInHHMM;//= "09:00"
        public string ControlOutHHMM;//= "18:00"
        public string Shift;//= ""
        public string Comment;//= ""

        public override string ToString()
        {
            return idCard + "\t" + FIO + "\t" + NAV + "\t" + this.Department + "\t" + DepartmentId + "\t" + DepartmentBossCode + "\t" +
                PositionInDepartment + "\t" + GroupPerson + "\t" + City + "\t" +
                ControlInSeconds + "\t" + ControlOutSeconds + "\t" + ControlInHHMM + "\t" + ControlOutHHMM + "\t" +
                Shift + "\t" + Comment;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            PersonFull df = obj as PersonFull;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    class PassByPoint
    {
        public string _id;
        public string _name;
        public string _direction;
        public string _server;

        public override string ToString()
        {
            return _id + "\t" + _name + "\t" + _direction + "\t" + _server;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            PassByPoint df = obj as PassByPoint;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    class OutReasons
    {
        public string _id;
        public string _name;
        public string _visibleName;
        public int _hourly;

        public override string ToString()
        {
            return _id + "\t" + _name + "\t" + _visibleName + "\t" + _hourly;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            OutReasons df = obj as OutReasons;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    class OutPerson
    {
        public string _reason_id;//= "0"
        public string _nav;//= ""
        public string _date;//= ""
        public int _from;//= 0
        public int _to;//= 0
        public int _hourly;//= 0

        public override string ToString()
        {
            return _reason_id + "\t" + _nav + "\t" + _date + "\t" + _from + "\t" + _to + "\t" + _hourly;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            OutPerson df = obj as OutPerson;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    struct PeopleShift
    {
        public string _nav;
        public string _dayStartShift;
        public int _MoStart;
        public int _MoEnd;
        public int _TuStart;
        public int _TuEnd;
        public int _WeStart;
        public int _WeEnd;
        public int _ThStart;
        public int _ThEnd;
        public int _FrStart;
        public int _FrEnd;
        public int _SaStart;
        public int _SaEnd;
        public int _SuStart;
        public int _SuEnd;
        public string _Status;
        public string _Comment;
    }

    class AmountMembersOfGroup
    {
        public int _amountMembers;
        public string _groupName;
        public string _emails;

        public override string ToString()
        {
            return _amountMembers + "\t" + _groupName + "\t" + _emails;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            AmountMembersOfGroup df = obj as AmountMembersOfGroup;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }
        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    interface IDepartment
    {
        string _departmentId { get; set; }
        string _departmentDescription { get; set; }
        string _departmentBossCode { get; set; }
    }

    class Department : IDepartment
    {
        public string _departmentId { get; set; }
        public string _departmentDescription { get; set; }
        public string _departmentBossCode { get; set; }

        public override string ToString()
        {
            return _departmentId + "\t" + _departmentDescription + "\t" + _departmentBossCode;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            Department df = obj as Department;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }
        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    class DepartmentFull : IDepartment
    {
        public string _departmentId { get; set; }
        public string _departmentDescription { get; set; }
        public string _departmentBossCode { get; set; }
        public string _departmentBossEmail { get; set; }

        public override string ToString()
        {
            return _departmentId + "\t" + _departmentDescription + "\t" + _departmentBossCode + "\t" + _departmentBossEmail;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            DepartmentFull df = obj as DepartmentFull;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }
        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    class DataGridViewSeekValuesInSelectedRow
    {
        public string[] values = new string[10];
        private bool correctData { get; set; }

        public void FindValuesInCurrentRow(DataGridView DGV, params string[] columnsName)
        {
            int IndexCurrentRow = DGV.CurrentRow.Index;
            correctData = (-1 < IndexCurrentRow & 0 < columnsName.Length & columnsName.Length < 11) ? true : false;

            if (correctData)
            {
                for (int i = 0; i < DGV.ColumnCount; i++)
                {
                    if (DGV.Columns[i].HeaderText == columnsName[0])
                    {
                        values[0] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 1 && DGV.Columns[i].HeaderText == columnsName[1])
                    {
                        values[1] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 2 && DGV.Columns[i].HeaderText == columnsName[2])
                    {
                        values[2] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 3 && DGV.Columns[i].HeaderText == columnsName[3])
                    {
                        values[3] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 4 && DGV.Columns[i].HeaderText == columnsName[4])
                    {
                        values[4] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 5 && DGV.Columns[i].HeaderText == columnsName[5])
                    {
                        values[5] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 6 && DGV.Columns[i].HeaderText == columnsName[6])
                    {
                        values[6] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 7 && DGV.Columns[i].HeaderText == columnsName[7])
                    {
                        values[7] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 8 && DGV.Columns[i].HeaderText == columnsName[8])
                    {
                        values[8] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length == 10 && DGV.Columns[i].HeaderText == columnsName[9])
                    {
                        values[9] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                }
            }
            else
            {
                throw new IndexOutOfRangeException("Taken collumns more then 10!");
            }
        }
    }

    static class DataTableExtensions
    {
        public static void SetColumnsOrder(this DataTable table, params string[] columnNames)
        {
            List<string> listColNames = columnNames.ToList();

            //Remove invalid column names.
            foreach (string colName in columnNames)
            {
                if (!table.Columns.Contains(colName))
                {
                    listColNames.Remove(colName);
                }
            }

            int columnIndex = 0;
            foreach (var columnName in listColNames)
            {
                table.Columns[columnName].SetOrdinal(columnIndex);
                columnIndex++;
            }
        }
    }

    static class DateTimeExtensions
    {
        public static DateTime LastDayOfMonth(this DateTime date)
        {
            return new DateTime(date.Year, date.Month, 1).AddMonths(1).AddDays(-1);
        }

        public static DateTime FirstDayOfMonth(this DateTime date)
        {
            return new DateTime(date.Year, date.Month, 1);
        }

        public static string ToMonthName(this DateTime dateTime)
        {
            return System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(dateTime.Month);
        }

        public static string ToShortMonthName(this DateTime dateTime)
        {
            return System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(dateTime.Month);
        }

        //
        public static string ToYYYYMMDD(this DateTime dateTime)
        {
            return dateTime.ToString("yyyy-MM-dd");
        }

        public static string ToYYYYMMDDHHMM(this DateTime dateTime)
        {
            return DateTime.Now.ToString("yyyy-MM-dd HH:mm");
        }

    }

    class EncryptDecrypt
    {
        /*
            string plainText = "Hello, World!";            
            string TestKey = "gFlMfLZu4unBGfBh9weIuLwlcXHSa59vxyUQWM3yh1M=";
            string TestIV = "rRNgGypocle9VG9bth6kxg==";

            var ed = new EncDec(TestKey, TestIV);
            var cypherText = ed.Encrypt(plainText);  //4U8SPwju5aH9pHwmiYd1jQ==
            var plainText2 = ed.Decrypt(cypherText); //4U8SPwju5aH9pHwmiYd1jQ==
*/
        private byte[] key;
        private byte[] IV;

        public void EncDec(string keyText, string ivText)
        {
            key = Convert.FromBase64String(keyText);
            IV = Convert.FromBase64String(ivText);
        }

        public void EncDecKeys(byte[] key, byte[] IV)
        {
            this.key = key;
            this.IV = IV;
        }

        public string Encrypt(string plainText)
        {
            var bytes = EncryptStringToBytes(plainText, key, IV);
            return Convert.ToBase64String(bytes);
        }

        private static byte[] EncryptStringToBytes(string plainText, byte[] Key, byte[] IV)
        {
            // Check arguments. 
            if (plainText == null || plainText.Length <= 0)
                throw new ArgumentNullException("plainText");
            if (Key == null || Key.Length <= 0)
                throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0)
                throw new ArgumentNullException("IV");
            byte[] encrypted;
            // Create an RijndaelManaged object 
            // with the specified key and IV. 
            using (RijndaelManaged rijAlg = new RijndaelManaged())
            {
                rijAlg.Key = Key;
                rijAlg.IV = IV;

                // Create a decrytor to perform the stream transform.
                ICryptoTransform encryptor = rijAlg.CreateEncryptor(rijAlg.Key, rijAlg.IV);

                // Create the streams used for encryption. 
                using (System.IO.MemoryStream msEncrypt = new System.IO.MemoryStream())
                {
                    using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                    {
                        using (System.IO.StreamWriter swEncrypt = new System.IO.StreamWriter(csEncrypt))
                        {
                            swEncrypt.Write(plainText);  //Write all data to the stream.
                        }
                        encrypted = msEncrypt.ToArray();
                    }
                }
            }

            // Return the encrypted bytes from the memory stream. 
            return encrypted;
        }

        public string Decrypt(string cipherText)
        {
            var bytes = Convert.FromBase64String(cipherText);
            return DecryptStringFromBytes(bytes, key, IV);
        }

        private static string DecryptStringFromBytes(byte[] cipherText, byte[] Key, byte[] IV)
        {
            // Check arguments. 
            if (cipherText == null || cipherText.Length <= 0)
                throw new ArgumentNullException("cipherText");
            if (Key == null || Key.Length <= 0)
                throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0)
                throw new ArgumentNullException("IV");

            // Declare the string used to hold  the decrypted text. 
            string plaintext = null;

            // Create an RijndaelManaged object with the specified key and IV. 
            using (RijndaelManaged rijAlg = new RijndaelManaged())
            {
                rijAlg.Key = Key;
                rijAlg.IV = IV;

                // Create a decrytor to perform the stream transform.
                ICryptoTransform decryptor = rijAlg.CreateDecryptor(rijAlg.Key, rijAlg.IV);

                // Create the streams used for decryption. 
                using (System.IO.MemoryStream msDecrypt = new System.IO.MemoryStream(cipherText))
                {
                    using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                    {
                        using (System.IO.StreamReader srDecrypt = new System.IO.StreamReader(csDecrypt))
                        {

                            // Read the decrypted bytes from the decrypting stream 
                            // and place them in a string.
                            plaintext = srDecrypt.ReadToEnd();
                        }
                    }
                }
            }

            return plaintext;
        }
    }

}
