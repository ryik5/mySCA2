﻿using System;
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
    abstract class ParameterOfConfiguration
    {
        public string ParameterName { get; set; }
        public string ParameterValue { get; set; }
        public string ParameterDescription { get; set; }
        public string isExample { get; set; }
        public bool isPassword { get; set; }
        public bool isRequired { get; set; }
        public abstract string SaveParameter();
        public abstract string GetParameter();
    }

    interface DBParameter
    {
        string DBName { get; set; }
        string TableName { get; set; }
        string DBConnectionString { get; set; }
        string DBQuery { get; set; }
        void Execute();
    }

    struct ParameterConfig
    {
        public string parameterName;
        public string parameterDescription;
        public string parameterValue;
        public bool isPassword;
        public string isExample;
    }

    class ParameterOfConfigurationInSQLiteDB : ParameterOfConfiguration
    {
        readonly byte[] btsMess1 = Convert.FromBase64String(@"OCvesunvXXsxtt381jr7vp3+UCwDbE4ebdiL1uinGi0="); //Key Encrypt
        readonly byte[] btsMess2 = Convert.FromBase64String(@"NO6GC6Zjl934Eh8MAJWuKQ=="); //Key Decrypt
        public System.IO.FileInfo databasePerson { get; set; }
        public override string SaveParameter()
        {
            if (isPassword && (ParameterValue == null || ParameterValue.Trim().Length == 0))
                throw new ArgumentNullException("ParameterValue as Paswsword is empty");
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
                        if (isPassword)
                            ParameterValueSave = EncryptionDecryptionCriticalData.EncryptStringToBase64Text(ParameterValue, btsMess1, btsMess2);
                        else ParameterValueSave = ParameterValue;
                        //("ConfigDB", "ParameterName , Value , Description , DateCreated , IsPassword , IsExample ");

                        using (SQLiteCommand sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'ConfigDB' (ParameterName, Value, Description, DateCreated, IsPassword, IsExample)" +
                                   " VALUES (@ParameterName, @Value, @Description, @DateCreated, @IsPassword, @IsExample)", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@ParameterName", DbType.String).Value = ParameterName;
                            sqlCommand.Parameters.Add("@Value", DbType.String).Value = ParameterValueSave;
                            sqlCommand.Parameters.Add("@Description", DbType.String).Value = ParameterDescription;
                            sqlCommand.Parameters.Add("@DateCreated", DbType.String).Value = DateTime.Now.ToYYYYMMDDHHMM();
                            sqlCommand.Parameters.Add("@IsPassword", DbType.Boolean).Value = isPassword;
                            sqlCommand.Parameters.Add("@IsExample", DbType.String).Value = isExample;
                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }
                        sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                        sqlCommand1.ExecuteNonQuery();
                    }
                }
                return "Ok";
            }
        }
        public override string GetParameter() { return ""; }
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

    struct MailingStructure
    {
        public string _sender ;
        public string _recipient ;
        public string _groupsReport ;
        public string _nameReport ;
        public string _descriptionReport ;
        public string _period ;
        public string _status ;
        public string _typeReport ;
        public string _dayReport ;
    }

    struct Person
    {
        public string FIO ;
        public string NAV ;
    }

    struct PersonFull
    {
        public int idCard;//= 0
        public string FIO;//= ""
        public string NAV;//= ""
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

        //   public string RealInHHMM;//= "09:00"
        //   public string RealOutHHMM;//= "18:00"

        //   public string RealWorkedDayHoursHHMM;//= "09:00"
        //   public string RealDate;//= ""
        //   public string RealDayOfWeek;//= ""

        //   public string serverSKD;//= ""
        //   public string namePassPoint;//= ""
        //   public string directionPass;//= ""
    }

    struct PassByPoint
    {
        public string _id ;
        public string _name ;
        public string _direction ;
        public string _server ;
    }

    struct OutReasons
    {
        public string _id ;
        public string _name ;
        public string _visibleName ;
        public int _hourly ;
    }

    struct OutPerson
    {
        public string _reason_id ;//= "0"
        public string _nav ;//= ""
        public string _date ;//= ""
        public int _from ;//= 0
        public int _to ;//= 0
        public int _hourly ;//= 0
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

    struct AmountMembersOfGroup
    {
        public int _amountMembers;
        public string _groupName;
        public string _emails;
    }

    struct Department
    {
        public string _departmentId;
        public string _departmentDescription;
        public string _departmentBossCode;
    }

  /*  struct PersonCodeEmail
    {
        public string _departmentBossCode;
        public string _departmentBossEmail;
    }*/

    struct DepartmentFull
    {
        public string _departmentId;
        public string _departmentDescription;
        public string _departmentBossCode;
        public string _departmentBossEmail;
    }

    struct ObjectsOfConfig
    {
        public string _parameter;
        public string _value;
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
