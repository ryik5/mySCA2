using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;

using ASTA.Classes.Security;

namespace ASTA.Classes
{
    public class ParameterConfig
    {
        public string name;
        public string description;
        public string value;
        public string created;
        public bool isSecret;
        public string isExample;

        public override string ToString()
        {
            return name + "\t" + description + "\t" + value + "\t" + created + "\t" +
                isSecret.ToString() + "\t" + isExample;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            ParameterConfig df = obj as ParameterConfig;
            if (df == null)
                return false;

            return ToString() == df.ToString();
        }
        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    internal class ParameterOfConfiguration
    {
        public string name { get; set; }
        public string value { get; set; }
        public string description { get; set; }
        public string isExample { get; set; }
        public bool isSecret { get; set; }

        public static ParameterOfConfigurationBuilder CreateParameter()
        {
            return new ParameterOfConfigurationBuilder();
        }

        public override string ToString()
        {
            return name + "\t" + value + "\t" + description + "\t" + isExample + "\t" + isSecret;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            ParameterOfConfiguration df = obj as ParameterOfConfiguration;
            if (df == null)
                return false;

            return ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    internal class ParameterOfConfigurationBuilder
    {
        private ParameterOfConfiguration _parameterOfConfiguration;

        public ParameterOfConfigurationBuilder()
        {
            _parameterOfConfiguration = new ParameterOfConfiguration();
        }
        public ParameterOfConfigurationBuilder SetName(string name)
        {
            _parameterOfConfiguration.name = name;
            return this;
        }
        public ParameterOfConfigurationBuilder SetValue(string name)
        {
            _parameterOfConfiguration.value = name;
            return this;
        }
        public ParameterOfConfigurationBuilder SetDescription(string name)
        {
            _parameterOfConfiguration.description = name;
            return this;
        }
        public ParameterOfConfigurationBuilder SetIsExample(string name)
        {
            _parameterOfConfiguration.isExample = name;
            return this;
        }
        public ParameterOfConfigurationBuilder SetIsSecret(bool state)
        {
            _parameterOfConfiguration.isSecret = state;
            return this;
        }

        public ParameterOfConfigurationBuilder SetParameter(ParameterConfig parameter)
        {
            _parameterOfConfiguration.name = parameter.name;
            _parameterOfConfiguration.value = parameter.value;
            _parameterOfConfiguration.description = parameter.description;
            _parameterOfConfiguration.isExample = parameter.isExample;
            _parameterOfConfiguration.isSecret = parameter.isSecret;
            return this;
        }



        public static implicit operator ParameterOfConfiguration(ParameterOfConfigurationBuilder parameter)
        {
            return parameter._parameterOfConfiguration;
        }
    }

    internal class ConfigurationOfASTA
    {
        ParameterOfConfiguration _parameterOfConfiguration;
        System.IO.FileInfo _databasePerson;

        public delegate void Status(object sender, TextEventArgs e);
        public event Status status;

        readonly byte[] keyEncryption = Convert.FromBase64String(@"OCvesunvXXsxtt381jr7vp3+UCwDbE4ebdiL1uinGi0="); //Key Encrypt
        readonly byte[] keyDencryption = Convert.FromBase64String(@"NO6GC6Zjl934Eh8MAJWuKQ=="); //Key Decrypt

        internal ConfigurationOfASTA(System.IO.FileInfo fileinfoOfMainDB)
        {
            _databasePerson = fileinfoOfMainDB;
        }

        internal string SaveParameter(ParameterOfConfiguration parameterOfConfiguration)
        {
            _parameterOfConfiguration = parameterOfConfiguration;

            if (_parameterOfConfiguration.name == null || _parameterOfConfiguration.name.Length == 0)
                throw new ArgumentNullException("any of Parameters should have a name");

            if (_databasePerson.Exists)
            {
                using (SQLiteConnection sqlConnection = new SQLiteConnection($"Data Source={_databasePerson};Version=3;"))
                {
                    sqlConnection.Open();

                    SQLiteCommand sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                    string ParameterValueSave = "";
                    if (_parameterOfConfiguration.isSecret)
                    {
                        ParameterValueSave =
                            (_parameterOfConfiguration.value == null || _parameterOfConfiguration.value.Length == 0) ?
                            "" :
                            EncryptionDecryptionCriticalData.EncryptStringToBase64Text(_parameterOfConfiguration.value, keyEncryption, keyDencryption);
                    }
                    else
                    { ParameterValueSave = _parameterOfConfiguration.value; }

                    using (SQLiteCommand sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'ConfigDB' (ParameterName, Value, Description, DateCreated, IsPassword, IsExample)" +
                               " VALUES (@ParameterName, @Value, @Description, @DateCreated, @IsPassword, @IsExample)", sqlConnection))
                    {
                        sqlCommand.Parameters.Add("@ParameterName", DbType.String).Value = _parameterOfConfiguration.name;
                        sqlCommand.Parameters.Add("@Value", DbType.String).Value = ParameterValueSave;
                        sqlCommand.Parameters.Add("@Description", DbType.String).Value = _parameterOfConfiguration.description;
                        sqlCommand.Parameters.Add("@DateCreated", DbType.String).Value = DateTime.Now.ToYYYYMMDDHHMM();
                        sqlCommand.Parameters.Add("@IsPassword", DbType.Boolean).Value = _parameterOfConfiguration.isSecret;
                        sqlCommand.Parameters.Add("@IsExample", DbType.String).Value = _parameterOfConfiguration.isExample;
                        try { sqlCommand.ExecuteNonQuery(); } catch { }
                    }
                    sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                }
                if (_parameterOfConfiguration?.value?.Length == 0)
                    return _parameterOfConfiguration.name + " - was saved in DB, but value is empty!";
                else
                    return _parameterOfConfiguration.name + " (" + _parameterOfConfiguration.value + ")" + " - was saved in DB!";
            }
            else
            {
                return _parameterOfConfiguration.name + " - wasn't saved!\nDB: " + _databasePerson.ToString() + " is not exist!";
            }

        }

        internal List<ParameterConfig> GetParameters(string parameter)
        {
            List<ParameterConfig> parametersConfig = new List<ParameterConfig>(); ;

            ParameterConfig parameterConfig = new ParameterConfig() { name = "", description = "", value = "", isSecret = false, isExample = "no" };

            string value, valueTmp, name, decrypt;

            if (_databasePerson.Exists)
            {
                using (SQLiteConnection sqlConnection = new SQLiteConnection($"Data Source={_databasePerson};Version=3;"))
                {
                    sqlConnection.Open();

                    try
                    {
                        using (SQLiteCommand sqlCommand = sqlConnection.CreateCommand())
                        {
                            if (parameter?.Length > 0 && parameter != "%%")
                            {
                                sqlCommand.CommandText = @"Select ParameterName, Value, Description, DateCreated, IsPassword, IsExample from ConfigDB where ParameterName=@parameter";
                                sqlCommand.Parameters.Add(new SQLiteParameter("@parameter") { Value = parameter });
                            }
                            else
                            {
                                sqlCommand.CommandText = @"Select ParameterName, Value, Description, DateCreated, IsPassword, IsExample from ConfigDB";
                            }
                            sqlCommand.CommandType = System.Data.CommandType.Text;

                            status?.Invoke(this, new TextEventArgs("ParameterOfConfigurationInSQLiteDB: " + sqlCommand.CommandText));

                            using (SQLiteDataReader reader = sqlCommand.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    name = reader["ParameterName"]?.ToString();
                                    if (name?.Length > 0)
                                    {
                                        parameterConfig = new ParameterConfig() { name = "", description = "", value = "", isSecret = false, isExample = "no" };
                                        decrypt = reader["IsPassword"]?.ToString();

                                        valueTmp = reader["Value"]?.ToString();
                                        if (decrypt == "1" && valueTmp?.Length > 0)
                                        {
                                            value = EncryptionDecryptionCriticalData.DecryptBase64ToString(valueTmp, keyEncryption, keyDencryption);
                                            parameterConfig.isSecret = true;
                                        }
                                        else
                                        {
                                            value = valueTmp;
                                            parameterConfig.isSecret = false;
                                        }
                                        parameterConfig.name = name;
                                        parameterConfig.value = value;
                                        parameterConfig.description = reader["Description"]?.ToString();
                                        parameterConfig.created = reader["DateCreated"]?.ToString();
                                        parameterConfig.isExample = reader["IsExample"]?.ToString();

                                        parametersConfig.Add(parameterConfig);

                                        status?.Invoke(this, new TextEventArgs("Read a parameter from DB: " + name));
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception expt)
                    {
                        status?.Invoke(this, new TextEventArgs("ParameterOfConfigurationInSQLiteDB, error: " + expt.ToString()));
                    }
                }
            }
            return parametersConfig;
        }
    }
}
