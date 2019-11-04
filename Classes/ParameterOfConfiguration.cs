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
        public string Name { get; set; }
        public string Value { get; set; }
        public string Description { get; set; }
        public string IsExample { get; set; }
        public bool IsSecret { get; set; }

        public static ParameterOfConfigurationBuilder CreateParameter()
        {
            return new ParameterOfConfigurationBuilder();
        }

        public override string ToString()
        {
            return Name + "\t" + Value + "\t" + Description + "\t" + IsExample + "\t" + IsSecret;
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
            _parameterOfConfiguration.Name = name;
            return this;
        }
        public ParameterOfConfigurationBuilder SetValue(string name)
        {
            _parameterOfConfiguration.Value = name;
            return this;
        }
        public ParameterOfConfigurationBuilder SetDescription(string name)
        {
            _parameterOfConfiguration.Description = name;
            return this;
        }
        public ParameterOfConfigurationBuilder SetIsExample(string name)
        {
            _parameterOfConfiguration.IsExample = name;
            return this;
        }
        public ParameterOfConfigurationBuilder SetIsSecret(bool state)
        {
            _parameterOfConfiguration.IsSecret = state;
            return this;
        }

        public ParameterOfConfigurationBuilder SetParameter(ParameterConfig parameter)
        {
            _parameterOfConfiguration.Name = parameter.name;
            _parameterOfConfiguration.Value = parameter.value;
            _parameterOfConfiguration.Description = parameter.description;
            _parameterOfConfiguration.IsExample = parameter.isExample;
            _parameterOfConfiguration.IsSecret = parameter.isSecret;
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

        public delegate void Status<TextEventArgs>(object sender, TextEventArgs e);
        public event Status<TextEventArgs> status;

        readonly byte[] keyEncryption = Convert.FromBase64String(@"OCvesunvXXsxtt381jr7vp3+UCwDbE4ebdiL1uinGi0="); //Key Encrypt
        readonly byte[] keyDencryption = Convert.FromBase64String(@"NO6GC6Zjl934Eh8MAJWuKQ=="); //Key Decrypt

        internal ConfigurationOfASTA(System.IO.FileInfo fileinfoOfMainDB)
        {
            _databasePerson = fileinfoOfMainDB;
        }

        internal string SaveParameter(ParameterOfConfiguration parameterOfConfiguration)
        {
            _parameterOfConfiguration = parameterOfConfiguration;

            if (_parameterOfConfiguration.Name == null || _parameterOfConfiguration.Name.Length == 0)
                throw new ArgumentNullException("any of Parameters should have a name");

            if (_databasePerson.Exists)
            {
                using (SQLiteConnection sqlConnection = new SQLiteConnection($"Data Source={_databasePerson};Version=3;"))
                {
                    sqlConnection.Open();

                    using (SQLiteCommand sqlCommand = new SQLiteCommand("begin", sqlConnection))
                    { sqlCommand.ExecuteNonQuery(); }
                    string ParameterValueSave = "";
                    if (_parameterOfConfiguration.IsSecret)
                    {
                        ParameterValueSave =
                            (_parameterOfConfiguration.Value == null || _parameterOfConfiguration.Value.Length == 0) ?
                            "" :
                            EncryptionDecryptionCriticalData.EncryptStringToBase64Text(_parameterOfConfiguration.Value, keyEncryption, keyDencryption);
                    }
                    else
                    { ParameterValueSave = _parameterOfConfiguration.Value; }

                    using (SQLiteCommand sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'ConfigDB' (ParameterName, Value, Description, DateCreated, IsPassword, IsExample)" +
                               " VALUES (@ParameterName, @Value, @Description, @DateCreated, @IsPassword, @IsExample)", sqlConnection))
                    {
                        sqlCommand.Parameters.Add("@ParameterName", DbType.String).Value = _parameterOfConfiguration.Name;
                        sqlCommand.Parameters.Add("@Value", DbType.String).Value = ParameterValueSave;
                        sqlCommand.Parameters.Add("@Description", DbType.String).Value = _parameterOfConfiguration.Description;
                        sqlCommand.Parameters.Add("@DateCreated", DbType.String).Value = DateTime.Now.ToYYYYMMDDHHMM();
                        sqlCommand.Parameters.Add("@IsPassword", DbType.Boolean).Value = _parameterOfConfiguration.IsSecret;
                        sqlCommand.Parameters.Add("@IsExample", DbType.String).Value = _parameterOfConfiguration.IsExample;
                        try { sqlCommand.ExecuteNonQuery(); } catch { }
                    }
                    using (SQLiteCommand sqlCommand = new SQLiteCommand("end", sqlConnection))
                    { sqlCommand.ExecuteNonQuery(); }
                }
                if (_parameterOfConfiguration?.Value?.Length == 0)
                    return _parameterOfConfiguration.Name + " - was saved in DB, but value is empty!";
                else
                    return _parameterOfConfiguration.Name + " (" + _parameterOfConfiguration.Value + ")" + " - was saved in DB!";
            }
            else
            {
                return _parameterOfConfiguration.Name + " - wasn't saved!\nDB: " + _databasePerson.ToString() + " is not exist!";
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
