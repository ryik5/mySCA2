using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;

namespace ASTA
{

    internal class ParameterConfig
    {
        public string parameterName;
        public string parameterDescription;
        public string Value;
        public string dateCreated;
        public bool isPassword;
        public string isExample;

        public override string ToString()
        {
            return parameterName + "\t" + parameterDescription + "\t" + Value + "\t" + dateCreated + "\t" +
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

    internal class ParameterOfConfiguration
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

        public override string ToString()
        {
            return ParameterName + "\t" + ParameterValue + "\t" + ParameterDescription + "\t" + isExample + "\t" + isPassword;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            ParameterOfConfiguration df = obj as ParameterOfConfiguration;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
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
        public ParameterOfConfigurationBuilder SetParameterName(string name)
        {
            _parameterOfConfiguration.ParameterName = name;
            return this;
        }
        public ParameterOfConfigurationBuilder SetParameterValue(string name)
        {
            _parameterOfConfiguration.ParameterValue = name;
            return this;
        }
        public ParameterOfConfigurationBuilder SetParameterDescription(string name)
        {
            _parameterOfConfiguration.ParameterDescription = name;
            return this;
        }
        public ParameterOfConfigurationBuilder SetIsExample(string name)
        {
            _parameterOfConfiguration.isExample = name;
            return this;
        }
        public ParameterOfConfigurationBuilder IsPassword(bool state)
        {
            _parameterOfConfiguration.isPassword = state;
            return this;
        }

        public static implicit operator ParameterOfConfiguration(ParameterOfConfigurationBuilder parameter)
        {
            return parameter._parameterOfConfiguration;
        }
    }

    internal class ParameterOfConfigurationInSQLiteDB
    {
        ParameterOfConfiguration _parameterOfConfiguration;
        System.IO.FileInfo _databasePerson;

        static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        readonly byte[] keyEncryption = Convert.FromBase64String(@"OCvesunvXXsxtt381jr7vp3+UCwDbE4ebdiL1uinGi0="); //Key Encrypt
        readonly byte[] keyDencryption = Convert.FromBase64String(@"NO6GC6Zjl934Eh8MAJWuKQ=="); //Key Decrypt

        public ParameterOfConfigurationInSQLiteDB(System.IO.FileInfo fileinfoOfMainDB)
        {
            _databasePerson = fileinfoOfMainDB;
        }

        public string SaveParameter(ParameterOfConfiguration parameterOfConfiguration)
        {
            _parameterOfConfiguration = parameterOfConfiguration;

            if (_parameterOfConfiguration.ParameterName == null || _parameterOfConfiguration.ParameterName.Length == 0)
                throw new ArgumentNullException("any of Parameters should have a name");

            if (_databasePerson.Exists)
            {
                using (SQLiteConnection sqlConnection = new SQLiteConnection($"Data Source={_databasePerson};Version=3;"))
                {
                    sqlConnection.Open();

                    SQLiteCommand sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                    string ParameterValueSave = "";
                    if (_parameterOfConfiguration.isPassword)
                    {
                        ParameterValueSave =
                            (_parameterOfConfiguration.ParameterValue == null || _parameterOfConfiguration.ParameterValue.Length == 0) ?
                            "" :
                            EncryptionDecryptionCriticalData.EncryptStringToBase64Text(_parameterOfConfiguration.ParameterValue, keyEncryption, keyDencryption);
                    }
                    else
                    { ParameterValueSave = _parameterOfConfiguration.ParameterValue; }

                    using (SQLiteCommand sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'ConfigDB' (ParameterName, Value, Description, DateCreated, IsPassword, IsExample)" +
                               " VALUES (@ParameterName, @Value, @Description, @DateCreated, @IsPassword, @IsExample)", sqlConnection))
                    {
                        sqlCommand.Parameters.Add("@ParameterName", DbType.String).Value = _parameterOfConfiguration.ParameterName;
                        sqlCommand.Parameters.Add("@Value", DbType.String).Value = ParameterValueSave;
                        sqlCommand.Parameters.Add("@Description", DbType.String).Value = _parameterOfConfiguration.ParameterDescription;
                        sqlCommand.Parameters.Add("@DateCreated", DbType.String).Value = DateTime.Now.ToYYYYMMDDHHMM();
                        sqlCommand.Parameters.Add("@IsPassword", DbType.Boolean).Value = _parameterOfConfiguration.isPassword;
                        sqlCommand.Parameters.Add("@IsExample", DbType.String).Value = _parameterOfConfiguration.isExample;
                        try { sqlCommand.ExecuteNonQuery(); } catch { }
                    }
                    sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                }
                if (_parameterOfConfiguration?.ParameterValue?.Length == 0)
                    return _parameterOfConfiguration.ParameterName + " - was saved in DB, but value is empty!";
                else
                    return _parameterOfConfiguration.ParameterName + " (" + _parameterOfConfiguration.ParameterValue + ")" + " - was saved in DB!";
            }
            else
            {
                return _parameterOfConfiguration.ParameterName + " - wasn't saved!\nDB: " + _databasePerson.ToString() + " is not exist!";
            }

        }

        public List<ParameterConfig> GetParameters(string parameter)
        {
            List<ParameterConfig> parametersConfig = new List<ParameterConfig>(); ;

            ParameterConfig parameterConfig = new ParameterConfig() { parameterName = "", parameterDescription = "", Value = "", isPassword = false, isExample = "no" };

            string value = ""; string valueTmp = ""; string name = ""; string decrypt = "";

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

                            logger.Trace("ParameterOfConfigurationInSQLiteDB: " + sqlCommand.CommandText);

                            using (SQLiteDataReader reader = sqlCommand.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    name = reader["ParameterName"]?.ToString();
                                    if (name?.Length > 0)
                                    {
                                        parameterConfig = new ParameterConfig() { parameterName = "", parameterDescription = "", Value = "", isPassword = false, isExample = "no" };
                                        decrypt = reader["IsPassword"]?.ToString();

                                        valueTmp = reader["Value"]?.ToString();
                                        if (decrypt == "1" && valueTmp?.Length > 0)
                                        {
                                            value = EncryptionDecryptionCriticalData.DecryptBase64ToString(valueTmp, keyEncryption, keyDencryption);
                                            parameterConfig.isPassword = true;
                                        }
                                        else
                                        {
                                            value = valueTmp;
                                            parameterConfig.isPassword = false;
                                        }
                                        parameterConfig.parameterName = name;
                                        parameterConfig.Value = value;
                                        parameterConfig.parameterDescription = reader["Description"]?.ToString();
                                        parameterConfig.dateCreated = reader["DateCreated"]?.ToString();
                                        parameterConfig.isExample = reader["IsExample"]?.ToString();

                                        parametersConfig.Add(parameterConfig);
                                        logger.Trace("Read a parameter from DB: " + name);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception expt)
                    {
                        logger.Trace("ParameterOfConfigurationInSQLiteDB, error: " + expt.ToString());
                    }
                }
            }
            return parametersConfig;
        }
    }
}
