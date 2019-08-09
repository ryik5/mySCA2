using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;

namespace ASTA
{

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
                throw new ArgumentNullException("any of Parameters should have a name");

            if (databasePerson.Exists)
            {
                using (SQLiteConnection sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();

                    SQLiteCommand sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                    string ParameterValueSave = "";
                    if (parameterOfConfiguration.isPassword)
                    {
                        ParameterValueSave =
                            (parameterOfConfiguration.ParameterValue == null || parameterOfConfiguration.ParameterValue.Length == 0) ?
                            "" :
                            EncryptionDecryptionCriticalData.EncryptStringToBase64Text(parameterOfConfiguration.ParameterValue, btsMess1, btsMess2);
                    }
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
                return parameterOfConfiguration.ParameterName + " - wasn't saved!\n" + databasePerson.ToString() + " is not exist!";
            }

        }

        public List<ParameterConfig> GetParameters(string parameter)
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

}
