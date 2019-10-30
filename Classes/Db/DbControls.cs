using System;
using System.Data;

namespace ASTA
{
    public abstract class DBConnector
    {
        public virtual void CloseConnection() { }
        public virtual void Dispose() { }
    }

    //SQL
    class SqlDbReader : IDisposable
    {
        System.Data.SqlClient.SqlConnection _sqlConnection;
        System.Data.SqlClient.SqlCommand _sqlCommand;
        static string _dbConnectionString;

        public SqlDbReader(string dbConnectionString)
        {
            _dbConnectionString = dbConnectionString;

            CheckDB(_dbConnectionString);
            ConnectToDB(_dbConnectionString);
        }

        private void CheckDB(string dbConnectionString)
        {
            if (dbConnectionString?.Trim()?.Length < 1)
            {
                throw new System.ArgumentException("Connection string can not be empty or short", "dbConnectionString");
            }
        }

        private void ConnectToDB(string dbConnectionString)
        {
            if (_sqlConnection != null)
            {
                Dispose();
            }
            _dbConnectionString = dbConnectionString;
            _sqlConnection = new System.Data.SqlClient.SqlConnection(_dbConnectionString);
            _sqlConnection.Open();
        }

        public System.Data.SqlClient.SqlDataReader GetData(string sqlQuery)
        {
            _sqlCommand = new System.Data.SqlClient.SqlCommand(sqlQuery, _sqlConnection);
            return _sqlCommand.ExecuteReader();
        }

        public void CloseConnection()
        {
            if (_sqlCommand != null)
            { _sqlCommand.Dispose(); }

            if (_sqlConnection != null)
            {
                _sqlConnection.Close();
                _sqlConnection.Dispose();
            }
        }

        public void Dispose()
        {
            CloseConnection();
        }
    }


    //MySQL
    class MySqlDbReader : IDisposable
    {
        MySql.Data.MySqlClient.MySqlConnection sqlConnection;
        MySql.Data.MySqlClient.MySqlCommand sqlCommand;

        public MySqlDbReader(string dbConnectionString)
        {
            if (dbConnectionString?.Length > 0)
            {
                ConnectToDB(dbConnectionString);
            }
            else
            {
                new NullReferenceException();
            }
        }

        private void ConnectToDB(string dbConnectionString)
        {
            sqlConnection = new MySql.Data.MySqlClient.MySqlConnection(dbConnectionString);
            sqlConnection.Open();
            /*
             try
    {
        connection.Open();
        return true;
    }
    catch (MySqlException ex)
    {
        //The two most common error numbers when connecting are as follows:
        //0: Cannot connect to server.
        //1045: Invalid user name and/or password.
        switch (ex.Number)
        {
            case 0:
                MessageBox.Show("Cannot connect to server.  Contact administrator");
                break;

            case 1045:
                MessageBox.Show("Invalid username/password, please try again");
                break;
        }
        return false;
    }
             */
        }

        public MySql.Data.MySqlClient.MySqlDataReader GetData(string sqlQuery)
        {
            sqlCommand = new MySql.Data.MySqlClient.MySqlCommand(sqlQuery, sqlConnection);
            return sqlCommand.ExecuteReader();
        }

        public void CloseDataAndConnection()
        {
            if (sqlCommand != null)
            { sqlCommand.Dispose(); }

            if (sqlConnection != null)
            {
                sqlConnection.Close();
                sqlConnection.Dispose();
            }
        }

        /*
         try
    {
        connection.Close();
        return true;
    }
    catch (MySqlException ex)
    {
        MessageBox.Show(ex.Message);
        return false;
    }
             */

        public void Dispose()
        {
            CloseDataAndConnection();
        }
    }


    //SQLite
    public abstract class SQLiteDbBase : DBConnector, IDisposable
    {
        public System.Data.SQLite.SQLiteConnection _sqlConnection;
        public System.Data.SQLite.SQLiteCommand _sqlCommand;
        string _dbConnectionString;

        public SQLiteDbBase(string dbConnectionString, System.IO.FileInfo dbFileInfo)
        {
            _dbConnectionString = dbConnectionString;
            CheckDB(_dbConnectionString, dbFileInfo);
            ConnectToDB(_dbConnectionString);
        }

        public string GetConnectionString()
        {
            return _dbConnectionString;
        }

        private void CheckDB(string dbConnectionString, System.IO.FileInfo dbFileInfo)
        {
            if (!dbFileInfo.Exists)
            {
                throw new System.ArgumentException("DB is not Exist", "dbFileInfo");
            }

            if (dbConnectionString?.Trim()?.Length < 1)
            {
                throw new System.ArgumentException("Connection string can not be Empty or short", "dbConnectionString");
            }
        }

        private void ConnectToDB(string dbConnectionString)
        {
            _sqlConnection = new System.Data.SQLite.SQLiteConnection(dbConnectionString);
            _sqlConnection.Open();
        }

        public override void CloseConnection()
        {
            if (_sqlCommand != null)
            { _sqlCommand?.Dispose(); }

            if (_sqlConnection != null)
            {
                try
                {
                    _sqlConnection?.Close();
                    _sqlConnection?.Dispose();
                }
                catch { }
            }
        }

        public override void Dispose()
        { CloseConnection(); }
    }

    class SqLiteDbReader : SQLiteDbBase, IDisposable
    {
        public SqLiteDbReader(string dbConnectionString, System.IO.FileInfo dbFileInfo) :
            base(dbConnectionString, dbFileInfo)
        { }

        public System.Data.SQLite.SQLiteDataReader GetData(string query)
        {
            _sqlCommand = new System.Data.SQLite.SQLiteCommand(query, _sqlConnection);
            return _sqlCommand.ExecuteReader();
        }

        public DataTable GetDataTable(string query)
        {
            using (DataTable dt = new DataTable())
            {
                using (var _sqlDataAdapter = new System.Data.SQLite.SQLiteDataAdapter(query, _sqlConnection))
                {
                    _sqlDataAdapter.Fill(dt);

                    return dt;
                }
            }
        }
    }

    class SqLiteDbWriter : SQLiteDbBase, IDisposable
    {
        public string Status { get; private set; }
        string temporaryResult;
        public SqLiteDbWriter(string dbConnectionString, System.IO.FileInfo dbFileInfo) :
            base(dbConnectionString, dbFileInfo)
        { }

        public void ExecuteQuery(System.Data.SQLite.SQLiteCommand sqlCommand)
        {
            Status = "Ok";

            if (sqlCommand == null)
            {
                Status = "Error. The SQLCommand can not be empty or null!";
                new ArgumentNullException();
            }

            using (var sqlCommand1 = new System.Data.SQLite.SQLiteCommand("begin", _sqlConnection))
            { sqlCommand1.ExecuteNonQuery(); }

            try { sqlCommand.ExecuteNonQuery(); }
            catch (Exception expt) { Status = expt.ToString(); }

            using (var sqlCommand1 = new System.Data.SQLite.SQLiteCommand("end", _sqlConnection))
            { sqlCommand1.ExecuteNonQuery(); }
        }

        public void ExecuteQuery(string query)
        {
            Status = "Ok";
            if (query == null)
            {
                Status = "Error. The query can not be empty or null!";
                new ArgumentNullException();
            }

            using (var sqlCommand = new System.Data.SQLite.SQLiteCommand(query, _sqlConnection))
            {
                try { sqlCommand.ExecuteNonQuery(); }
                catch (Exception expt) { Status = expt.ToString(); }
            }
        }

        public void ExecuteQueryBegin()
        {
            Status = string.Empty;
            using (var sqlCommand = new System.Data.SQLite.SQLiteCommand("begin", _sqlConnection))
            { sqlCommand.ExecuteNonQuery(); }
        }

        public void ExecuteQueryEnd()
        {
            using (var sqlCommand = new System.Data.SQLite.SQLiteCommand("end", _sqlConnection))
            { sqlCommand.ExecuteNonQuery(); }
        }

        public void ExecuteQueryForBulkStepByStep(System.Data.SQLite.SQLiteCommand sqlCommand)
        {
            temporaryResult = "Ok";
            if (sqlCommand == null)
            {
                temporaryResult = "Error. The SQLCommand can not be empty or null!";
                new ArgumentNullException();
            }

            try { sqlCommand.ExecuteNonQuery(); }
            catch (Exception expt) { temporaryResult = expt.Message; }
            Status += temporaryResult;
        }

        public void ExecuteQueryForBulkStepByStep(string query)
        {
            temporaryResult = "Ok";
            if (query?.Length == 0)
            {
                temporaryResult = "Error. The SQLCommand can not be empty or null!";
                new ArgumentNullException();
            }
            using (var sqlCommand = new System.Data.SQLite.SQLiteCommand(query, _sqlConnection))
            {
                try { sqlCommand.ExecuteNonQuery(); }
                catch (Exception expt) { temporaryResult = expt.Message; }
            }
            Status += temporaryResult;
        }
    }
}
