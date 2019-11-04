﻿using System;
using System.Data;

namespace ASTA.Classes
{
    //SQL
    class SqlDbReader : IDisposable
    {
        System.Data.SqlClient.SqlConnection _sqlConnection;
        static string _dbConnectionString;

        public delegate void Message(object sender, TextEventArgs e);
        public event Message Status;

        public SqlDbReader(string dbConnectionString)
        {
            _dbConnectionString = dbConnectionString;

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

        public System.Data.SqlClient.SqlDataReader GetData(string query)
        {
            CheckDB(_dbConnectionString);
            ConnectToDB(_dbConnectionString);

            Status?.Invoke(this, new TextEventArgs("query: " + query));

            using (var sqlCommand = new System.Data.SqlClient.SqlCommand(query, _sqlConnection))
            { return sqlCommand.ExecuteReader(); }
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        private void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
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

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // This code added to correctly implement the disposable pattern.
        public virtual void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }
        #endregion
    }


        //MySQL
        class MySqlDbReader : IDisposable
    {
        MySql.Data.MySqlClient.MySqlConnection sqlConnection;
        public delegate void Message(object sender, TextEventArgs e);
        public event Message Status;
        string _dbConnectionString;

        public MySqlDbReader(string dbConnectionString)
        {
            _dbConnectionString = dbConnectionString;
        }

        private void ConnectToDB(string dbConnectionString)
        {
            if (dbConnectionString?.Length > 0)
            {
                if (sqlConnection != null)
                {
                    Dispose();
                }

                sqlConnection = new MySql.Data.MySqlClient.MySqlConnection(dbConnectionString);
                sqlConnection.Open();
            }
            else
            {
                new NullReferenceException();
            }

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

        public MySql.Data.MySqlClient.MySqlDataReader GetData(string query)
        {
            ConnectToDB(_dbConnectionString);
            Status?.Invoke(this, new TextEventArgs("query: " + query));
            using (var sqlCommand = new MySql.Data.MySqlClient.MySqlCommand(query, sqlConnection))
            { return sqlCommand.ExecuteReader(); }
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (sqlConnection != null)
                    {
                        try
                        {
                            sqlConnection?.Close();
                            sqlConnection?.Dispose();
                        }
                        catch { }
                    }
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }
        #endregion
    }


    //SQLite
    public abstract class SQLiteDbBase :  IDisposable
    {
        public delegate void Message(object sender, TextEventArgs e);

        public System.Data.SQLite.SQLiteConnection sqlConnection;
        public System.Data.SQLite.SQLiteCommand _sqlCommand;
        string _dbConnectionString;

        protected SQLiteDbBase(string dbConnectionString, System.IO.FileInfo dbFileInfo)
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
            if (!(dbFileInfo?.Length>0))
                throw new System.ArgumentException("dbFileInfo cannot be null!");

            if (!dbFileInfo.Exists)
                throw new System.ArgumentException("dbFileInfo is not exist");

            if (!(dbConnectionString?.Trim()?.Length >0))
                throw new System.ArgumentException("dbConnectionString string can not be Empty or short");
        }

        private void ConnectToDB(string dbConnectionString)
        {
            if (sqlConnection != null)
            {
                Dispose();
            }

            sqlConnection = new System.Data.SQLite.SQLiteConnection(dbConnectionString);
            sqlConnection.Open();
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (sqlConnection != null)
                    {
                        try
                        {
                            sqlConnection?.Close();
                            sqlConnection?.Dispose();
                        }
                        catch { }
                    }
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }
        #endregion
    }

    internal class SqLiteDbReader : SQLiteDbBase, IDisposable
    {
        public event Message Status;

        public SqLiteDbReader(string dbConnectionString, System.IO.FileInfo dbFileInfo) :
            base(dbConnectionString, dbFileInfo)
        { }

        public System.Data.SQLite.SQLiteDataReader GetData(string query)
        {
            Status?.Invoke(this, new TextEventArgs("query: "+ query));
            using (var _sqlCommand = new System.Data.SQLite.SQLiteCommand(query, sqlConnection))
            {
                Status?.Invoke(this, new TextEventArgs("query: " + query));
                return _sqlCommand.ExecuteReader(); 
            }
        }

        public DataTable GetDataTable(string query)
        {
            using (DataTable dt = new DataTable())
            {
                using (var _sqlDataAdapter = new System.Data.SQLite.SQLiteDataAdapter(query, sqlConnection))
                {
                    Status?.Invoke(this, new TextEventArgs("query: " + query));

                    _sqlDataAdapter.Fill(dt);
                    return dt;
                }
            }
        }
    }

   internal class SqLiteDbWriter : SQLiteDbBase, IDisposable
    {
        public SqLiteDbWriter(string dbConnectionString, System.IO.FileInfo dbFileInfo) :
            base(dbConnectionString, dbFileInfo)
        { }

        public event Message Status;

        public void ExecuteQuery(System.Data.SQLite.SQLiteCommand sqlCommand)
        {
            if (sqlCommand == null)
            {
                Status?.Invoke(this, new TextEventArgs("Error. The SQLCommand can not be empty or null!"));
                new ArgumentNullException();
            }

            using (var sqlCommand1 = new System.Data.SQLite.SQLiteCommand("begin", sqlConnection))
            { sqlCommand1.ExecuteNonQuery(); }

            try
            {
                sqlCommand.ExecuteNonQuery();
                Status?.Invoke(this, new TextEventArgs("ExecuteQuery - Ok"));
            }
            catch (Exception expt)
            { Status?.Invoke(this, new TextEventArgs("Error! " + expt.ToString())); }

            using (var sqlCommand1 = new System.Data.SQLite.SQLiteCommand("end", sqlConnection))
            { sqlCommand1.ExecuteNonQuery(); }
        }

        public void ExecuteQuery(string query)
        {
            if (query == null)
            {
                Status?.Invoke(this, new TextEventArgs("Error. The query can not be empty or null!"));
                new ArgumentNullException();
            }

            using (var sqlCommand = new System.Data.SQLite.SQLiteCommand(query, sqlConnection))
            {
                try {
                    sqlCommand.ExecuteNonQuery();
                    Status?.Invoke(this, new TextEventArgs("query: " + query+" - ok"));
                }
                catch (Exception expt)
                { Status?.Invoke(this, new TextEventArgs("query: "+ query+" ->Error! " + expt.ToString())); }
            }
        }

        public void ExecuteQueryForBulkStepByStep(System.Data.SQLite.SQLiteCommand sqlCommand)
        {
            if (sqlCommand == null)
            {
                Status?.Invoke(this, new TextEventArgs("Error. The SQLCommand can not be empty or null!"));
                new ArgumentNullException();
            }

            try
            {
                sqlCommand.ExecuteNonQuery();
                Status?.Invoke(this, new TextEventArgs("ExecuteQueryForBulkStepByStep - Ok"));
            }
            catch (Exception expt)
            { Status?.Invoke(this, new TextEventArgs("ExecuteQueryForBulkStepByStep -> Error! " + expt.ToString())); }
        }
    }
}
