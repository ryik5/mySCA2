using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASTA
{
    class SqlDbTableReader : IDisposable
    {
        string _dbConnectionString = null;
        string _sqlQuery = null;
        System.Data.SqlClient.SqlConnection sqlConnection;
        System.Data.SqlClient.SqlCommand sqlCommand;

        public SqlDbTableReader(string dbConnectionString)
        {
            _dbConnectionString = dbConnectionString;
            ConnectToDB();
        }

        private void ConnectToDB()
        {
            sqlConnection = new System.Data.SqlClient.SqlConnection(_dbConnectionString);
            sqlConnection.Open();
        }

        public System.Data.SqlClient.SqlDataReader GetDataFromDB(string sqlQuery)
        {
            _sqlQuery = sqlQuery;
            sqlCommand = new System.Data.SqlClient.SqlCommand(_sqlQuery, sqlConnection);
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

        public void Dispose()
        {
            CloseDataAndConnection();
        }

    }

    class MySqlDbTableReader : IDisposable
    {
        private string _dbConnectionString = null;
        private string _sqlQuery = null;
        MySql.Data.MySqlClient.MySqlConnection sqlConnection;
        MySql.Data.MySqlClient.MySqlCommand sqlCommand;

        public MySqlDbTableReader(string dbConnectionString)
        {
            _dbConnectionString = dbConnectionString;
            ConnectToDB();
        }

        private void ConnectToDB()
        {
            sqlConnection = new MySql.Data.MySqlClient.MySqlConnection(_dbConnectionString);
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
            _sqlQuery = sqlQuery;
            sqlCommand = new MySql.Data.MySqlClient.MySqlCommand(_sqlQuery, sqlConnection);
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

}
