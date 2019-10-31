using System.Collections.Generic;
using System.Text;

namespace ASTA.Classes
{
    // <summary>
    /// Represents a single SQLite table.
    /// </summary>
    public class DbTable
    {
        #region Overrided Methods
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(this, obj))
                return true;

            DbTable dst = obj as DbTable;
            if (dst == null)
                return false;

            if (_tableName != dst._tableName)
                return false;

            if (_columns.Count != dst._columns.Count)
                return false;

            foreach (DbColumn col in _columns)
            {
                bool found = false;
                foreach (DbColumn dcol in dst._columns)
                {
                    if (dcol.ColumnName == col.ColumnName)
                    {
                        found = true;
                        if (!dcol.Equals(col))
                            return false;
                    }
                } // foreach
                if (!found)
                    return false;
            } // foreach

            return true;
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        public override string ToString()
        {
            return GetCreateTableStatement(_tableName);
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Returns a CREATE TABLE DDL statement with the specified table name.
        /// </summary>
        /// <param name="tableName">The name of the table</param>
        /// <returns>The CREATE TABLE SQL statement</returns>
        public string GetCreateTableStatement(string tableName)
        {
            if (_primaryKeys.Count == 1 && _primaryKeys[new List<string>(_primaryKeys.Keys)[0]].IsAutoIncrement)
            {
                return "CREATE TABLE [" + tableName + "] (\r\n" +
                    GetColumnStatements(_columns) + ");\r\n";
            }
            else
            {
                return "CREATE TABLE [" + tableName + "] (\r\n" +
                    GetColumnStatements(_columns) +
                    GetPrimaryKeySection(_primaryKeys) + ");\r\n";
            }
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// The table name
        /// </summary>
        public string TableName
        {
            get { return _tableName; }
            set { _tableName = value; }
        }

        /// <summary>
        /// Set/Get the database column objects of the table.
        /// </summary>
        public List<DbColumn> Columns
        {
            get { return _columns; }
            set
            {
                _columns = value;
                _primaryKeys = new Dictionary<string, DbColumn>();
                foreach (DbColumn col in value)
                {
                    if (col.IsPrimaryKey)
                        _primaryKeys.Add(col.ColumnName, col);
                } // foreach
            }
        }

        /// <summary>
        /// Get/Set the primary keys of the table.
        /// </summary>
        public Dictionary<string, DbColumn> PrimaryKeys
        {
            get { return _primaryKeys; }
            set { _primaryKeys = value; }
        }
        #endregion

        #region Private Methods
        private string GetColumnStatements(List<DbColumn> columns)
        {
            Dictionary<string, DbColumn> mcols = new Dictionary<string, DbColumn>();
            List<string> colNames = new List<string>();
            foreach (DbColumn col in columns)
            {
                colNames.Add(col.ColumnName);
                mcols.Add(col.ColumnName, col);
            }
            colNames.Sort();

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < colNames.Count; i++)
            {
                string cname = colNames[i];
                sb.Append(mcols[cname].ToString());
                if (i < colNames.Count - 1)
                    sb.Append(",\n");
            } // for

            return sb.ToString();
        }

        private string GetPrimaryKeySection(Dictionary<string, DbColumn> pkeys)
        {
            List<string> pnames = new List<string>(pkeys.Keys);
            if (pnames.Count == 0)
                return string.Empty;
            pnames.Sort();

            StringBuilder sb = new StringBuilder();
            sb.Append(",\nPRIMARY KEY (");
            for (int i = 0; i < pnames.Count; i++)
            {
                sb.Append("[" + pnames[i] + "]");
                if (i < pnames.Count - 1)
                    sb.Append(", ");
            } // for
            sb.Append(")\n");

            return sb.ToString();
        }

        #endregion

        #region Private Variables
        private string _tableName;
        private List<DbColumn> _columns;
        private Dictionary<string, DbColumn> _primaryKeys;
        #endregion
    }
}
