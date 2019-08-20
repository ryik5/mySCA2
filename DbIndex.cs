using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASTA
{
    /// <summary>
    /// Represents a single SQLite table index object
    /// </summary>
    public class DbIndex
    {
        #region Overrided Methods
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(this, obj))
                return true;

            DbIndex dst = obj as DbIndex;
            if (dst == null)
                return false;

            if (_tableName != dst._tableName)
                return false;
            if (_isUnique != dst._isUnique)
                return false;
            if (_indexName != dst._indexName)
                return false;

            if (_columns.Count != dst._columns.Count)
                return false;

            foreach (string colName in _columns.Keys)
            {
                if (!dst._columns.ContainsKey(colName))
                    return false;
                if (_columns[colName] != dst._columns[colName])
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
            return "CREATE " + (_isUnique ? "UNIQUE " : string.Empty) +
                "INDEX [" + _indexName + "] ON [" + _tableName + "] (" +
                GetColumnNames(_columns) + ")";
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Get/Set the uniqueness of the index.
        /// </summary>
        public bool IsUnique
        {
            get { return _isUnique; }
            set { _isUnique = value; }
        }

        /// <summary>
        /// The name of the index
        /// </summary>
        public string IndexName
        {
            get { return _indexName; }
            set { _indexName = value; }
        }

        /// <summary>
        /// The name of the table on which the index operates.
        /// </summary>
        public string TableName
        {
            get { return _tableName; }
            set { _tableName = value; }
        }

        /// <summary>
        /// The index columns
        /// </summary>
        public Dictionary<string, DbSortOrder> IndexColumns
        {
            get { return _columns; }
            set { _columns = value; }
        }
        #endregion

        #region Private Methods
        private string GetColumnNames(Dictionary<string, DbSortOrder> cols)
        {
            StringBuilder sb = new StringBuilder();
            List<string> colNames = new List<string>(cols.Keys);
            colNames.Sort();
            for (int i = 0; i < colNames.Count; i++)
            {
                string cname = colNames[i];
                sb.Append("[" + cname + "]");
                if (cols[cname] == DbSortOrder.Descending)
                    sb.Append(" DESC");
                if (i < colNames.Count - 1)
                    sb.Append(", ");
            } // for

            return sb.ToString();
        }
        #endregion

        #region Private Variables
        private bool _isUnique;
        private string _indexName;
        private string _tableName;
        private Dictionary<string, DbSortOrder> _columns;
        #endregion
    }

    /// <summary>
    /// List the possible sort orders
    /// </summary>
    public enum DbSortOrder
    {
        /// <summary>
        /// Invalid value
        /// </summary>
        None = 0,

        /// <summary>
        /// ASCENDING sort order
        /// </summary>
        Ascending = 1,

        /// <summary>
        /// DESCENDING sort order
        /// </summary>
        Descending = 2,
    }
}
