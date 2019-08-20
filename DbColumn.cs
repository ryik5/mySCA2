using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASTA
{
    public class DbColumn
    {
        #region Overrided Methods
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(this, obj))
                return true;

            DbColumn dst = obj as DbColumn;
            if (dst == null)
                return false;

            if (_collation != dst._collation)
                return false;
            if (_columnName != dst._columnName)
                return false;
            if (_dbType != dst._dbType)
                return false;
            if (_defaultValue == null && dst._defaultValue != null ||
                _defaultValue != null && dst._defaultValue == null ||
                (_defaultValue != null && dst._defaultValue != null &&
                !_defaultValue.Equals(dst._defaultValue)))
                return false;
            if (_defaultFunction != dst._defaultFunction)
                return false;
            if (_isNullable != dst._isNullable)
                return false;
            if (_isPrimaryKey != dst._isPrimaryKey)
                return false;
            if (_size != dst._size)
                return false;
            if (_precision != dst._precision)
                return false;
            if (_isUnique != dst._isUnique)
                return false;
            if (_isAutoIncrement != dst._isAutoIncrement)
                return false;

            return true;
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        public override string ToString()
        {
            if (_isPrimaryKey && _isAutoIncrement)
            {
                return "[" + _columnName + "] " +
                    GetSqlType(_dbType, _size, _precision) +
                    " PRIMARY KEY AUTOINCREMENT " +
                    (_isNullable ? " NULL" : " NOT NULL") +
                    GetDefaultValueString() +
                    (_isUnique ? " UNIQUE" : string.Empty) +
                    (_collation != null ? (" COLLATE " + _collation) : string.Empty);
            }
            else
            {
                return "[" + _columnName + "] " +
                    GetSqlType(_dbType, _size, _precision) +
                    (_isNullable ? " NULL" : " NOT NULL") +
                    GetDefaultValueString() +
                    (_isUnique ? " UNIQUE" : string.Empty) +
                    (_collation != null ? (" COLLATE " + _collation) : string.Empty);
            }
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Get/Set the column name of the column.
        /// </summary>
        public string ColumnName
        {
            get { return _columnName; }
            set { _columnName = value; }
        }

        /// <summary>
        /// Get/Set the primary key property of the column.
        /// </summary>
        public bool IsPrimaryKey
        {
            get { return _isPrimaryKey; }
            set { _isPrimaryKey = value; }
        }

        /// <summary>
        /// Get/Set indication of the primary is using the auto-incrementing algorithm
        /// </summary>
        public bool IsAutoIncrement
        {
            get { return _isAutoIncrement; }
            set { _isAutoIncrement = value; }
        }

        /// <summary>
        /// Get/Set the DB type of the column.
        /// </summary>
        public DbType ColumnType
        {
            get { return _dbType; }
            set { _dbType = value; }
        }

        /// <summary>
        /// Get/Set the size of the column object
        /// </summary>
        public int ColumnSize
        {
            get { return _size; }
            set { _size = value; }
        }

        /// <summary>
        /// Get/Set the precision of the column
        /// </summary>
        public int ColumnPrecision
        {
            get { return _precision; }
            set { _precision = value; }
        }

        /// <summary>
        /// Get/Set the nullability property of the column (NULL / NOT NULL)
        /// </summary>
        public bool IsNullable
        {
            get { return _isNullable; }
            set { _isNullable = value; }
        }

        /// <summary>
        /// Get/Set the default value of the column.
        /// </summary>
        public object DefaultValue
        {
            get { return _defaultValue; }
            set { _defaultValue = value; }
        }

        /// <summary>
        /// Get/Set the default function of the column.
        /// </summary>
        public string DefaultFunction
        {
            get { return _defaultFunction; }
            set { _defaultFunction = value; }
        }

        /// <summary>
        /// Get/Set the name of the collation of the column (NULL if no collation is
        /// associated with the column).
        /// </summary>
        public string Collation
        {
            get { return _collation; }
            set { _collation = value; }
        }

        /// <summary>
        /// The column has UNIQUE constraint.
        /// </summary>
        public bool IsUnique
        {
            get { return _isUnique; }
            set { _isUnique = value; }
        }
        #endregion

        #region Private Methods
        private string GetDefaultValueString()
        {
            if (_defaultFunction != null)
                return " DEFAULT " + _defaultFunction;

            object value = _defaultValue;
            if (value == null)
                return string.Empty;

            if (value is string)
                return " DEFAULT '" + (string)value + "'";
            else if (value is int)
                return " DEFAULT " + (int)value;
            else if (value is double)
                return " DEFAULT " + (double)value;
            else
                throw new Exception();// DbUpgradeException.SchemaIsNotSupported();
        }

        private string GetSqlType(DbType type, int size, int precision)
        {
            string stype;
            switch (type)
            {
                case DbType.AnsiString:
                    stype = "varchar";
                    break;

                case DbType.AnsiStringFixedLength:
                    stype = "char";
                    break;

                case DbType.String:
                    stype = "nvarchar";
                    break;

                case DbType.Binary:
                    stype = "blob";
                    break;

                case DbType.Boolean:
                    stype = "boolean";
                    break;

                case DbType.Byte:
                    stype = "tinyint";
                    break;

                case DbType.DateTime:
                    stype = "timestamp";
                    break;

                case DbType.Double:
                    stype = "float";
                    break;

                case DbType.Int16:
                    stype = "smallint";
                    break;

                case DbType.Int32:
                    stype = "int";
                    break;

                case DbType.Int64:
                    stype = "integer";
                    break;

                default:
                    throw new Exception();// DbUpgradeException.SchemaIsNotSupported();
            } // switch

            if (size != -1)
            {
                if (precision != -1)
                    return stype + "(" + size + "," + precision + ")";
                else
                    return stype + "(" + size + ")";
            }
            else
                return stype;
        }
        #endregion


        #region Private Variables
        private int _size = -1;
        private int _precision = -1;
        private DbType _dbType;
        private bool _isNullable;
        private object _defaultValue;
        private string _columnName;
        private string _collation;
        private bool _isPrimaryKey;
        private bool _isUnique;
        private string _defaultFunction;
        private bool _isAutoIncrement;
        #endregion
    }
}
