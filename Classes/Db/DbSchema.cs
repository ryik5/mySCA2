using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Text;
using System.Text.RegularExpressions;

namespace ASTA.Classes
{
    /// <summary>
    /// Represents a complete SQLite DB schema
    /// </summary>
    public class DbSchema
    {
        #region Object Overrides
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(this, obj))
                return true;

            DbSchema dst = obj as DbSchema;
            if (dst == null)
                return false;

            if (_tables.Count != dst._tables.Count)
                return false;
            if (_indexes.Count != dst._indexes.Count)
                return false;

            foreach (string tableName in _tables.Keys)
            {
                if (!dst._tables.ContainsKey(tableName))
                    return false;
            }

            foreach (string indexName in _indexes.Keys)
            {
                if (!dst._indexes.ContainsKey(indexName))
                    return false;
            }

            foreach (string tableName in _tables.Keys)
            {
                DbTable table1 = _tables[tableName];
                DbTable table2 = dst._tables[tableName];
                if (!table1.Equals(table2))
                    return false;
            } // foreach

            foreach (string indexName in _indexes.Keys)
            {
                DbIndex index1 = _indexes[indexName];
                DbIndex index2 = dst._indexes[indexName];
                if (!index1.Equals(index2))
                    return false;
            }

            return true;
        }

        public override int GetHashCode()
        {
            // TODO: return the hash code by doing the hash on the stringified 
            // representation of the schema.
            return base.GetHashCode();
        }

        public override string ToString()
        {
            //  TODO: return the stringified version of the schema.
            return base.ToString();
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Return all table meta data in the schema
        /// </summary>
        public Dictionary<string, DbTable> Tables
        {
            get { return _tables; }
        }

        /// <summary>
        /// Return all index meta data in the schema
        /// </summary>
        public Dictionary<string, DbIndex> Indexes
        {
            get { return _indexes; }
        }
        #endregion

        #region Public Static Methods

        /// <summary>
        /// Parse the specified SQL string and creates the corresponding
        /// DbSchema
        /// </summary>
        /// <param name="sql">The SQL string to parse</param>
        /// <returns>The DB schema that corresponds to the SQL string.</returns>
        public static DbSchema ParseSql(string sql)
        {
            DbSchema schema = new DbSchema();
            string newSQL = sql.Replace("IF NOT EXISTS", ""); //remove from query of checking to exist the table in DB

            do
            {
                Match m1 = _tableHeader.Match(newSQL);
                Match m2 = _indexHeader.Match(newSQL);
                if (m1.Success && (!m2.Success || m2.Success && m2.Index > m1.Index))
                {
                    DbTable table = ParseDbTable(ref newSQL);
                    schema._tables.Add(table.TableName, table);
                }
                else
                {
                    if (m2.Success)
                    {
                        DbIndex index = ParseDbIndex(ref newSQL);
                        schema._indexes.Add(index.IndexName, index);
                    }
                    else
                        break;
                } // else
            } while (true);

            return schema;
        }

        /// <summary>
        /// Loads the DB schema of the specified DB file
        /// </summary>
        /// <param name="dbPath">The path to the DB file to load</param>
        /// <returns>The DB schema of that file</returns>
        public static DbSchema LoadDB(string dbPath)
        {
            DbSchema schema = new DbSchema();

            SQLiteConnectionStringBuilder builder = new SQLiteConnectionStringBuilder
            {
                DataSource = dbPath,
                PageSize = 4096,
                UseUTF16Encoding = true
            };

            using (SQLiteConnection conn = new SQLiteConnection(builder.ConnectionString))
            {
                conn.Open();

                using (SQLiteCommand query = new SQLiteCommand(@"SELECT * FROM SQLITE_MASTER", conn))
                {
                    using (SQLiteDataReader reader = query.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string type = (string)reader["type"];
                            string name = (string)reader["name"];
                            string tblName = (string)reader["tbl_name"];

                            // Ignore SQLite internal indexes and tables
                            if (name.StartsWith("sqlite_"))
                                continue;
                            if (reader["sql"] == DBNull.Value)
                                continue;

                            string sql = (string)reader["sql"];

                            if (type == "table")
                                schema._tables.Add(tblName, ParseDbTable(ref sql));
                            else if (type == "index")
                                schema._indexes.Add(name, ParseDbIndex(ref sql));
                            else
                                throw DbUpgradeException.SchemaIsNotSupported();
                        } // while
                    } // using
                }// using
            } // using

            return schema;
        }
        #endregion

        #region Private Static Methods

        /// <summary>
        /// Parse the specified CREATE INDEX statement and returned the 
        /// index representation as a DbIndex instance.
        /// </summary>
        /// <param name="sql">The CREATE INDEX sql statement</param>
        /// <returns>The DbIndex representation of the table.</returns>
        private static DbIndex ParseDbIndex(ref string sql)
        {
            DbIndex index = new DbIndex();
            Match m = _indexHeader.Match(sql);
            if (m.Success)
            {
                if (m.Groups[1].Success)
                    index.IsUnique = true;

                int start = m.Index + m.Length;
                index.IndexName = ParsePotentiallyDelimitedToken(sql, ref start);
                if (index.IndexName == null)
                    throw DbUpgradeException.SchemaIsNotSupported();

                // Search for occurence of "ON"
                int offset = sql.IndexOf("ON", start);
                if (offset == -1)
                    throw DbUpgradeException.SchemaIsNotSupported();
                start = offset + 2;

                index.TableName = ParsePotentiallyDelimitedToken(sql, ref start);
                if (index.TableName == null)
                    throw DbUpgradeException.SchemaIsNotSupported();

                sql = sql.Substring(start);
                if (!ScanToken(ref sql, "("))
                    throw DbUpgradeException.SchemaIsNotSupported();

                string cols = null;
                for (int i = 0; i < sql.Length; i++)
                {
                    if (sql[i] == ')')
                    {
                        cols = sql.Substring(0, i);
                        if (i < sql.Length - 1)
                            sql = sql.Substring(i + 1);
                        else
                            sql = string.Empty;
                        break;
                    }
                } // for
                if (cols == null)
                    throw DbUpgradeException.SchemaIsNotSupported();

                string[] parts = cols.Split(',');
                Dictionary<string, DbSortOrder> icols =
                    new Dictionary<string, DbSortOrder>();
                foreach (string p in parts)
                {
                    string cn = p.Trim();
                    string[] cparts = cn.Split(' ', '\t');
                    DbSortOrder order = DbSortOrder.Ascending;
                    if (cparts.Length == 2)
                    {
                        if (cparts[1].ToUpper() == "DESC")
                            order = DbSortOrder.Descending;
                        else
                            throw DbUpgradeException.SchemaIsNotSupported();
                    }

                    icols.Add(cparts[0].Trim().Trim('[', ']', '\'', '`', '\"'), order);
                }
                if (icols.Count == 0)
                    throw DbUpgradeException.SchemaIsNotSupported();
                index.IndexColumns = icols;

                return index;
            }
            else
                throw DbUpgradeException.SchemaIsNotSupported();
        }

        /// <summary>
        /// Parses the specified CREATE TABLE SQL statement and return
        /// the table representation as a DbTable instance.
        /// </summary>
        /// <param name="sql">The CREATE TABLE sql statement</param>
        /// <returns>The DbTable representation of the table.</returns>
        public static DbTable ParseDbTable(ref string sql)
        {
            DbTable table = new DbTable
            {
                TableName = ParseTableName(ref sql)
            };
            table.Columns = ParseTableColumns(ref sql, table.TableName);
            return table;
        }

        /// <summary>
        /// Parse the table columns section of the CREATE TABLE DDL statement.
        /// </summary>
        /// <param name="sql">The SQL statement to parse</param>
        /// <param name="tableName">Name of the table.</param>
        /// <returns>
        /// The list of DbColumn objects that represent the meta
        /// data about all table columns.
        /// </returns>
        public static List<DbColumn> ParseTableColumns(ref string sql, string tableName)
        {
            // Skip '('
            if (!ScanToken(ref sql, "("))
                throw DbUpgradeException.SchemaIsNotSupported();

            List<string> primaryKeys = new List<string>();
            List<DbColumn> res = new List<DbColumn>();
            do
            {
                if (sql.Trim().StartsWith(")"))
                    break;
                DbColumn col = ParseColumn(ref sql);
                if (col != null)
                {
                    res.Add(col);
                    if (!ScanToken(ref sql, ","))
                    {
                        if (ScanToken(ref sql, ")"))
                            break;
                        else
                            DbUpgradeException.SchemaIsNotSupported();
                    } // if
                }
                else
                {
                    // Try to parse as a PRIMARY KEY section
                    List<string> keys = ParsePrimaryKeys(ref sql);
                    if (keys != null)
                        primaryKeys.AddRange(keys);
                    else
                          throw DbUpgradeException.SchemaIsNotSupported();
                } // else
            } while (true);

            // Apply any primary keys found
            foreach (string pkey in primaryKeys)
            {
                bool found = false;
                foreach (DbColumn col in res)
                {
                    if (col.ColumnName == pkey)
                    {
                        col.IsPrimaryKey = true;
                        found = true;
                        break;
                    }
                } // foreach

                if (!found)
                    throw DbUpgradeException.InvalidPrimaryKeySection(tableName);
            } // foreach

            return res;
        }

        /// <summary>
        /// Parse the specified SQL statement as a PRIMARY KEYS() section.
        /// </summary>
        /// <param name="sql">The SQL statement to parse</param>
        /// <returns>The list of primary keys (if recognized as a PRIMARY KEYS
        /// statement) or null if not recognized.</returns>
        private static List<string> ParsePrimaryKeys(ref string sql)
        {
            if (!ScanToken(ref sql, "PRIMARY KEY"))
                return null;
            if (!ScanToken(ref sql, "("))
                return null;

            string keys = null;
            for (int i = 0; i < sql.Length; i++)
            {
                if (sql[i] == ')')
                {
                    keys = sql.Substring(0, i);
                    if (i < sql.Length - 1)
                        sql = sql.Substring(i + 1);
                    else
                        sql = string.Empty;
                    break;
                }
            } // for

            if (keys != null)
            {
                string[] parts = keys.Split(',');
                List<string> res = new List<string>();
                foreach (string p in parts)
                {
                    string key = p.Trim().Trim('[', ']', '\'', '`', '\"');
                    res.Add(key);
                } // foreach

                return res;
            }
            else
                throw DbUpgradeException.SchemaIsNotSupported();
        }

        /// <summary>
        /// Parse a single table column row.
        /// </summary>
        /// <param name="sql">The SQL statement</param>
        /// <returns>The DbColumn instance for the column.</returns>
        private static DbColumn ParseColumn(ref string sql)
        {
            // In case this is a PRIMARY KEY constraint - return null immediatly
            if (sql.TrimStart(' ', '\t', '\n', '\r').StartsWith("PRIMARY KEY"))
                return null;

            string columnName = null;
            DbType columnType = DbType.Int32;
            int columnSize = -1;
            int columnPrecision = -1;
            DbColumn res = new DbColumn();
            if (ParseColumnNameType(ref sql, ref columnName, ref columnType, ref columnSize, ref columnPrecision))
            {
                res.ColumnName = columnName;
                res.ColumnSize = columnSize;
                res.ColumnType = columnType;
                res.ColumnPrecision = columnPrecision;

                int index = FindClosingRightParens(sql);
                if (index == -1)
                    throw DbUpgradeException.SchemaIsNotSupported();
                int index2 = sql.IndexOf(",");
                if (index2 != -1 && index2 < index)
                    index = index2;
                string rest = sql.Substring(0, index);
                sql = sql.Substring(index);

                // TRUE by default
                res.IsNullable = true;

                // Parse the list of column constraints
                Match m = _columnConstraints.Match(rest);
                while (m.Success)
                {
                    if (m.Groups[2].Success)
                    {
                        if (m.Groups[2].Value == "NOT NULL")
                            res.IsNullable = false;
                        else
                            res.IsNullable = true;
                    }
                    else if (m.Groups[1].Success)
                    {
                        res.IsPrimaryKey = true;
                    }
                    else if (m.Groups[5].Success)
                    {
                        res.DefaultValue = null;
                        res.DefaultFunction = m.Groups[5].Value;
                    }
                    else if (m.Groups[7].Success)
                    {
                        res.DefaultValue = m.Groups[7].Value;
                        res.DefaultFunction = null;
                    }
                    else if (m.Groups[9].Success)
                    {
                        res.DefaultValue = double.Parse(m.Groups[9].Value);
                        res.DefaultFunction = null;
                    }
                    else if (m.Groups[11].Success)
                    {
                        res.DefaultValue = int.Parse(m.Groups[11].Value);
                        res.DefaultFunction = null;
                    }
                    else if (m.Groups[13].Success)
                    {
                        res.Collation = m.Groups[13].Value;
                    }
                    else if (m.Groups[14].Success)
                    {
                        res.IsUnique = true;
                    }
                    else if (m.Groups[15].Success)
                    {
                        res.IsAutoIncrement = true;
                    }
                    else
                        throw DbUpgradeException.InternalSoftwareError();

                    rest = rest.Substring(m.Index + m.Length);
                    m = _columnConstraints.Match(rest);
                } // while                

                return res;
            } // if
            else
            {
                // This is not a column - maybe it is a primary key declaration
                return null;
            } // else
        }

        /// <summary>
        /// Try to parse the SQL string as a column name
        /// </summary>
        /// <param name="sql">The SQL string to parse</param>
        /// <param name="colName">The column name</param>
        /// <param name="colType">The column type</param>
        /// <param name="colSize">The column size</param>
        /// <returns>TRUE if parsed successfully, FALSE otherwise.</returns>
        private static bool ParseColumnNameType(ref string sql, ref string colName, ref DbType colType, ref int colSize, ref int colPrecision)
        {
            int index = 0;

            colName = ParsePotentiallyDelimitedToken(sql, ref index);
            if (colName == null)
                return false;

            string ctype = ParsePotentiallyDelimitedToken(sql, ref index);
            if (ctype == null)
                return false;
            colType = GetDbType(ctype);

            // See if there are parentheses with a digit inside
            colSize = ParsePotentialTypeSize(sql, out colPrecision, ref index);

            if (index < sql.Length)
                sql = sql.Substring(index);
            else
                sql = string.Empty;

            return true;
        }

        private static int ParsePotentialTypeSize(string sql, out int precision, ref int index)
        {
            int size = -1;
            precision = -1;
            int saved = index;

            // Skip white space
            for (; index < sql.Length; index++)
            {
                if (!Char.IsWhiteSpace(sql[index]))
                    break;
            }
            if (index == sql.Length)
            {
                index = saved;
                return -1;
            }

            if (sql[index] == '(')
            {
                // Parse the size component
                StringBuilder digit = new StringBuilder();
                for (index++; index < sql.Length; index++)
                {
                    if (!Char.IsWhiteSpace(sql[index]))
                        break;
                }
                if (index == sql.Length || !Char.IsDigit(sql[index]))
                {
                    index = saved;
                    return -1;
                }
                for (; index < sql.Length; index++)
                {
                    if (!Char.IsDigit(sql[index]))
                        break;
                    else
                        digit.Append(sql[index]);
                }
                if (index == sql.Length)
                {
                    index = saved;
                    return -1;
                }
                string tmpsize = digit.ToString();

                // Skip white space
                for (; index < sql.Length; index++)
                {
                    if (!Char.IsWhiteSpace(sql[index]))
                        break;
                }
                if (index == sql.Length)
                {
                    index = saved;
                    return -1;
                }

                if (sql[index] == ',')
                {
                    // The size has a precision component that we need to parse also.

                    // Skip white space
                    for (index++; index < sql.Length; index++)
                    {
                        if (!Char.IsWhiteSpace(sql[index]))
                            break;
                    }
                    if (index == sql.Length)
                    {
                        index = saved;
                        return -1;
                    }

                    // Read the precision component
                    digit = new StringBuilder();
                    for (; index < sql.Length; index++)
                    {
                        if (!Char.IsDigit(sql[index]))
                            break;
                        else
                            digit.Append(sql[index]);
                    }
                    if (index == sql.Length)
                    {
                        index = saved;
                        return -1;
                    }
                    string tmpprec = digit.ToString();
                    if (!int.TryParse(tmpprec, out precision))
                        throw DbUpgradeException.SchemaIsNotSupported();
                }

                for (; index < sql.Length; index++)
                {
                    if (sql[index] == ')')
                        break;
                }
                if (index == sql.Length)
                {
                    index = saved;
                    return -1;
                }
                index++;
                if (!int.TryParse(tmpsize, out size))
                    throw DbUpgradeException.SchemaIsNotSupported();
                return size;
            } // if
            else
            {
                index = saved;
                return -1;
            }
        }

        /// <summary>
        /// Parse a potentially delimited token. Update the index parameter according
        /// to the found token.
        /// </summary>
        /// <param name="sql">The SQL string to parse</param>
        /// <param name="index">The index into the SQL string where search starts.</param>
        /// <returns>The found token.</returns>
        private static string ParsePotentiallyDelimitedToken(string sql, ref int index)
        {
            int saved = index;

            // Eat any trailing whitespace
            for (; index < sql.Length; index++)
            {
                if (!Char.IsWhiteSpace(sql[index]))
                    break;
            } // for

            // If the current character is a [ or " or ' - store that character
            // and try to match against the next one.
            bool hasDelim = false;
            char delim = ' ';
            if (sql[index] == '"' || sql[index] == '\'' || sql[index] == '[' || sql[index] == '`')
            {
                hasDelim = true;
                if (sql[index] == '[')
                    delim = ']';
                else
                    delim = sql[index];
            }

            if (hasDelim)
            {
                index++; // Skip the start delimiter character
                int startIndex = index;
                for (; index < sql.Length; index++)
                {
                    if (sql[index] == delim)
                        break;
                }
                if (index >= sql.Length)
                {
                    index = saved;
                    return null;
                }

                string res = sql.Substring(startIndex, index - startIndex);
                index++; // Skip the end delimiter
                return res;
            }
            else
            {
                int startIndex = index;
                for (; index < sql.Length; index++)
                {
                    if (Char.IsWhiteSpace(sql[index]) ||
                        !(Char.IsLetterOrDigit(sql[index]) || sql[index] == '_'))
                        break;
                }
                if (index >= sql.Length)
                {
                    index = saved;
                    return null;
                }
                string res = sql.Substring(startIndex, index - startIndex);
                return res;
            } // else
        }

        /// <summary>
        /// Find the right parentheses that closes a column statement
        /// </summary>
        /// <param name="sql">The string to search</param>
        /// <returns>THe index of the closing right parentheses or -1 if not found.</returns>
        private static int FindClosingRightParens(string sql)
        {
            int pcount = 0;
            for (int i = 0; i < sql.Length; i++)
            {
                if (sql[i] == '(')
                    pcount++;
                else if (sql[i] == ')')
                {
                    pcount--;
                    if (pcount == -1)
                        return i;
                } // else
            } // for

            return -1;
        }

        /// <summary>
        /// Translate from the type name to the corresponding DB type.
        /// </summary>
        /// <param name="type">The name of the type</param>
        /// <returns>The DbType for that type.</returns>
        private static DbType GetDbType(string type)
        {
            type = type.ToLower();
            if (type == "int")
                return DbType.Int32;
            if (type == "tinyint")
                return DbType.Byte;
            if (type == "smallint")
                return DbType.Int16;
            if (type == "integer")
                return DbType.Int64;
            if (type == "nvarchar")
                return DbType.String;
            if (type == "varchar")
                return DbType.AnsiString;
            if (type == "char")
                return DbType.AnsiStringFixedLength;
            if (type == "text")
                return DbType.String;
            if (type == "boolean")
                return DbType.Boolean;
            if (type == "bit")
                return DbType.Boolean;
            if (type == "timestamp")
                return DbType.DateTime;
            if (type == "float")
                return DbType.Double;
            if (type == "real")
                return DbType.Single;
            if (type == "blob")
                return DbType.Binary;
            return DbType.String;
            //throw DbUpgradeException.SchemaIsNotSupported();
        }

        /// <summary>
        /// Parse the header part of a table DDL (CREATE TABLE 'name')
        /// </summary>
        /// <param name="sql">The SQL string to parse</param>
        /// <returns>The name of the parsed table.</returns>
        public static string ParseTableName(ref string sql)
        {
            Match m = _tableHeader.Match(sql);

            if (m.Success)
            {
                int index = m.Index + m.Length;
                string tableName = ParsePotentiallyDelimitedToken(sql, ref index);
                if (tableName != null)
                {
                    sql = sql.Substring(index);
                    return tableName;
                }
            }

            throw DbUpgradeException.SchemaIsNotSupported();
        }

        /// <summary>
        /// Scans the specified string in search of the specified token.
        /// </summary>
        /// <param name="str">The string to scan</param>
        /// <param name="token">The token to search for</param>
        /// <returns>TRUE if the token was found (str is adjusted to point right after
        /// the token in the input string) or FALSE if not found (str is 
        /// left unchanged).</returns>
        private static bool ScanToken(ref string str, string token)
        {
            int mindex = 0;
            int index = 0;
            for (index = 0; index < str.Length && mindex < token.Length; index++)
            {
                if (str[index] == token[mindex])
                    mindex++;
                else if (Char.IsWhiteSpace(str, index))
                    mindex = 0;
                else
                    return false;
            } // for
            if (mindex == token.Length)
            {
                if (str.Length <= index)
                    str = string.Empty;
                else
                    str = str.Substring(index);
                return true;
            }
            else
                return false;
        }
        #endregion

        #region Private Variables
        private static Regex _sizeRx = new Regex(@"\s*\(\s*(\d+)\s*\)");
        private static Regex _tableHeader = new Regex(@"CREATE\s+TABLE\s+");
        private static Regex _tableExistedHeader = new Regex(@"CREATE\s+TABLE\s+IFs+NOTs+EXISTS");
        private static Regex _columnConstraints =
            new Regex(@"(PRIMARY KEY)|(NOT NULL|NULL)|(DEFAULT\s+((CURRENT_TIMESTAMP|CURRENT_DATE|CURRENT_TIME)|(\'([^\']*)\')|(\(?(\-?\d+\.\d+)\)?)|(\(?(\-?\d+)\)?)))|(COLLATE\s+([a-zA-Z_][a-zA-Z0-9_]*))|(UNIQUE)|(AUTOINCREMENT)");
        private static Regex _indexHeader =
            new Regex(@"CREATE\s+(UNIQUE\s+)?INDEX\s+");
        private Dictionary<string, DbIndex> _indexes = new Dictionary<string, DbIndex>();
        private Dictionary<string, DbTable> _tables = new Dictionary<string, DbTable>();
        #endregion
    }

}
