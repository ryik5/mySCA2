﻿using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ASTA.Classes
{
    internal static class DataTableExtensions
    {
        public static void SetColumnsOrder(this DataTable table, params string[] columnNames)
        {
            List<string> listColNames = columnNames.ToList();

            //Remove invalid column names.
            foreach (string colName in columnNames)
            {
                if (!table.Columns.Contains(colName))
                {
                    listColNames.Remove(colName);
                }
            }

            int columnIndex = 0;
            foreach (var columnName in listColNames)
            {
                table.Columns[columnName].SetOrdinal(columnIndex);
                columnIndex++;
            }
        }
    }
}