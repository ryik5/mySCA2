using System.Collections;
using System.Data;
using System.Windows.Forms;

namespace ASTA.Classes
{
    public class DataGridViewOperations
    {
        public static int ColumnsCount(DataGridView dgv) //add string into  from other threads
        {
            return dgv?.ColumnCount ?? -1;
        }

        public static int RowsCount(DataGridView dgv) //add string into  from other threads
        {
            return dgv?.Rows?.Count ?? -1;
        }

        public static string ColumnName(DataGridView dgv, int indexColumn) //add string into  from other threads
        {
            return indexColumn < 0 ? null : dgv?.Columns[indexColumn]?.HeaderText;
        }

        public string CurrentCellValue(DataGridView dgv) //from other threads
        {
            int currentRowIndex = dgv?.CurrentRow?.Index ?? -1;
            int currentColumnIndex = dgv?.CurrentCell?.ColumnIndex ?? -1;

            return currentRowIndex < 0 || currentColumnIndex < 0
                ? null
                : dgv?.Rows[currentRowIndex]?.Cells[currentColumnIndex]?.Value?.ToString()?.Trim() ?? null;
        }

        public static int CurrentRowIndex(DataGridView dgv) //add string into  from other threads
        {
            return dgv?.CurrentRow?.Index ?? -1;
        }

        public static int CurrentColumnIndex(DataGridView dgv) //add string into  from other threads
        {
            return dgv?.CurrentCell?.ColumnIndex ?? -1;
        }

        public void AddDataTable(DataGridView dgv, DataTable dt)
        {
            if (dt != null && dt?.Columns?.Count > 0 && dt?.Rows?.Count > 0)
            {
                dgv.DataSource = dt;
            }
            else
            {
                dgv.DataSource = new ArrayList();
            }
        }

        public void Paint(DataGridView dgv, string columnName, string desiredData)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row?.Cells[columnName]?.Value?.ToString() == desiredData)
                    row.DefaultCellStyle.BackColor = System.Drawing.Color.Red;
                else
                    row.DefaultCellStyle.BackColor = System.Drawing.Color.White;
            }
        }


        public void Show(DataGridView dgv, bool visible)
        {
            dgv.Visible = visible;

            if (visible)
            { dgv.Refresh(); }
        }

        //   public string[] cellValue = new string[10];
        private bool CheckCorrectData { get; set; }

        public string[] FindValuesInCurrentRow(DataGridView dgv, params string[] columnsName)
        {
            string[] cellValue = null;
            int IndexCurrentRow = CurrentRowIndex(dgv);

            CheckCorrectData = (-1 < IndexCurrentRow & 0 < columnsName.Length);

            if (CheckCorrectData)
            {
                cellValue = new string[columnsName.Length];
                for (int i = 0; i < dgv.ColumnCount; i++)
                {
                    for (int j = 0; j < columnsName.Length; j++)
                    {
                        if (columnsName.Length > j && dgv.Columns[i].HeaderText == columnsName[j])
                        {
                            cellValue[j] = dgv.Rows[IndexCurrentRow]?.Cells[i]?.Value?.ToString();
                            break;
                        }
                    }
                }
            }

            return cellValue;
        }
    }
}
