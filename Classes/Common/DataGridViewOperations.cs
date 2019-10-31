using System.Collections;
using System.Data;
using System.Windows.Forms;

namespace ASTA.Classes
{
    public class DataGridViewOperations
    {
        public DataGridViewOperations() { }

        public int ColumnsCount(DataGridView dgv) //add string into  from other threads
        {
            int iDgv = dgv?.ColumnCount ?? -1;
            return iDgv;
        }

        public int RowsCount(DataGridView dgv) //add string into  from other threads
        {
            int iDgv = dgv?.Rows?.Count ?? -1;
            return iDgv;
        }

        public string ColumnName(DataGridView dgv, int indexColumn) //add string into  from other threads
        {
            string sDgv = null;
            if (indexColumn < 0)
            {
                return sDgv;
            }

            sDgv = dgv?.Columns[indexColumn]?.HeaderText;
            return sDgv;
        }

        public string CellValue(DataGridView dgv, int indexRow, int indexCells) //from other threads
        {
            string sDgv = null;
            if (indexRow < 0 || indexCells < 0)
            {
                return sDgv;
            }

            sDgv = dgv?.Rows[indexRow]?.Cells[indexCells]?.Value?.ToString();
            return sDgv;
        }

        public string CurrentCellValue(DataGridView dgv) //from other threads
        {
            int currentRowIndex = dgv?.CurrentRow?.Index ?? -1;
            int currentColumnIndex = dgv?.CurrentCell?.ColumnIndex ?? -1;

            string sDgv = null;

            if (currentRowIndex < 0 || currentColumnIndex < 0)
            {
                return sDgv;
            }

            sDgv = dgv?.Rows[currentRowIndex]?.Cells[currentColumnIndex]?.Value?.ToString()?.Trim() ?? null;

            return sDgv;
        }

        public int CurrentRowIndex(DataGridView dgv) //add string into  from other threads
        {
            int iDgv = dgv?.CurrentRow?.Index ?? -1;
            return iDgv;
        }

        public int CurrentColumnIndex(DataGridView dgv) //add string into  from other threads
        {
            int iDgv = dgv?.CurrentCell?.ColumnIndex ?? -1;
            return iDgv;
        }

        public void ShowData(DataGridView dgv, DataTable dt)
        {
            dgv.Visible = false;
            // clear datasource
            ArrayList Empty = new ArrayList();
            dgv.DataSource = Empty;

            if (dt != null && dt?.Columns?.Count > 0 && dt?.Rows?.Count > 0)
            {
                dgv.DataSource = dt;
            }

            dgv.Visible = true;
            dgv.Refresh();
        }

        public void ShowData(DataGridView dgv, object obj)
        {
            dgv.Visible = false;

            // clear datasource
            ArrayList Empty = new ArrayList();
            dgv.DataSource = Empty;

            if (obj != null)
            {
                dgv.DataSource = obj;
            }

            dgv.Visible = true;
            dgv.Refresh();
        }

        public string[] cellValue = new string[10];
        private bool correctData { get; set; }

        public void FindValuesInCurrentRow(DataGridView dgv, params string[] columnsName)
        {
            int IndexCurrentRow = -1;
            try { IndexCurrentRow = dgv.CurrentRow.Index; } catch { }

            correctData = (-1 < IndexCurrentRow & 0 < columnsName.Length & columnsName.Length < 11) ? true : false;

            if (correctData)
            {
                for (int i = 0; i < dgv.ColumnCount; i++)
                {
                    if (dgv.Columns[i].HeaderText == columnsName[0])
                    {
                        cellValue[0] = dgv.Rows[IndexCurrentRow]?.Cells[i]?.Value?.ToString();
                    }
                    else if (columnsName.Length > 1 && dgv.Columns[i].HeaderText == columnsName[1])
                    {
                        cellValue[1] = dgv.Rows[IndexCurrentRow]?.Cells[i]?.Value?.ToString();
                    }
                    else if (columnsName.Length > 2 && dgv.Columns[i].HeaderText == columnsName[2])
                    {
                        cellValue[2] = dgv.Rows[IndexCurrentRow]?.Cells[i]?.Value?.ToString();
                    }
                    else if (columnsName.Length > 3 && dgv.Columns[i].HeaderText == columnsName[3])
                    {
                        cellValue[3] = dgv.Rows[IndexCurrentRow]?.Cells[i]?.Value?.ToString();
                    }
                    else if (columnsName.Length > 4 && dgv.Columns[i].HeaderText == columnsName[4])
                    {
                        cellValue[4] = dgv.Rows[IndexCurrentRow]?.Cells[i]?.Value?.ToString();
                    }
                    else if (columnsName.Length > 5 && dgv.Columns[i].HeaderText == columnsName[5])
                    {
                        cellValue[5] = dgv.Rows[IndexCurrentRow]?.Cells[i]?.Value?.ToString();
                    }
                    else if (columnsName.Length > 6 && dgv.Columns[i].HeaderText == columnsName[6])
                    {
                        cellValue[6] = dgv.Rows[IndexCurrentRow]?.Cells[i]?.Value?.ToString();
                    }
                    else if (columnsName.Length > 7 && dgv.Columns[i].HeaderText == columnsName[7])
                    {
                        cellValue[7] = dgv.Rows[IndexCurrentRow]?.Cells[i]?.Value?.ToString();
                    }
                    else if (columnsName.Length > 8 && dgv.Columns[i].HeaderText == columnsName[8])
                    {
                        cellValue[8] = dgv.Rows[IndexCurrentRow]?.Cells[i]?.Value?.ToString();
                    }
                    else if (columnsName.Length == 10 && dgv.Columns[i].HeaderText == columnsName[9])
                    {
                        cellValue[9] = dgv.Rows[IndexCurrentRow]?.Cells[i]?.Value?.ToString();
                    }
                }
            }
            else
            {
                return;
            }
        }
    }
}
