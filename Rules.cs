using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace mySCA2
{
    class Rules
    {



        /*
         dataView.RowFilter = "id = 10";      // no special character in column name "id"
         dataView.RowFilter = "$id = 10";     // no special character in column name "$id"
         dataView.RowFilter = "[#id] = 10";   // special character "#" in column name "#id"
         dataView.RowFilter = "[[id\]] = 10"; // special characters in column name "[id]"
         dataView.RowFilter = "Name = 'John'"        // string value
         dataView.RowFilter = "Name = 'John ''A'''"  // string with single quotes "John 'A'"
         dataView.RowFilter = String.Format("Name = '{0}'", "John 'A'".Replace("'", "''"));
         dataView.RowFilter = "Year = 2008"          // integer value
         dataView.RowFilter = "Price = 1199.9"       // float value
         dataView.RowFilter = String.Format(CultureInfo.InvariantCulture.NumberFormat, "Price = {0}", 1199.9f);
         dataView.RowFilter = "Date = #12/31/2008#"          // date value (time is 00:00:00)
         dataView.RowFilter = "Date = #2008-12-31#"          // also this format is supported
         dataView.RowFilter = "Date = #12/31/2008 16:44:58#" // date and time value
         dataView.RowFilter = String.Format(CultureInfo.InvariantCulture.DateTimeFormat, "Date = #{0}#", new DateTime(2008, 12, 31, 16, 44, 58));
         dataView.RowFilter = "Date = '31.12.2008 16:44:58'" // if current culture is German
         dataView.RowFilter = "Price = '1199.90'"            // if current culture is English
         dataView.RowFilter = "Num = 10"             // number is equal to 10
         dataView.RowFilter = "Date < #1/1/2008#"    // date is less than 1/1/2008
         dataView.RowFilter = "Name <> 'John'"       // string is not equal to 'John'
         dataView.RowFilter = "Name >= 'Jo'"         // string comparison
         dataView.RowFilter = "Id IN (1, 2, 3)"                    // integer values
         dataView.RowFilter = "Price IN (1.0, 9.9, 11.5)"          // float values
         dataView.RowFilter = "Name IN ('John', 'Jim', 'Tom')"     // string values
         dataView.RowFilter = "Date IN (#12/31/2008#, #1/1/2009#)" // date time values
         dataView.RowFilter = "Id NOT IN (1, 2, 3)"  // values not from the list
         dataView.RowFilter = "Name LIKE 'j*'"       // values that start with 'j'
         dataView.RowFilter = "Name LIKE '%jo%'"     // values that contain 'jo'
         dataView.RowFilter = "Name NOT LIKE 'j*'"   // values that don't start with 'j'     
         dataView.RowFilter = "City = 'Tokyo' AND (Age < 20 OR Age > 60)";    // operator AND has precedence over OR operator, parenthesis are needed
         dataView.RowFilter = "City <> 'Tokyo' AND City <> 'Paris'";  // following examples do the same
         dataView.RowFilter = "NOT City = 'Tokyo' AND NOT City = 'Paris'";
         dataView.RowFilter = "NOT (City = 'Tokyo' OR City = 'Paris')";
         dataView.RowFilter = "City NOT IN ('Tokyo', 'Paris')";
         dataView.RowFilter = "MotherAge - Age < 20";   // people with young mother
         dataView.RowFilter = "Age % 10 = 0";           // people with decennial birthday
         dataView.RowFilter = "Salary > AVG(Salary)";  // select people with above-average salary
         dataView.RowFilter = "COUNT(Child.IdOrder) > 5";    // select orders which have more than 5 items
         dataView.RowFilter = "SUM(Child.Price) >= 500";  // select orders which total price (sum of items prices) is greater or equal $500
          */

        /*
        //datatable
        var table = new DataTable();

         //get some data
         using (var conn = new SqlConnection(yourSqlConn))
           {
             var comm = new SqlCommand(@"select * from someTable order by someColumn", conn);
             comm.CommandType = CommandType.Text;
             conn.Open();
             var data = comm.ExecuteReader();
             table.Load(data);
            }

         //bind to some control (repeater)
          rptFirstList.DataSource = table;
          rptFirstList.DataBind();

        //first way
        foreach (DataRow dr in table.Select("someColumn = 'value'","someColumnToSort"))    //loop through rows and import based on filter
         {secondTable.ImportRow(dr); }

          //second way
          var secondTable = new DataTable();
          secondTable = table.Copy();
          secondTable.DefaultView.Sort = "someOtherColumn";

       DataTable copyDataTable;
       copyDataTable = table.Copy();
       copyDataTable.Columns.Remove("ColB");

        or

       int columnIndex = 1;//this will remove the second column
       DataTable copyDataTable;
       copyDataTable = table.Copy();
       copyDataTable.Columns.RemoveAt(columnIndex);


       System.Data.DataView view = new System.Data.DataView(yourOriginalTable);
       System.Data.DataTable selected = 
       view.ToTable("Selected", false, "col1", "col2", "col6", "col7", "col3");             

      dtPeopleTemp = dtPeople.Clone();
      string ss = "search_text";
      for (int j = 0; j < dtPeople.Rows.Count; j++)
      {
          if (ss == dtPeople.Rows[j]["Name_Collumn"].ToString())
          {
              dtPeopleTemp.ImportRow(dtPeople.Rows[j]);
          }
      }*/



        /*
 string dataSource = "Database.s3db";
using (SQLiteConnection connection = new SQLiteConnection())
{
connection.ConnectionString = "Data Source=" + dataSource;
    connection.Open();
using (SQLiteCommand command = new SQLiteCommand(connection))
{
  command.CommandText = "update Example set Info = :info, Text = :text where ID=:id";

    command.Parameters.Add("info", DbType.String).Value = textBox2.Text; 
  command.Parameters.Add("text", DbType.String).Value = textBox3.Text; 
  command.Parameters.Add("id", DbType.String).Value = textBox1.Text; 
  command.ExecuteNonQuery();
}
}
*/


        /*
        private void GetRowsByFilter()
        {
            DataTable table = dtPeople.Copy();
            // Presuming the DataTable has a column named Date.
            string expression;
            expression = "Date > #1/1/00#";
            DataRow[] foundRows;

            // Use the Select method to find all rows matching the filter.
            foundRows = table.Select(expression);

            // Print column 0 of each returned row.
            for (int i = 0; i < foundRows.Length; i++)
            {
                Console.WriteLine(foundRows[i][0]);
            }
        }
        */


        //Make default sort
        /* try
         {
             DataView dv = new DataView(dtPeople);
             dv.Sort = "[Фамилия Имя Отчество] ASC, [Дата регистрации] ASC, [Время регистрации] ASC";
             DataTable dtTemp = dv.ToTable();
             dtPeople.Dispose();
             dtPeople = dtTemp.Copy();
         }catch(Exception expt) { MessageBox.Show(expt.ToString()); }
         */
        //dt.DefaultView.Sort = "EmpID,EmpName Desc"


        /*
dtPersonTemp = new DataTable();
dtPersonTemp = dtPersonRegisteredFull.Copy();
dtPersonTemp.Columns[27].ColumnMapping = MappingType.Hidden;
// copyDataTable.Columns.RemoveAt(0); //Remove column with index 0
*/


        // добавить ссылку в проект на System.Data.DataSetExtensions
        /*var results = from myRow in dt.AsEnumerable()
         where myRow.Field<int>("RowNo") == 1
         select myRow;
        DataTable dataresult = results.CopyToDataTable();*/


        //var rows = dt.Select("col1 > 5");
        //foreach (var row in rows)
        //    {row.Delete();}
        //DataView dv = ft.DefaultView;
        //dv.Sort = "occr desc";
        //DataTable sortedDT = dv.ToTable();
        //
        //dataTable.DefaultView.Sort = "Col1, Col2, Col3"
        //
        //DataRow[] foundRows=table.Select("Date = '1/31/1979' or OrderID = 2", "CompanyName ASC");
        //DataTable dt = foundRows.CopyToDataTable();
        //
        //DataRow[] dataRows = table.Select().OrderBy(u => u["EmailId"]).ToArray();
        //
        //DataTable dt = new DataTable();         
        //dt.DefaultView.Sort = "Column_name desc";
        //dt = dt.DefaultView.ToTable();


        //DataView dv = ft.DefaultView;
        //dv.Sort = "occr desc";
        //DataTable sortedDT = dv.ToTable();
        //
        //dataTable.DefaultView.Sort = "Col1, Col2, Col3"
        //
        //DataRow[] foundRows=table.Select("Date = '1/31/1979' or OrderID = 2", "CompanyName ASC");
        //DataTable dt = foundRows.CopyToDataTable();
        //
        //DataRow[] dataRows = table.Select().OrderBy(u => u["EmailId"]).ToArray();
        //
        //DataTable dt = new DataTable();         
        //dt.DefaultView.Sort = "Column_name desc";
        //dt = dt.DefaultView.ToTable();

        /* //searh min and max value. works
int minTimeLevel = int.MaxValue;
int maxTimeLevel = int.MinValue;
foreach (DataRow dr in dataTableSource.Rows)
{
    int accountLevel = dr.Field<int>("Время регистрации");
    minTimeLevel = Math.Min(minTimeLevel, accountLevel);
    maxTimeLevel = Math.Max(maxTimeLevel, accountLevel);
}
//or
int minLevel = Convert.ToInt32(dataTableSource.Compute("min([Время регистрации])", string.Empty));
//or
List<int> levels = dataTableSource.AsEnumerable().Select(al => al.Field<int>("Время регистрации")).Distinct().ToList();
int min = levels.Min();
int max = levels.Max();*/

        /*
         
         
         */
    }
}
