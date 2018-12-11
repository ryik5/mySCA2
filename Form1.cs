using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
//using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

// for Crypography
//using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace mySCA2
{
    public partial class FormPersonViewerSCA : Form
    {
        private System.Diagnostics.FileVersionInfo myFileVersionInfo;
        private string myRegKey = @"SOFTWARE\RYIK\SCA2";

        private string strVersion;
        private ContextMenu contextMenu1;
        private bool buttonAboutForm;

        private readonly System.IO.FileInfo databasePerson = new System.IO.FileInfo(@".\person.db");
        private string nameOfLastTableFromDB = "PersonRegistered";
        private bool bLoaded = false;
        private int iCounterLine = 0;
        private List<string> listFIO = new List<string>(); // List of FIO and identity of data
        private List<string> listPoints = new List<string>(); // List of all Points of SCA
        private List<string> listRegistrations = new List<string>(); // List whole of registration of the selected person at All servers
        int iFIO = 0;
        private List<string> lListFIOTemp = new List<string>();
        private decimal dControlHourSelected = 9;
        private decimal dControlMinuteSelected = 0;
        private bool bErrorData = false;

        //Visual of registration
        private PictureBox pictureBox1 = new PictureBox();
        private Bitmap bmp = new Bitmap(1, 1);

        private List<string> selectedDates = new List<string>();
        private string[] myBoldedDates;
        private List<string> boldeddDates = new List<string>();
        private string[] workSelectedDays;

        //Page of "Settings of Programm"
        private Label labelServer1;
        private TextBox textBoxServer1;
        private Label labelServer1UserName;
        private TextBox textBoxServer1UserName;
        private Label labelServer1UserPassword;
        private TextBox textBoxServer1UserPassword;

        private Color clrRealRegistration = Color.PaleGreen;
        private string sLastSelectedElement = "MainForm";

        //Settings of Programm
        private bool bServer1Exist = false;
        private string sServer1 = "";
        private string sServer1Registry = "";
        private string sServer1UserName = "";
        private string sServer1UserNameRegistry = "";
        private string sServer1UserPassword = "";
        private string sServer1UserPasswordRegistry = "";
        private readonly byte[] btsMess1 = Convert.FromBase64String(@"OCvesunvXXsxtt381jr7vp3+UCwDbE4ebdiL1uinGi0="); //Key Encrypt
        private readonly byte[] btsMess2 = Convert.FromBase64String(@"NO6GC6Zjl934Eh8MAJWuKQ=="); //Key Decrypt

        private OpenFileDialog openFileDialog1 = new OpenFileDialog();
        private List<string> listGroups = new List<string>();

        private DataTable dtPeople = new DataTable("People");
        private DataColumn[] dcPeople ={
                                  new DataColumn("№ п/п",typeof(int)),//0
                                  new DataColumn("Фамилия Имя Отчество",typeof(string)),//1
                                  new DataColumn("NAV-код",typeof(string)),
                                  new DataColumn("Группа",typeof(string)),
                                  new DataColumn("Время прихода,часы",typeof(string)),
                                  new DataColumn("Время прихода,минут",typeof(string)), //5
                                  new DataColumn("Время прихода",typeof(decimal)),
                                  new DataColumn("Время ухода,часы",typeof(string)),
                                  new DataColumn("Время ухода,минут",typeof(string)),
                                  new DataColumn("Время ухода",typeof(decimal)),
                                  new DataColumn("№ пропуска",typeof(int)), //10
                                  new DataColumn("Отдел",typeof(string)),
                                  new DataColumn("Дата регистрации",typeof(string)),
                                  new DataColumn("Время регистрации,часы",typeof(string)),
                                  new DataColumn("Время регистрации,минут",typeof(string)),
                                  new DataColumn("Время регистрации",typeof(decimal)), //15
                                  new DataColumn("Реальное время ухода,часы",typeof(string)),
                                  new DataColumn("Реальное время ухода,минут",typeof(string)),
                                  new DataColumn("Реальное время ухода",typeof(decimal)), //18
                                  new DataColumn("Сервер СКД",typeof(string)), //19
                                  new DataColumn("Имя точки прохода",typeof(string)), //20
                                  new DataColumn("Направление прохода",typeof(string)), //21
                                  new DataColumn("Время прихода ЧЧ:ММ",typeof(string)),
                                  new DataColumn("Время ухода ЧЧ:ММ",typeof(string)),
                                  new DataColumn("Реальное время прихода ЧЧ:ММ",typeof(string)),
                                  new DataColumn("Реальное время ухода ЧЧ:ММ",typeof(string)), //25
                                  new DataColumn("Реальное отработанное время",typeof(decimal)), //26
                                  new DataColumn("Реальное отработанное время ЧЧ:ММ",typeof(string)), //27
    };

        private DataTable dtPersonTemp = new DataTable("PersonTemp");
        private DataTable dtPersonRegisteredFull = new DataTable("PersonRegisteredFull");
        private DataTable dtPersonRegistered = new DataTable("PersonRegistered");
        private DataTable dtPersonGroup = new DataTable("PersonGroup");
        private DataTable dtPersonsLastList = new DataTable("PersonsLastList");
        private DataTable PersonsLastComboList = new DataTable("PersonsLastComboList");


        //Color Control elements of Person depending of the selected MenuItem  
        private Color labelGroupCurrentBackColor;
        private Color textBoxGroupCurrentBackColor;
        private Color labelGroupDescriptionCurrentBackColor;
        private Color textBoxGroupDescriptionCurrentBackColor;
        private Color comboBoxFioCurrentBackColor;
        private Color textBoxFIOCurrentBackColor;
        private Color textBoxNavCurrentBackColor;



        public FormPersonViewerSCA()
        { InitializeComponent(); }

        private void Form1_Load(object sender, EventArgs e)
        { Form1Load(); }

        private void Form1Load()
        {
            CheckSavedDataInRegistry();

            myFileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath);
            strVersion = myFileVersionInfo.Comments + " ver." + myFileVersionInfo.FileVersion + " " + myFileVersionInfo.LegalCopyright;
            StatusLabel1.Text = myFileVersionInfo.ProductName + " ver." + myFileVersionInfo.FileVersion + " " + myFileVersionInfo.LegalCopyright;
            StatusLabel1.Alignment = ToolStripItemAlignment.Right;

            contextMenu1 = new ContextMenu();  //Context Menu on notify Icon
            notifyIcon1.ContextMenu = contextMenu1;
            contextMenu1.MenuItems.Add("About", AboutSoft);
            contextMenu1.MenuItems.Add("Exit", ApplicationExit);
            notifyIcon1.Text = myFileVersionInfo.ProductName + "\nv." + myFileVersionInfo.FileVersion + "\n" + myFileVersionInfo.CompanyName;
            this.Text = myFileVersionInfo.Comments;

            groupBoxProperties.Visible = false;

            GetFioItem.BackColor = Color.LightSkyBlue;
            HelpAboutItem.BackColor = Color.PaleGreen;
            ExitItem.BackColor = Color.DarkOrange;
            DeleteGroupItem.BackColor = Color.DarkOrange;

            //Set up Starting values
            dateTimePickerStart.CustomFormat = "yyyy-MM-dd";
            dateTimePickerEnd.CustomFormat = "yyyy-MM-dd";
            dateTimePickerStart.Format = DateTimePickerFormat.Custom;
            dateTimePickerEnd.Format = DateTimePickerFormat.Custom;
            dateTimePickerStart.MinDate = DateTime.Parse("2016-01-01");
            dateTimePickerEnd.MinDate = DateTime.Parse("2016-01-01");
            dateTimePickerStart.MaxDate = DateTime.Now;
            dateTimePickerEnd.MaxDate = DateTime.Parse("2022-12-01");
            dateTimePickerStart.Value = DateTime.Parse(DateTime.Now.Year + "-" + DateTime.Now.Month + "-01");
            dateTimePickerEnd.Value = DateTime.Now;
            numUpDownHour.Value = 9;
            numUpDownMinute.Value = 0;
            PersonOrGroupItem.Text = "Работа с одной персоной";
            toolTip1.SetToolTip(textBoxGroup, "Создать группу");
            toolTip1.SetToolTip(textBoxGroupDescription, "Изменить описание группы");
            StatusLabel2.Text = "";

            TryMakeDB();
            UpdateTableOfDB();

            SetTechInfoIntoDB();
            BoldAnualDates();
            WriteLastParametersIntoVariable();
            CheckForIllegalCrossThreadCalls = false;
            AddAnualDateItem.Enabled = false;
            DeleteAnualDateItem.Enabled = false;
            EnterEditAnualItem.Enabled = true;

            MembersGroupItem.Enabled = false;
            AddPersonToGroupItem.Enabled = false;
            CreateGroupItem.Enabled = false;
            DeleteGroupItem.Visible = false;
            DeletePersonFromGroupItem.Visible = false;
            CheckBoxesFiltersAll_Enable(false);
            //UpdateControllingItem.Visible = false;
            ReportsItem.Visible = false;
            VisualItem.Visible = true;
            VisualItem.Enabled = false;
            VisualWorkedTimeItem.Visible = false;
            ExportIntoExcelItem.Enabled = false;

            int IndexCurrentRow = 2;
            int IndexColumn1 = 0;           // индекс 1-й колонки в датагрид
            int IndexColumn2 = 1;           // индекс 2-й колонки в датагрид

            string sFIO = "";
            try
            {
                textBoxFIO.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString();
                sFIO = textBoxFIO.Text;
            }
            catch { }
            try { textBoxNav.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn2].Value.ToString(); } catch { }

            numUpDownHour.Value = 9;
            numUpDownMinute.Value = 0;

            try
            {
                comboBoxFio.SelectedIndex = comboBoxFio.FindString(sFIO); //ищем в комбобокс-e выбранный ФИО и устанавливаем на него фокус
                if (comboBoxFio.FindString(sFIO) != -1 && ShortFIO(sFIO).Length > 3)
                {
                    StatusLabel2.Text = @"Выбран: " + ShortFIO(sFIO) + @" |  Всего ФИО: " + iFIO;
                }
                else if (ShortFIO(sFIO).Length < 3 && iFIO > 0)
                { StatusLabel2.Text = @"Всего ФИО: " + iFIO; }
            }
            catch { StatusLabel2.Text = " Начните работу с кнопки - \"Получить ФИО\""; }

            dtPeople.Columns.AddRange(dcPeople);

            dtPersonTemp = dtPeople.Clone();  //Copy only structure(Name of collumns)
            dtPersonRegisteredFull = dtPeople.Clone();  //Copy only structure(Name of collumns)
            dtPersonRegistered = dtPeople.Clone();  //Copy only structure(Name of collumns)
            dtPersonGroup = dtPeople.Clone();  //Copy only structure(Name of collumns)
            dtPersonsLastList = dtPeople.Clone();  //Copy only structure(Name of collumns)
            PersonsLastComboList = dtPeople.Clone();  //Copy only structure(Name of collumns)

            dataGridView1.ShowCellToolTips = true;
        }


        //todo converting string to decimal
        private decimal TimeStringToDecimal(string s)
        {
            decimal timeConverted = 0;
            if (s.Contains(':'))
            {
                string[] time = s.Split(':');
                timeConverted = TryParseStringToDecimal(time[0]) + (TryParseStringToDecimal(time[1]) + 1) / 60 - (1 / 60);
            }
            return timeConverted;
        }

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

        private void AboutSoft(object sender, EventArgs e) //Кнопка "О программе"
        { AboutSoft(); }

        private void AboutSoft()
        {
            AboutBox1 aboutBox = new AboutBox1();
            aboutBox.Show();
            buttonAboutForm = aboutBox.OKButtonClicked;
        }

        private void ApplicationExit(object sender, EventArgs e)
        { ApplicationExit(); }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        { ApplicationExit(); }

        private void ApplicationExit()
        { Application.Exit(); }

        private void TryMakeDB()
        {
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'PersonRegisteredFull' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, FIO TEXT, NAV TEXT, iDCard TEXT, DateRegistered TEXT, " +
                    "HourComming TEXT, MinuteComming TEXT, Comming REAL, HourControlling TEXT, MinuteControlling TEXT, Controlling REAL," +
                    " HourControllingOut TEXT, MinuteControllingOut TEXT, ControllingOut REAL,  WorkedOut REAL, ServerOfRegistration TEXT, Reserv1 TEXT, Reserv2 TEXT);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'PersonRegistered' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, FIO TEXT, NAV TEXT, iDCard TEXT, DateRegistered TEXT, " +
                    "HourComming TEXT, MinuteComming TEXT, Comming REAL, HourControlling TEXT, MinuteControlling TEXT, Controlling REAL, " +
                    "HourControllingOut TEXT, MinuteControllingOut TEXT, ControllingOut REAL,  WorkedOut REAL, ServerOfRegistration TEXT, Reserv1 TEXT, Reserv2 TEXT);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'PersonTemp' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, FIO TEXT, NAV TEXT, iDCard TEXT, DateRegistered TEXT, " +
                    "HourComming TEXT, MinuteComming TEXT, Comming REAL, HourControlling TEXT, MinuteControlling TEXT, Controlling REAL, HourControllingOut TEXT, " +
                    "MinuteControllingOut TEXT, ControllingOut REAL,  WorkedOut REAL, ServerOfRegistration TEXT, Reserv1 TEXT, Reserv2 TEXT);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'PersonGroup' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, FIO TEXT, NAV TEXT, GroupPerson TEXT, " +
                    "HourControlling TEXT, MinuteControlling TEXT, Controlling REAL, HourControllingOut TEXT, MinuteControllingOut TEXT, ControllingOut REAL, " +
                    "Reserv1 TEXT, Reserv2 TEXT, UNIQUE ('FIO', 'NAV', 'GroupPerson') ON CONFLICT REPLACE);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'PersonGroupDesciption' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, GroupPerson TEXT, GroupPersonDescription TEXT, Reserv1 TEXT, Reserv2 TEXT, " +
                    "UNIQUE ('GroupPerson') ON CONFLICT REPLACE);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'TechnicalInfo' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, PCName TEXT, POName TEXT, POVersion TEXT, LastDateStarted TEXT, Reserv1 TEXT, " +
                    "Reserv2 TEXT, Reserverd3 TEXT);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'BoldedDates' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, BoldedDate TEXT, NAV TEXT, Groups TEXT, Reserv1 TEXT, Reserv2 TEXT);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'MySettings' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, MyParameterName TEXT, MyParameterValue TEXT, Reserv1 TEXT, Reserv2 TEXT);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'ProgramSettings' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, PoParameterName TEXT, PoParameterValue TEXT, Reserv1 TEXT, Reserv2 TEXT, " +
                   "UNIQUE (PoParameterName) ON CONFLICT REPLACE);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'EquipmentSettings' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, EquipmentParameterName TEXT, EquipmentParameterValue TEXT, EquipmentParameterServer TEXT," +
                    "Reserv1, Reserv2, UNIQUE ('EquipmentParameterName', 'EquipmentParameterServer') ON CONFLICT REPLACE);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'PersonsLastList' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, PersonsList TEXT, " +
                    "Reserv1 TEXT, Reserv2 TEXT, UNIQUE ('PersonsList', Reserv1) ON CONFLICT REPLACE);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'PersonsLastComboList' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, ComboList TEXT, " +
                    "Reserv1 TEXT, Reserv2 TEXT, UNIQUE ('ComboList', Reserv1) ON CONFLICT REPLACE);", databasePerson);
        }

        private void UpdateTableOfDB()
        {
            TryUpdateStructureSqlDB("PersonRegisteredFull", "FIO TEXT, NAV TEXT, iDCard TEXT, DateRegistered TEXT, " +
                    "HourComming TEXT, MinuteComming TEXT, Comming REAL, HourControlling TEXT, MinuteControlling TEXT, Controlling REAL," +
                    " HourControllingOut TEXT, MinuteControllingOut TEXT, ControllingOut REAL,  WorkedOut REAL, ServerOfRegistration TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);

            TryUpdateStructureSqlDB("PersonRegistered", "FIO TEXT, NAV TEXT, iDCard TEXT, DateRegistered TEXT, " +
                    "HourComming TEXT, MinuteComming TEXT, Comming REAL, HourControlling TEXT, MinuteControlling TEXT, Controlling REAL, " +
                    "HourControllingOut TEXT, MinuteControllingOut TEXT, ControllingOut REAL,  WorkedOut REAL, ServerOfRegistration TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);


            TryUpdateStructureSqlDB("PersonTemp", "FIO TEXT, NAV TEXT, iDCard TEXT, DateRegistered TEXT, " +
                    "HourComming TEXT, MinuteComming TEXT, Comming REAL, HourControlling TEXT, MinuteControlling TEXT, Controlling REAL, HourControllingOut TEXT, " +
                    "MinuteControllingOut TEXT, ControllingOut REAL,  WorkedOut REAL, ServerOfRegistration TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);

            TryUpdateStructureSqlDB("PersonGroup", "FIO TEXT, NAV TEXT, GroupPerson TEXT, HourControlling TEXT, MinuteControlling TEXT, Controlling REAL, " +
                    "HourControllingOut TEXT, MinuteControllingOut TEXT, ControllingOut REAL, Reserv1 TEXT, Reserv2 TEXT", databasePerson);

            TryUpdateStructureSqlDB("PersonGroupDesciption", "GroupPerson TEXT, GroupPersonDescription TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
            TryUpdateStructureSqlDB("TechnicalInfo", "PCName TEXT, POName TEXT, POVersion TEXT, LastDateStarted TEXT, Reserv1 TEXT, Reserv2 TEXT, Reserverd3 TEXT", databasePerson);
            TryUpdateStructureSqlDB("BoldedDates", "BoldedDate TEXT, NAV TEXT, Groups TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
            TryUpdateStructureSqlDB("MySettings", "MyParameterName TEXT, MyParameterValue TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
            TryUpdateStructureSqlDB("ProgramSettings", " PoParameterName TEXT, PoParameterValue TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
            TryUpdateStructureSqlDB("EquipmentSettings", "EquipmentParameterName TEXT, EquipmentParameterValue TEXT, EquipmentParameterServer TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
            TryUpdateStructureSqlDB("PersonsLastList", "PersonsList TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
            TryUpdateStructureSqlDB("PersonsLastComboList", "ComboList TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
        }

        private void SetTechInfoIntoDB() //Write into DB Technical Info
        {
            string guid = System.Runtime.InteropServices.Marshal.GetTypeLibGuidForAssembly(System.Reflection.Assembly.GetExecutingAssembly()).ToString(); // получаем GIUD приложения
            string pcname = Environment.MachineName + "|" + Environment.OSVersion;
            string poname = myFileVersionInfo.FileName + "|" + myFileVersionInfo.ProductName;
            string poversion = myFileVersionInfo.FileVersion;
            string LastDateStarted = dateTimePickerEnd.Value.ToString();
            string Reserv1 = Environment.UserName;
            string Reserv2 = Environment.WorkingSet.ToString();

            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'TechnicalInfo' (PCName, POName, POVersion, LastDateStarted, Reserv1, Reserv2, Reserverd3) " +
                        " VALUES (@PCName, @POName, @POVersion, @LastDateStarted, @Reserv1, @Reserv2, @Reserverd3)", sqlConnection))
                    {
                        sqlCommand.Parameters.Add("@PCName", DbType.String).Value = pcname;
                        sqlCommand.Parameters.Add("@POName", DbType.String).Value = poname;
                        sqlCommand.Parameters.Add("@POVersion", DbType.String).Value = poversion;
                        sqlCommand.Parameters.Add("@LastDateStarted", DbType.String).Value = LastDateStarted;
                        sqlCommand.Parameters.Add("@Reserv1", DbType.String).Value = Reserv1;
                        sqlCommand.Parameters.Add("@Reserv2", DbType.String).Value = Reserv2;
                        sqlCommand.Parameters.Add("@Reserverd3", DbType.String).Value = guid;
                        try { sqlCommand.ExecuteNonQuery(); } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                    }
                }
            }
            myFileVersionInfo = null; guid = null; pcname = null; poname = null; poversion = null;
            LastDateStarted = null; Reserv1 = null; Reserv2 = null;
        }

        private void WriteLastParametersIntoVariable()   //Select Previous Data from DB and write it into the combobox and Parameters
        {
            int iCombo = 0;
            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();

                    using (var sqlCommand = new SQLiteCommand("SELECT PersonsList FROM PersonsLastList;", sqlConnection))
                    {
                        using (var reader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in reader)
                            {
                                try
                                {
                                    if (record != null && record.ToString().Length > 0)
                                    {
                                        listFIO.Add(record["PersonsList"].ToString().Trim());
                                    }
                                }
                                catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                            }
                        }
                    }

                    using (var sqlCommand = new SQLiteCommand("SELECT ComboList FROM PersonsLastComboList;", sqlConnection))
                    {
                        using (var reader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in reader)
                            {
                                try
                                {
                                    if (record != null && record.ToString().Length > 0)
                                    {
                                        _comboBoxFioAdd(record["ComboList"].ToString().Trim());
                                        iCombo++;
                                    }
                                }
                                catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                            }
                        }
                    }

                    using (var sqlCommand = new SQLiteCommand("SELECT PoParameterName, PoParameterValue  FROM ProgramSettings;", sqlConnection))
                    {
                        using (var reader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in reader)
                            {
                                try
                                {
                                    if (record != null && record.ToString().Length > 0)
                                    {
                                        if (record["PoParameterName"].ToString().Trim() == "clrRealRegistration")
                                            clrRealRegistration = Color.FromName(record["PoParameterValue"].ToString());
                                    }
                                }
                                catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                            }
                        }
                    }

                    using (var sqlCommand = new SQLiteCommand("SELECT EquipmentParameterName, EquipmentParameterValue, EquipmentParameterServer, Reserv1, Reserv2  FROM EquipmentSettings;", sqlConnection))
                    {
                        using (var reader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in reader)
                            {
                                try
                                {
                                    if (record != null && record.ToString().Length > 0)
                                    {
                                        if (record["EquipmentParameterValue"].ToString().Trim() == "Server1" && record["EquipmentParameterName"].ToString().Trim() == "Server1UserName")
                                        {
                                            sServer1 = sServer1Registry.Length == 0 ? record["EquipmentParameterServer"].ToString() : sServer1Registry;
                                            sServer1UserName = sServer1UserNameRegistry.Length == 0 ? record["Reserv1"].ToString() : sServer1UserNameRegistry;
                                            sServer1UserPassword = sServer1UserPasswordRegistry.Length == 0 ? record["Reserv2"].ToString() : sServer1UserPasswordRegistry;
                                        }
                                    }
                                }
                                catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                            }
                        }
                    }

                }
            }
            iFIO = iCombo;
            try { comboBoxFio.SelectedIndex = 0; } catch { }
        }

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
        private void ExecuteSql(string SqlQuery, System.IO.FileInfo FileDB) //Prepare DB and execute of SQL Query
        {
            if (!System.IO.File.Exists(databasePerson.FullName))
            { SQLiteConnection.CreateFile(databasePerson.FullName); }
            using (var connection = new SQLiteConnection($"Data Source={databasePerson.FullName};Version=3;"))
            {
                connection.Open();
                try
                {
                    using (var command = new SQLiteCommand(SqlQuery, connection))
                    { command.ExecuteNonQuery(); }
                }
                catch (Exception expt) { MessageBox.Show(expt.ToString()); }
            }
            SqlQuery = null;
        }

        private void TryUpdateStructureSqlDB(string tableName, string listCollumnsWithType, System.IO.FileInfo FileDB) //Update Table in DB and execute of SQL Query
        {
            using (var connection = new SQLiteConnection($"Data Source={databasePerson.FullName};Version=3;"))
            {
                connection.Open();
                foreach (string collumn in listCollumnsWithType.Split(','))
                {
                    using (var command = new SQLiteCommand("ALTER TABLE " + tableName + " ADD COLUMN " + collumn, connection))
                        try { command.ExecuteNonQuery(); } catch { }
                }
            }
        }

        //void ShowDataTableQuery(
        private void ShowDataTableQuery(System.IO.FileInfo databasePerson, string myTable, string mySqlQuery = "SELECT DISTINCT FIO AS 'Фамилия Имя Отчество', NAV AS 'NAV-код', " +
            " DateRegistered AS 'Дата регистрации', HourComming AS 'Время прихода, часы',  MinuteComming AS 'Время прихода, минуты', ServerOfRegistration AS 'Сервер', " +
            " HourControlling AS 'Контрольное время, часы', MinuteControlling AS 'Контрольное время, минуты', Reserv1 AS 'Точка прохода', Reserv2 AS 'Направление'",
            string mySqlWhere = "ORDER BY FIO, DateRegistered, Comming") //Query data from the Table of the DB
        {
            if (databasePerson.Exists)
            {
                // dtPeople.Clear();
                iCounterLine = 0;
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    using (var sqlDA = new SQLiteDataAdapter(mySqlQuery + " FROM '" + myTable + "' " + mySqlWhere + "; ", sqlConnection))
                    {
                        var dt = new DataTable();
                        //dtPeople 
                        sqlDA.Fill(dt);
                        _dataGridViewSource(dt);
                    }

                    using (var sqlCommand = new SQLiteCommand("Select * FROM '" + myTable + "' " + mySqlWhere + "; ", sqlConnection))
                    {
                        using (var sqlReader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in sqlReader)
                            {
                                try
                                {
                                    if (record?["id"] != null)
                                    { iCounterLine++; }
                                }
                                catch { }
                            }
                        }
                    }
                }
            }
            sLastSelectedElement = "dataGridView";
        }

        /// <summary>
        /// ///////////////////////////////////////////////////////////////////////////
        /// </summary>

        private void ShowDatatableOnDatagridview(DataTable dt, int[] collumnsHide, DataGridView dgv) //Query data from the Table of the DB
        {
            /*
            dtPersonTemp = new DataTable();
            dtPersonTemp = dtPersonRegisteredFull.Copy();
            dtPersonTemp.Columns[27].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[26].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[25].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[24].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[23].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[22].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[20].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[19].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[18].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[15].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[10].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[9].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[6].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[5].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[0].ColumnMapping = MappingType.Hidden;*/
            // copyDataTable.Columns.RemoveAt(0); //Remove column with index 0

            DataTable dataTable = new DataTable();
            dataTable = dt.Copy();

            for (int i = 0; i < collumnsHide.Length; i++)
            {
                dtPersonTemp.Columns[i].ColumnMapping = MappingType.Hidden;
            }

            _dataGridViewSource(dt);
            sLastSelectedElement = "dataGridView";
        }



        /// <summary>
        /// ///////////////////////////////////////////////////////////////////////////
        /// </summary>


        private void CopyWholeDataFromOneTableIntoAnother(System.IO.FileInfo databasePerson, string myTableInto, string myTableFrom) //Copy into Table from other Table
        {
            DeleteAllDataInTableQuery(databasePerson, myTableInto);

            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    using (var sqlCommand = new SQLiteCommand("INSERT INTO " + myTableInto + " Select * FROM " + myTableFrom, sqlConnection))
                    { try { sqlCommand.ExecuteNonQuery(); } catch { } }
                }
            }
            myTableInto = null; myTableFrom = null;
        }

        private void CopyDataOneNavUserIntoAnotherTable(System.IO.FileInfo databasePerson, string myTableInto, string myTableFrom, string myNAV) //Copy into Table from other Table
        {
            DeleteAllDataInTableQuery(databasePerson, myTableInto);

            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    using (var sqlCommand = new SQLiteCommand("INSERT INTO " + myTableInto + " Select * FROM " + myTableFrom + " where NAV like '" + myNAV + "'", sqlConnection))
                    { try { sqlCommand.ExecuteNonQuery(); } catch { } }
                }
            }
            myTableFrom = null; myTableInto = null; myNAV = null;
        }


        private void DeleteTable(System.IO.FileInfo databasePerson, string myTable) //Delete All data from the selected Table of the DB (both parameters are string)
        {
            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    using (var sqlCommand = new SQLiteCommand("DROP Table if exists '" + myTable + "';", sqlConnection))
                    { try { sqlCommand.ExecuteNonQuery(); } catch { } }
                    using (var sqlCommand = new SQLiteCommand("vacuum;", sqlConnection))   //vacuum DB
                    {
                        sqlCommand.ExecuteNonQuery();
                        sqlCommand.Dispose();
                    }
                    sqlConnection.Close();
                    sqlConnection.Dispose();
                }
                GC.Collect();
            }
            myTable = null;
        }

        private void DeleteAllDataInTableQuery(System.IO.FileInfo databasePerson, string myTable) //Delete All data from the selected Table of the DB (both parameters are string)
        {
            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    using (var sqlCommand = new SQLiteCommand("DELETE FROM '" + myTable + "';", sqlConnection))
                    { try { sqlCommand.ExecuteNonQuery(); } catch { } }

                    //vacuum DB
                    using (var sqlCommand = new SQLiteCommand("vacuum;", sqlConnection))
                    { sqlCommand.ExecuteNonQuery(); }
                }
            }
            myTable = null;
        }

        private void DeleteDataTableQueryNAV(System.IO.FileInfo databasePerson, string myTable,
            string mySqlParameter1 = "", string mySqlData1 = "",
            string mySqlParameter2 = "", string mySqlData2 = "") //Delete data from the Table of the DB by NAV (both parameters are string)
        {
            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    if (mySqlParameter1.Trim().Length > 0 && mySqlParameter2.Trim().Length == 0)
                    {
                        using (var sqlCommand = new SQLiteCommand("DELETE FROM '" + myTable + "' Where NAV like '" + _textBoxReturnText(textBoxNav) + "' AND " + mySqlParameter1 + "= @" + mySqlParameter1 + ";", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@" + mySqlParameter1, DbType.String).Value = mySqlData1;
                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }
                    }
                    if (mySqlParameter1.Trim().Length > 0 && mySqlParameter2.Trim().Length > 0)
                    {
                        using (var sqlCommand = new SQLiteCommand("DELETE FROM '" + myTable + "' Where NAV like '" + _textBoxReturnText(textBoxNav) + "' AND " + mySqlParameter1 + "= @" + mySqlParameter1 +
                            " AND " + mySqlParameter2 + "= @" + mySqlParameter2 + ";", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@" + mySqlParameter1, DbType.String).Value = mySqlData1;
                            sqlCommand.Parameters.Add("@" + mySqlParameter2, DbType.String).Value = mySqlData2;
                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }
                    }
                    //vacuum DB
                    using (var sqlCommand = new SQLiteCommand("vacuum;", sqlConnection))
                    { try { sqlCommand.ExecuteNonQuery(); } catch { } }
                }
            }
            myTable = null; mySqlParameter1 = null; mySqlData1 = null; mySqlParameter2 = null; mySqlData2 = null;
        }

        private void DeleteDataTableQuery(System.IO.FileInfo databasePerson, string myTable, string mySqlNav,
        string mySqlParameter1 = "", string mySqlData1 = "",
            string mySqlParameter2 = "", string mySqlData2 = "") //Delete data from the Table of the DB by NAV (both parameters are string)
        {
            if (databasePerson.Exists)
            {
                using (SQLiteConnection sqlConnection = new SQLiteConnection(string.Format("Data Source={0};Version=3;", databasePerson)))
                {
                    sqlConnection.Open();
                    if (mySqlParameter1.Trim().Length > 0 && mySqlParameter2.Trim().Length == 0)
                    {
                        using (var sqlCommand = new SQLiteCommand("DELETE FROM '" + myTable + "' " + mySqlNav + " AND " + mySqlParameter1 + "= @" + mySqlParameter1 + ";", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@" + mySqlParameter1, DbType.String).Value = mySqlData1;
                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }
                    }
                    if (mySqlParameter1.Trim().Length > 0 && mySqlParameter2.Trim().Length > 0)
                    {
                        using (SQLiteCommand sqlCommand = new SQLiteCommand(
                            $"DELETE FROM \'{myTable}\' {mySqlNav} AND {mySqlParameter1}= @{mySqlParameter1} AND {mySqlParameter2}= @{mySqlParameter2};", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@" + mySqlParameter1, DbType.String).Value = mySqlData1;
                            sqlCommand.Parameters.Add("@" + mySqlParameter2, DbType.String).Value = mySqlData2;
                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }
                    }

                    using (var sqlCommand = new SQLiteCommand("vacuum;", sqlConnection))  //vacuum DB
                    { try { sqlCommand.ExecuteNonQuery(); } catch { } }
                }
            }
            myTable = null; mySqlParameter1 = null; mySqlData1 = null; mySqlParameter2 = null; mySqlData2 = null; mySqlNav = null;
        }

        private void DeleteDataTableOneQuery(System.IO.FileInfo databasePerson, string myTable, string mySqlParameter1 = "", string mySqlData1 = "") //Delete data from the Table of the DB by NAV (both parameters are string)
        {
            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    if (mySqlParameter1.Length > 0)
                    {
                        using (var sqlCommand = new SQLiteCommand("DELETE FROM '" + myTable + "' Where " + mySqlParameter1 + "= @" + mySqlParameter1 + ";", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@" + mySqlParameter1, DbType.String).Value = mySqlData1;
                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }
                    }

                    using (var sqlCommand = new SQLiteCommand("vacuum;", sqlConnection))   //vacuum DB
                    { try { sqlCommand.ExecuteNonQuery(); } catch { } }
                }
            }
            myTable = null; mySqlParameter1 = null; mySqlData1 = null;
        }

        private void DeleteDataTableQueryLess(System.IO.FileInfo databasePerson, string myTable,
            string mySqlParameter2 = "", decimal mySqlData2 = 9) //Delete data from the Table of the DB (both parameters are string)
        {
            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    if (mySqlParameter2.Trim().Length > 0)
                    {
                        using (var sqlCommand = new SQLiteCommand("DELETE FROM '" + myTable + "' Where NAV like '" + _textBoxReturnText(textBoxNav) + "' AND " + mySqlParameter2 + " < @" + mySqlParameter2 + ";", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@" + mySqlParameter2, DbType.Decimal).Value = mySqlData2;
                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }
                    }
                    //vacuum DB
                    using (var sqlCommand = new SQLiteCommand("vacuum;", sqlConnection))
                    { try { sqlCommand.ExecuteNonQuery(); } catch { } }
                }
            }
            mySqlParameter2 = null;
        }





        private async void GetFio_Click(object sender, EventArgs e)
        {
            _MenuItemEnabled(QuickLoadDataItem, false);
            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(QuickSettingsItem, false);
            _MenuItemEnabled(AnualDatesMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            _MenuItemEnabled(GetFioItem, false);

            await Task.Run(() => CheckAliveServer(sServer1, sServer1UserName, sServer1UserPassword));

            if (bServer1Exist)
            {
                Task.Run(() => _timer1Enabled(true));
                _ProgressBar1Value0();
                dataGridView1.Visible = true;
                pictureBox1.Visible = false;

                DeleteAllDataInTableQuery(databasePerson, "PersonTemp");
                DeleteAllDataInTableQuery(databasePerson, "PersonRegistered");
                DeleteAllDataInTableQuery(databasePerson, "PersonRegisteredFull");

                await Task.Run(() => GetFioFromServers());

                int[] emptyArray = { 0};
                ShowDatatableOnDatagridview(dtPeople, emptyArray, dataGridView1);
                //ShowDataTableQuery(databasePerson, "PersonTemp");

                await Task.Run(() => panelViewResize());
            }
            else { GetInfoSetup(); }

            _MenuItemEnabled(QuickSettingsItem, true);
        }

        private void CheckAliveServer(string serverName, string userName, string userPasswords) //Get the list of registered users
        {
            bServer1Exist = false;
            string stringConnection;
            _toolStripStatusLabelSetText(StatusLabel2, "Проверка сервера " + serverName + ". Ждите окончания процесса...");
            stimerPrev = "Проверка сервера " + serverName + ". Ждите окончания процесса...";
            stringConnection = "Data Source=" + serverName + "\\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID=" + userName + ";Password=" + userPasswords + "; Connect Timeout=5";

            try
            {
                using (var sqlConnection = new SqlConnection(stringConnection))
                {
                    sqlConnection.Open();
                    string query = "SELECT id FROM OBJ_PERSON ";
                    using (var sqlCommand = new SqlCommand(query, sqlConnection))
                    {
                        using (var sqlReader = sqlCommand.ExecuteReader())
                        {
                            bServer1Exist = true;
                        }
                    }
                }
            }
            catch
            { bServer1Exist = false; }

            if (bServer1Exist)
            { _MenuItemEnabled(GetFioItem, true); }
            else
            {
                _toolStripStatusLabelSetText(StatusLabel2, "Ошибка доступа к " + serverName + " SQL БД СКД-сервера!");
                _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                stimerPrev = serverName + " не доступен или неправильная авторизация";

                _MenuItemEnabled(QuickLoadDataItem, false);
                CheckBoxesFiltersAll_Enable(false);
                VisualItemsAll_Enable(false);
                MessageBox.Show(serverName + " не доступен или не правильные имя/пароль", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            stringConnection = null;
        }

        private void GetFioFromServers() //Get the list of registered users
        {
            dtPeople.Clear();

            string stringConnection;
            List<string> ListFIOTemp = new List<string>();
            listFIO = new List<string>();
            try
            {
                _comboBoxFioClr();
                _toolStripStatusLabelSetText(StatusLabel2, "Запрашиваю списки персонала с " + sServer1 + ". Ждите окончания процесса...");
                stimerPrev = "Запрашиваю списки персонала с " + sServer1 + ". Ждите окончания процесса...";
                stringConnection = "Data Source=" + sServer1 + "\\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID=" + sServer1UserName + ";Password=" + sServer1UserPassword + "; Connect Timeout=60";
                using (var sqlConnection = new SqlConnection(stringConnection))
                {
                    sqlConnection.Open();

                    //Check string!!!   "SELECT id, name, surname, patronymic, id, tabnum FROM OBJ_PERSON "
                    string query = "SELECT id, name, surname, patronymic, id, tabnum FROM OBJ_PERSON ";
                    using (var sqlCommand = new SqlCommand(query, sqlConnection))
                    {
                        using (var sqlReader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in sqlReader)
                            {
                                _ProgressWork1();
                                try
                                {
                                    string id = record?["id"].ToString();
                                    if (record?["name"].ToString().Trim().Length > 0)
                                    {
                                        DataRow row = dtPeople.NewRow();
                                        row[1] = record["name"].ToString().Trim() + " " + record["surname"].ToString().Trim() + " " + record["patronymic"].ToString().Trim();
                                        row[2] = record["tabnum"].ToString().Trim();
                                        dtPeople.Rows.Add(row);

                                        listFIO.Add(record["name"].ToString().Trim() + "|" + record["surname"].ToString().Trim() + "|" + record["patronymic"].ToString().Trim() + "|" + record["id"].ToString().Trim() + "|" +
                                                    record["tabnum"].ToString().Trim() + "|" + sServer1);
                                        ListFIOTemp.Add(record["name"].ToString().Trim() + " " + record["surname"].ToString().Trim() + " " + record["patronymic"].ToString().Trim() + "|" + record["tabnum"].ToString().Trim());
                                    }
                                }
                                catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                                _ProgressWork1();
                            }
                        }
                    }
                }

                _toolStripStatusLabelSetText(StatusLabel2, "Все списки с ФИО с серверов СКД успешно получены");
                stimerPrev = "Все списки с ФИО с сервера СКД успешно получены";
            }
            catch (Exception Expt)
            {
                bServer1Exist = false;
                stimerPrev = "Сервер не доступен или неправильная авторизация";
                MessageBox.Show(Expt.ToString(), @"Сервер не доступен или неправильная авторизация", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //Remove dublicat and output Result into combobox1
            IEnumerable<string> ListFIOCombo = ListFIOTemp.Distinct();
            iFIO = 0;

            if (databasePerson.Exists && bServer1Exist)
            {
                DeleteAllDataInTableQuery(databasePerson, "PersonsLastList");
                DeleteAllDataInTableQuery(databasePerson, "PersonsLastComboList");

                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();

                    var sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();

                    foreach (var str in listFIO.ToArray())
                    {
                        if (str != null && str.Length > 1)
                        {
                            using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'PersonsLastList' (PersonsList, Reserv1, Reserv2) " +
                                " VALUES (@PersonsList, @Reserv1, @Reserv2)", sqlConnection))
                            {
                                sqlCommand.Parameters.Add("@PersonsList", DbType.String).Value = str;
                                sqlCommand.Parameters.Add("@Reserv1", DbType.String).Value = "";
                                sqlCommand.Parameters.Add("@Reserv2", DbType.String).Value = "";
                                try { sqlCommand.ExecuteNonQuery(); } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                            }
                        }
                    }
                    foreach (string str in ListFIOCombo.ToArray())
                    {
                        if (str != null && str.Length > 1)
                        {
                            using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'PersonsLastComboList' (ComboList, Reserv1, Reserv2) " +
                            " VALUES (@ComboList, @Reserv1, @Reserv2)", sqlConnection))
                            {
                                sqlCommand.Parameters.Add("@ComboList", DbType.String).Value = str;
                                sqlCommand.Parameters.Add("@Reserv1", DbType.String).Value = "";
                                sqlCommand.Parameters.Add("@Reserv2", DbType.String).Value = "";
                                try { sqlCommand.ExecuteNonQuery(); } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                            }
                        }
                    }
                    sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                    sqlCommand1.Dispose();
                }

                foreach (string str in ListFIOCombo.ToArray())
                { _comboBoxFioAdd(str); iFIO++; }
                try
                { _comboBoxFioIndex(0); }
                catch { };
                _timer1Enabled(false);
                _toolStripStatusLabelSetText(StatusLabel2, "Получено ФИО - " + iFIO + " ");
                _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);
                _MenuItemEnabled(QuickLoadDataItem, true);
            }
            _toolStripStatusLabelForeColor(StatusLabel1, Color.Black);
            stringConnection = null;
            ListFIOCombo = null;
            ListFIOTemp = null;
            if (!bServer1Exist)
            {
                _toolStripStatusLabelSetText(StatusLabel2, "Ошибка доступа к SQL БД СКД-серверов!");
                _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                _MenuItemEnabled(QuickLoadDataItem, false);
                MessageBox.Show("Проверьте правильность написания серверов,\nимя и пароль sa-администратора,\nа а также доступность серверов и их баз!");
            }
            _ProgressBar1Value100();
            _MenuItemEnabled(GetFioItem, true);
            _MenuItemEnabled(FunctionMenuItem, true);
            _MenuItemEnabled(QuickSettingsItem, true);
            _MenuItemEnabled(AnualDatesMenuItem, true);
            _MenuItemEnabled(GroupsMenuItem, true);
        }

        /*
        private void FilterItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Visible = true;
            pictureBox1.Visible = false;

            DeleteAllDataInTableQuery(databasePerson, "PersonTemp");
            if (nameOfLastTableFromDB == "PersonGroupDesciption" || nameOfLastTableFromDB == "PersonGroup")
            {
                GetGroupInfoFromDB();

                if (textBoxGroup.Text.Trim().Length > 0)
                {
                    string[] sCell;
                    foreach (string sRow in lListFIOTemp.ToArray())
                    {
                        sCell = Regex.Split(sRow, "[|]"); //FIO|NAV|H|M
                        if (sCell[0].Length > 1)
                        {
                            _textBoxSetText(textBoxFIO, sCell[0]);   //иммитируем выбор данных
                            _textBoxSetText(textBoxNav, sCell[1]);   //Select person                  
                            
                            dControlHourSelected = Convert.ToDecimal(sCell[2]);
                            dControlMinuteSelected = Convert.ToDecimal(sCell[3]);
                            FilterDataByNav();
                        }
                    }
                }
                nameOfLastTableFromDB = "PersonGroup";
            }
            else
            {
                if (!checkBoxReEnter.Checked)
                { CopyWholeDataFromOneTableIntoAnother(databasePerson, "PersonTemp", "PersonRegistered"); }
                else
                { FilterDataByNav(); }
                nameOfLastTableFromDB = "PersonRegistered";
            }
            ShowDataTableQuery(databasePerson, "PersonTemp");

            //FilterItem.BackColor = SystemColors.Control;
            panelViewResize();
        }*/



        private void FilterDataByNav()    //Copy Data from PersonRegistered into PersonTemp by Filter(NAV and anual dates or minimalTime or dayoff)
        {
            if (_CheckboxChecked(checkBoxReEnter)) //checkBoxReEnter.Checked
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    HashSet<string> AllDateRegistration = new HashSet<string>();
                    using (var sqlCommand = new SQLiteCommand("SELECT  *, MIN(Comming) AS FirstRegistered FROM PersonRegistered  " +
                        " WHERE NAV like '" + _textBoxReturnText(textBoxNav) + "' GROUP BY FIO, NAV, DateRegistered ORDER BY DateRegistered ASC;", sqlConnection))
                    {                                                                                         //, min (PersonRegistered.comming) AS FirstRegistered
                        using (var reader = sqlCommand.ExecuteReader())
                        {
                            string stringDateRegistered = null;
                            foreach (DbDataRecord record in reader)
                            {
                                try
                                {
                                    if (record != null)
                                    {
                                        stringDateRegistered = null;
                                        stringDateRegistered =
                                            record["FIO"].ToString().Trim() + "|" +
                                            record["NAV"].ToString().Trim() + "|" +
                                            record["iDCard"].ToString().Trim() + "|" +
                                            record["DateRegistered"].ToString().Trim() + "|" +
                                            record["HourComming"].ToString().Trim() + "|" +
                                            record["MinuteComming"].ToString().Trim() + "|" +
                                            record["Comming"].ToString().Trim() + "|" +
                                            record["HourControlling"].ToString().Trim() + "|" +
                                            record["MinuteControlling"].ToString().Trim() + "|" +
                                            record["Controlling"].ToString().Trim() + "|" +
                                            record["ServerOfRegistration"].ToString().Trim() + "|" +
                                            record["Reserv1"].ToString().Trim() + "|" +
                                            record["Reserv2"].ToString().Trim()
                                            ;
                                        AllDateRegistration.Add(stringDateRegistered);

                                        dControlHourSelected = TryParseStringToDecimal(record["HourControlling"].ToString().Trim());
                                        dControlMinuteSelected = TryParseStringToDecimal(record["MinuteControlling"].ToString().Trim());
                                    }
                                }
                                catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                            }
                        }
                    }

                    var sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    //Write found dates with the first time of registration
                    using (var sqlCommand = new SQLiteCommand("INSERT INTO 'PersonTemp' (FIO, NAV, DateRegistered, HourComming, MinuteComming, Comming, HourControlling, MinuteControlling, Controlling, ServerOfRegistration, Reserv1, Reserv2) " +
                        "VALUES (@FIO, @NAV, @DateRegistered, @HourComming, @MinuteComming, @Comming, @HourControlling, @MinuteControlling, @Controlling, @ServerOfRegistration, @Reserv1, @Reserv2)", sqlConnection))
                    {
                        sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                        sqlCommand1.ExecuteNonQuery();

                        foreach (var rowData in AllDateRegistration.ToArray())
                        {
                            var cellData = Regex.Split(rowData, "[|]");

                            sqlCommand.Parameters.Add("@FIO", DbType.String).Value = cellData[0];
                            sqlCommand.Parameters.Add("@NAV", DbType.String).Value = cellData[1];
                            // sqlCommand.Parameters.Add("@iDCard", DbType.String).Value = cellData[2];
                            sqlCommand.Parameters.Add("@DateRegistered", DbType.String).Value = cellData[3];
                            sqlCommand.Parameters.Add("@HourComming", DbType.String).Value = cellData[4];
                            sqlCommand.Parameters.Add("@MinuteComming", DbType.String).Value = cellData[5];
                            sqlCommand.Parameters.Add("@Comming", DbType.Decimal).Value = Convert.ToDecimal(cellData[6]);
                            sqlCommand.Parameters.Add("@HourControlling", DbType.String).Value = cellData[7];
                            sqlCommand.Parameters.Add("@MinuteControlling", DbType.String).Value = cellData[8];
                            sqlCommand.Parameters.Add("@Controlling", DbType.Decimal).Value = Convert.ToDecimal(cellData[9]);
                            sqlCommand.Parameters.Add("@ServerOfRegistration", DbType.String).Value = cellData[10];
                            sqlCommand.Parameters.Add("@Reserv1", DbType.String).Value = cellData[11];
                            sqlCommand.Parameters.Add("@Reserv2", DbType.String).Value = cellData[12];
                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }

                        sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                        sqlCommand1.ExecuteNonQuery();
                    }
                    AllDateRegistration.Clear();
                    sqlCommand1.Dispose();
                }
            }
            else
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    HashSet<string> AllDateRegistration = new HashSet<string>();
                    using (var sqlCommand = new SQLiteCommand("Select * FROM PersonRegistered  " +
                        " where NAV like '" + _textBoxReturnText(textBoxNav) + "' order by FIO, DateRegistered, Comming ASC;", sqlConnection))
                    {
                        using (var reader = sqlCommand.ExecuteReader())
                        {
                            string stringDateRegistered = null;
                            foreach (DbDataRecord record in reader)
                            {
                                try
                                {
                                    if (record != null)
                                    {
                                        stringDateRegistered = null;
                                        stringDateRegistered =
                                            record["FIO"].ToString().Trim() + "|" +
                                            record["NAV"].ToString().Trim() + "|" +
                                            record["iDCard"].ToString().Trim() + "|" +
                                            record["DateRegistered"].ToString().Trim() + "|" +
                                            record["HourComming"].ToString().Trim() + "|" +
                                            record["MinuteComming"].ToString().Trim() + "|" +
                                            record["Comming"].ToString().Trim() + "|" +
                                            record["HourControlling"].ToString().Trim() + "|" +
                                            record["MinuteControlling"].ToString().Trim() + "|" +
                                            record["Controlling"].ToString().Trim() + "|" +
                                            record["ServerOfRegistration"].ToString().Trim() + "|" +
                                            record["Reserv1"].ToString().Trim() + "|" +
                                            record["Reserv2"].ToString().Trim()
                                            ;
                                        AllDateRegistration.Add(stringDateRegistered);

                                        dControlHourSelected = Convert.ToDecimal(record["HourControlling"].ToString().Trim());
                                        dControlMinuteSelected = Convert.ToDecimal(record["MinuteControlling"].ToString().Trim());
                                    }
                                }
                                catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                            }
                        }
                    }

                    var sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    //Write found dates with the first time of registration
                    using (var sqlCommand = new SQLiteCommand("INSERT INTO 'PersonTemp' (FIO, NAV, DateRegistered, HourComming, MinuteComming, Comming, HourControlling, MinuteControlling, Controlling, ServerOfRegistration, Reserv1, Reserv2) " +
                        "VALUES (@FIO, @NAV, @DateRegistered, @HourComming, @MinuteComming, @Comming, @HourControlling, @MinuteControlling, @Controlling, @ServerOfRegistration, @Reserv1, @Reserv2)", sqlConnection))
                    {
                        sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                        sqlCommand1.ExecuteNonQuery();

                        foreach (var rowData in AllDateRegistration.ToArray())
                        {
                            var cellData = Regex.Split(rowData, "[|]");

                            sqlCommand.Parameters.Add("@FIO", DbType.String).Value = cellData[0];
                            sqlCommand.Parameters.Add("@NAV", DbType.String).Value = cellData[1];
                            //        sqlCommand.Parameters.Add("@iDCard", DbType.String).Value = cellData[2];
                            sqlCommand.Parameters.Add("@DateRegistered", DbType.String).Value = cellData[3];
                            sqlCommand.Parameters.Add("@HourComming", DbType.String).Value = cellData[4];
                            sqlCommand.Parameters.Add("@MinuteComming", DbType.String).Value = cellData[5];
                            sqlCommand.Parameters.Add("@Comming", DbType.Decimal).Value = Convert.ToDecimal(cellData[6]);
                            sqlCommand.Parameters.Add("@HourControlling", DbType.String).Value = cellData[7];
                            sqlCommand.Parameters.Add("@MinuteControlling", DbType.String).Value = cellData[8];
                            sqlCommand.Parameters.Add("@Controlling", DbType.Decimal).Value = Convert.ToDecimal(cellData[9]);
                            sqlCommand.Parameters.Add("@ServerOfRegistration", DbType.String).Value = cellData[10];
                            sqlCommand.Parameters.Add("@Reserv1", DbType.String).Value = cellData[11];
                            sqlCommand.Parameters.Add("@Reserv2", DbType.String).Value = cellData[12];
                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }

                        sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                        sqlCommand1.ExecuteNonQuery();
                    }
                    AllDateRegistration.Clear();
                    sqlCommand1.Dispose();
                }
            }

            if (_CheckboxChecked(checkBoxWeekend))//checkBoxWeekend.Checked
            { DeleteAnualDates(databasePerson, "PersonTemp"); }

            if (_CheckboxChecked(checkBoxStartWorkInTime)) //checkBoxStartWorkInTime.Checked
            { DeleteDataTableQueryLess(databasePerson, "PersonTemp", "Comming", dControlHourSelected + (dControlMinuteSelected + 1) / 60 - (1 / 60)); }
        }

        private void DeleteAnualDates(System.IO.FileInfo databasePerson, string myTable) //Exclude Anual Days from the table "PersonTemp" DB
        {
            var oneDay = TimeSpan.FromDays(1);

            var mySelectedStartDay = new DateTime(dateTimePickerStart.Value.Year, dateTimePickerStart.Value.Month, dateTimePickerStart.Value.Day);
            var mySelectedEndDay = new DateTime(dateTimePickerEnd.Value.Year, dateTimePickerEnd.Value.Month, dateTimePickerEnd.Value.Day);
            int myYearNow = DateTime.Now.Year;
            var myMonthCalendar = new MonthCalendar();

            myMonthCalendar.MaxSelectionCount = 60;
            myMonthCalendar.SelectionRange = new SelectionRange(mySelectedStartDay, mySelectedEndDay);
            myMonthCalendar.FirstDayOfWeek = Day.Monday;

            for (int year = 0; year < 4; year++)
            {
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow - year, 1, 1));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow - year, 1, 2));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow - year, 3, 8));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow - year, 5, 1));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow - year, 5, 2));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow - year, 5, 9));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow - year, 6, 28));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow - year, 8, 24));    // (plavayuschaya data)
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow - year, 10, 16));   // (plavayuschaya data)
            }

            // Алгоритм для вычисления католической Пасхи http://snippets.dzone.com/posts/show/765
            int Y = myYearNow;
            int a = Y % 19;
            int b = Y / 100;
            int c = Y % 100;
            int d = b / 4;
            int e = b % 4;
            int f = (b + 8) / 25;
            int g = (b - f + 1) / 3;
            int h = (19 * a + b - d - g + 15) % 30;
            int i = c / 4;
            int k = c % 4;
            int L = (32 + 2 * e + 2 * i - h - k) % 7;
            int m = (a + 11 * h + 22 * L) / 451;
            int monthEaster = (h + L - 7 * m + 114) / 31;
            int dayEaster = ((h + L - 7 * m + 114) % 31) + 1;

            //Easter - Paskha
            myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, monthEaster, dayEaster) + oneDay);
            //Independence day
            DateTime dayBolded = new DateTime(myYearNow, 8, 24);
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Sunday:
                    myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, 8, 24) + oneDay);    // (plavayuschaya data)
                    break;
                default:
                    break;
            }
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Saturday:
                    myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, 8, 24) + oneDay + oneDay);    // (plavayuschaya data)
                    break;
                default:
                    break;
            }
            //day of Ukraine Force
            dayBolded = new DateTime(myYearNow, 10, 16);
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Sunday:
                    myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, 10, 16) + oneDay);    // (plavayuschaya data)
                    break;
                default:
                    break;
            }
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Saturday:
                    myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, 10, 16) + oneDay + oneDay);    // (plavayuschaya data)
                    break;
                default:
                    break;
            }

            string singleDate = null;

            for (var myDate = myMonthCalendar.SelectionStart; myDate <= myMonthCalendar.SelectionEnd; myDate += oneDay)
            {
                if (myDate.DayOfWeek == DayOfWeek.Saturday || myDate.DayOfWeek == DayOfWeek.Sunday)
                {
                    singleDate = Regex.Split(myDate.ToString("yyyy-MM-dd"), " ")[0].Trim();
                    DeleteDataTableQueryNAV(databasePerson, myTable, "DateRegistered", singleDate);
                }
            }
            foreach (var myAnualDate in myMonthCalendar.AnnuallyBoldedDates)
            {
                for (var myDate = myMonthCalendar.SelectionStart; myDate <= myMonthCalendar.SelectionEnd; myDate += oneDay)
                {
                    if (myDate == myAnualDate)
                    {
                        singleDate = Regex.Split(myDate.ToString("yyyy-MM-dd"), " ")[0].Trim();
                        DeleteDataTableQueryNAV(databasePerson, "PersonTemp", "DateRegistered", singleDate);
                    }
                }
            }
        }

        private void BoldAnualDates() //Excluded Anual Days from the table "PersonTemp" DB
        {
            var oneDay = TimeSpan.FromDays(1);
            var twoDays = TimeSpan.FromDays(2);
            var fiftyDays = TimeSpan.FromDays(50);

            var mySelectedStartDay = new DateTime(dateTimePickerStart.Value.Year, dateTimePickerStart.Value.Month, dateTimePickerStart.Value.Day);
            var mySelectedEndDay = new DateTime(dateTimePickerEnd.Value.Year, dateTimePickerEnd.Value.Month, dateTimePickerEnd.Value.Day);
            int myYearNow = DateTime.Now.Year;

            monthCalendar.MaxSelectionCount = 60;
            monthCalendar.SelectionRange = new SelectionRange(mySelectedStartDay, mySelectedEndDay);
            monthCalendar.FirstDayOfWeek = Day.Monday;

            //Start of the Block Bolded days
            monthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, 1, 1));
            monthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, 1, 2));
            monthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, 1, 7));
            monthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, 3, 8));
            monthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, 5, 1));
            monthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, 5, 2));
            monthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, 5, 9));
            monthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, 6, 28));
            monthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, 8, 24));    // (plavayuschaya data)
            monthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, 10, 16));   // (plavayuschaya data)

            // Алгоритм для вычисления католической Пасхи    http://snippets.dzone.com/posts/show/765
            int Y = myYearNow;
            int a = Y % 19;
            int b = Y / 100;
            int c = Y % 100;
            int d = b / 4;
            int e = b % 4;
            int f = (b + 8) / 25;
            int g = (b - f + 1) / 3;
            int h = (19 * a + b - d - g + 15) % 30;
            int i = c / 4;
            int k = c % 4;
            int L = (32 + 2 * e + 2 * i - h - k) % 7;
            int m = (a + 11 * h + 22 * L) / 451;
            int monthEaster = (h + L - 7 * m + 114) / 31;
            int dayEaster = ((h + L - 7 * m + 114) % 31) + 1;

            monthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, monthEaster, dayEaster) + oneDay);   //Easter - Paskha

            //Independence day
            DateTime dayBolded = new DateTime(myYearNow, 8, 24);
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Sunday:
                    monthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, 8, 24) + oneDay);
                    break;
                default:
                    break;
            }
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Saturday:
                    monthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, 8, 24) + twoDays);
                    break;
                default:
                    break;
            }

            //day of Ukraine Force
            dayBolded = new DateTime(myYearNow, 10, 16);
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Sunday:
                    monthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, 10, 16) + oneDay);
                    break;
                default:
                    break;
            }
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Saturday:
                    monthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, 10, 16) + twoDays);
                    break;
                default:
                    break;
            }

            //Cristmas day
            dayBolded = new DateTime(myYearNow, 7, 1);
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Sunday:
                    monthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, 7, 1) + oneDay);
                    break;
                default:
                    break;
            }
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Saturday:
                    monthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, 7, 1) + twoDays);
                    break;
                default:
                    break;
            }

            //Troitsa
            dayBolded = new DateTime(myYearNow, monthEaster, dayEaster) + fiftyDays;
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Sunday:
                    monthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, monthEaster, dayEaster) + fiftyDays + oneDay);
                    break;
                default:
                    break;
            }
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Saturday:
                    monthCalendar.AddAnnuallyBoldedDate(new DateTime(myYearNow, monthEaster, dayEaster) + fiftyDays + twoDays);
                    break;
                default:
                    break;
            }

            //incorrect for the days less 50 after and before every New Year
            for (var myDate = monthCalendar.SelectionStart - fiftyDays - fiftyDays; myDate <= monthCalendar.SelectionEnd + fiftyDays + fiftyDays; myDate += oneDay)     // Sunday and Saturday
            {
                if (myDate.DayOfWeek == DayOfWeek.Saturday || myDate.DayOfWeek == DayOfWeek.Sunday)
                    monthCalendar.AddAnnuallyBoldedDate(myDate);
            }

            var today = DateTime.Today;
            monthCalendar.SelectionStart = today;
            monthCalendar.SelectionEnd = today;
            monthCalendar.Update();
        }

        private void ExportDatagridToExcel()  //Export to Excel from DataGridView
        {
            _MenuItemEnabled(QuickLoadDataItem, false);
            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(QuickSettingsItem, false);
            _MenuItemEnabled(AnualDatesMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            _MenuItemEnabled(QuickSettingsItem, false);
            //_MenuItemEnabled(ViewMenuItem, false);
            VisualItemsAll_Enable(false);
            _controlEnable(dataGridView1, false);

            int iDGCollumns = _dataGridView1ColumnCount();
            int iDGRows = _dataGridView1RowsCount();
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);           //Книга
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);    //Таблица.
            ExcelApp.Columns.ColumnWidth = iDGCollumns;

            for (int i = 0; i < iDGCollumns; i++)
            {
                ExcelApp.Cells[1, 1 + i] = _dataGridView1ColumnHeaderText(i);
                ExcelApp.Columns[1 + i].NumberFormat = "@";
                ExcelApp.Columns[1 + i].AutoFit();
            }

            for (int i = 0; i < iDGRows; i++)
            {
                for (int j = 0; j < iDGCollumns; j++)
                { ExcelApp.Cells[i + 2, j + 1] = _dataGridView1CellValue(i, j); }
            }

            ExcelApp.Visible = true;      //Вызываем нашу созданную эксельку.
            ExcelApp.UserControl = true;
            _ChangeMenuItemBackColor(ExportIntoExcelItem, SystemColors.Control);
            stimerPrev = "";
            _timer1Enabled(false);
            _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);
            sLastSelectedElement = "ExportExcel";
            iDGCollumns = 0; iDGRows = 0;
            _toolStripStatusLabelSetText(StatusLabel2, "Готово!");
            _MenuItemEnabled(QuickLoadDataItem, true);
            _MenuItemEnabled(FunctionMenuItem, true);
            _MenuItemEnabled(QuickSettingsItem, true);
            _MenuItemEnabled(AnualDatesMenuItem, true);
            _MenuItemEnabled(GroupsMenuItem, true);
            _MenuItemEnabled(QuickSettingsItem, true);
            //_MenuItemEnabled(ViewMenuItem, true);
            VisualItemsAll_Enable(true);
            _controlEnable(dataGridView1, true);
        }

        private void releaseObject(object obj) //for function - ExportDatagridToExcel()
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object of Excel \n" + ex);
            }
            finally
            { GC.Collect(); }
        }

        private async void Export_Click(object sender, EventArgs e)
        {
            Task.Run(() => _timer1Enabled(true));
            Task.Run(() => _toolStripStatusLabelSetText(StatusLabel2, "Генерирую Excel-файл"));
            stimerPrev = "Наполняю файл данными из текущей таблицы";

            await Task.Run(() => ExportDatagridToExcel());
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        { SelectDataFromCombobox(); }

        private void SelectDataFromCombobox()
        {
            string sComboboxFIO;
            textBoxFIO.Text = "";
            textBoxNav.Text = "";
            CheckBoxesFiltersAll_Enable(false);

            if (nameOfLastTableFromDB == "PersonGroup")
            {
                labelGroup.BackColor = Color.PaleGreen;
            }
            try
            {
                sComboboxFIO = comboBoxFio.SelectedItem.ToString().Trim();
                textBoxNav.Text = Regex.Split(sComboboxFIO, "[|]")[1].Trim();
                textBoxFIO.Text = Regex.Split(sComboboxFIO, "[|]")[0].Trim();
                StatusLabel2.Text = @"Выбран: " + ShortFIO(textBoxFIO.Text) + @" |  Всего ФИО: " + iFIO;
            }
            catch { }
            if (comboBoxFio.SelectedIndex > -1)
            {
                QuickLoadDataItem.BackColor = Color.PaleGreen;
                groupBoxPeriod.BackColor = Color.PaleGreen;
                groupBoxTime.BackColor = Color.PaleGreen;
                groupBoxRemoveDays.BackColor = SystemColors.Control;
            }
            sComboboxFIO = null;
            nameOfLastTableFromDB = "PersonRegistered";
        }

        private void CreateGroupItem_Click(object sender, EventArgs e)
        {
            if (textBoxGroup.Text.Trim().Length > 0)
            {
                decimal controlStartHours = numUpDownHour.Value + (numUpDownMinute.Value + 1) / 60 - (1 / 60);
                using (var connection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    connection.Open();
                    using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PersonGroupDesciption' (GroupPerson, GroupPersonDescription) " +
                                            "VALUES (@GroupPerson, @GroupPersonDescription )", connection))
                    {
                        command.Parameters.Add("@GroupPerson", DbType.String).Value = textBoxGroup.Text.Trim();
                        command.Parameters.Add("@GroupPersonDescription", DbType.String).Value = textBoxGroupDescription.Text.Trim();
                        try { command.ExecuteNonQuery(); } catch { }
                    }
                }
                try { StatusLabel2.Text = "Группа - \"" + textBoxGroup.Text.Trim() + "\" создана"; } catch { }
                controlStartHours = 0;
                PersonOrGroupItem.Text = "Работа с одной персоной";
                nameOfLastTableFromDB = "PersonGroup";

            }
            ListGroup();
        }

        private void ListGroupItem_Click(object sender, EventArgs e)
        { ListGroup(); }

        private void ListGroup()
        {
            groupBoxProperties.Visible = false;
            dataGridView1.Visible = true;

            ShowDataTableQuery(databasePerson, "PersonGroupDesciption", "SELECT GroupPerson, GroupPersonDescription ", " group by GroupPerson ORDER BY GroupPerson asc; ");

            try
            {
                textBoxGroup.Text = dataGridView1.Rows[0].Cells[0].Value.ToString();
                textBoxGroupDescription.Text = dataGridView1.Rows[0].Cells[1].Value.ToString();
                StatusLabel2.Text = @"Выбрана группа: " + textBoxGroup.Text + @" |  Всего ФИО: " + iFIO;

                QuickLoadDataItem.BackColor = Color.PaleGreen;
                groupBoxPeriod.BackColor = Color.PaleGreen;
                groupBoxTime.BackColor = Color.PaleGreen;
                groupBoxRemoveDays.BackColor = SystemColors.Control;
            }
            catch { }
            DeleteGroupItem.Visible = true;
            ExportIntoExcelItem.Enabled = true;

            nameOfLastTableFromDB = "PersonGroupDesciption";
            MembersGroupItem.Enabled = true;
            PersonOrGroupItem.Text = "Работа с одной персоной";
        }

        private void MembersGroupItem_Click(object sender, EventArgs e)
        {
            bErrorData = false;
            MembersGroup();
        }

        private void MembersGroup()
        {
            groupBoxProperties.Visible = false;
            dataGridView1.Visible = true;
            if (textBoxGroup.Text.Trim().Length > 0)
            {
                ShowDataTableQuery(databasePerson, "PersonGroup",
                  "SELECT FIO AS 'Фамилия Имя Отчество', NAV AS 'NAV-код', GroupPerson AS 'Группа'," +
                  " HourControlling AS 'Контрольное время, часы', MinuteControlling AS 'Контрольное время, минуты' ",
                  " Where GroupPerson like '" + textBoxGroup.Text.Trim() + "' ORDER BY FIO");
            }
            if (!bErrorData)
                try
                {
                    comboBoxFio.SelectedIndex = comboBoxFio.FindString(textBoxFIO.Text);
                    textBoxNav.Text = dataGridView1.Rows[0].Cells[1].Value.ToString();
                    textBoxGroup.Text = dataGridView1.Rows[0].Cells[2].Value.ToString();
                    textBoxFIO.Text = dataGridView1.Rows[0].Cells[0].Value.ToString();
                    StatusLabel2.Text = @"Выбрана группа: " + textBoxGroup.Text + @" |  Всего ФИО: " + iFIO;

                    QuickLoadDataItem.BackColor = Color.PaleGreen;
                    groupBoxPeriod.BackColor = Color.PaleGreen;
                    groupBoxTime.BackColor = Color.PaleGreen;
                    groupBoxRemoveDays.BackColor = SystemColors.Control;
                }
                catch { }

            if (iCounterLine > 0)
            {
                PersonOrGroupItem.Text = "Работа с одной персоной";
                nameOfLastTableFromDB = "PersonGroup";
                DeleteGroupItem.Visible = true;
                DeletePersonFromGroupItem.Visible = true;
            }
            else
                ListGroup();
        }

        private void AddPersonToGroupItem_Click(object sender, EventArgs e) //Add the selected person into the named group
        { AddPersonToGroup(); }


        private void importPeopleInLocalDBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ImportTextToTable();
            ImporTableToLocalDB(databasePerson.ToString());
        }

        private void ImportTextToTable() //Fill dtPeople
        {
            List<string> listRows = LoadDataIntoList();
            listGroups = new List<string>();

            string getThreeRows = "";
            getThreeRows = "Маска:\nФИО\tNAV-код\tГруппа\tВремя прихода,часы\tВремя прихода,минуты\tВремя ухода,часы\tВремя ухода,минуты\n\nДанные:\n";
            if (listRows.Count > 0)
            {
                getThreeRows += listRows.ElementAt(0) + "\n";
                if (listRows.Count > 1) getThreeRows += listRows.ElementAt(1) + "\n";
                if (listRows.Count > 2) getThreeRows += listRows.ElementAt(2) + "\n";

                DialogResult result = MessageBox.Show(
                      "Проверьте первые строки файла.\n" +
                      "Первая строка - маска для импорта. Строка заканчивается ячейкой \"Время ухода,минуты\" Разделитель - табуляция\n" +
                      "Если порядок ячеек соответствует маске, то\nдля продолжения импорта нажмите \"Да\":\n\n" + getThreeRows,
                      "Внимание!",
                      MessageBoxButtons.YesNo,
                      MessageBoxIcon.Exclamation,
                      MessageBoxDefaultButton.Button1);
                if (result == DialogResult.Yes)
                {
                    dtPeople.Rows.Clear();
                    DataRow row = dtPeople.NewRow();

                    foreach (string strRow in listRows)
                    {

                        string[] cell = strRow.Split('\t');
                        if (cell.Length == 7)
                        {
                            row[0] = cell[0];
                            row[1] = cell[1];
                            row[2] = cell[2];
                            listGroups.Add(cell[2]);
                            row[3] = cell[3];
                            row[4] = cell[4];
                            row[5] = TryParseStringToDecimal(cell[3]) + (TryParseStringToDecimal(cell[4]) + 1) / 60 - (1 / 60);
                            row[6] = cell[5];
                            row[7] = cell[6];
                            row[8] = TryParseStringToDecimal(cell[5]) + (TryParseStringToDecimal(cell[6]) + 1) / 60 - (1 / 60);

                            dtPeople.Rows.Add(row);
                            row = dtPeople.NewRow();
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("выбранный файл пустой, или \nне подходит для импорта.");
            }
        }

        private decimal TryParseStringToDecimal(string str)
        {
            decimal result = 0;
            try { result = decimal.Parse(str); } catch { result = 0; }
            return result;
        }

        private List<string> LoadDataIntoList() //max List length = 10 000 rows
        {
            int listMaxLength = 10000;
            List<string> listValue = new List<string>(listMaxLength);
            string s = "";
            string filepathLoadedData = "";
            int i = 0; // it is not empty's rows in the selected file

            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Текстовые файлы (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.ShowDialog();
            filepathLoadedData = openFileDialog1.FileName;
            if (filepathLoadedData == null || filepathLoadedData.Length < 1)
            {
                MessageBox.Show("Did not select File!");
            }
            else
            {
                try
                {
                    var Coder = Encoding.GetEncoding(1251);
                    using (System.IO.StreamReader Reader = new System.IO.StreamReader(filepathLoadedData, Coder))
                    {
                        StatusLabel1.Text = "Обрабатываю файл:  " + filepathLoadedData;
                        while ((s = Reader.ReadLine()) != null && i < listMaxLength)
                        {
                            if (s.Trim().Length > 0)
                            {
                                listValue.Add(s.Trim());
                                i++;
                            }
                        }
                    }
                }
                catch (Exception expt) { MessageBox.Show("Error was happened on " + i + " row\n" + expt.ToString()); }
                if (i > listMaxLength - 10 || i == 0)
                {
                    MessageBox.Show("Error was happened on " + i + " row\n You've been chosen the long file!");
                }
            }
            return listValue;
        }

        private void ImporTableToLocalDB(string pathToPersonDB) //use listGroups
        {
            using (var connection = new SQLiteConnection($"Data Source={pathToPersonDB};Version=3;"))
            {
                connection.Open();

                //import groups
                SQLiteCommand commandTransaction = new SQLiteCommand("begin", connection);
                commandTransaction.ExecuteNonQuery();
                using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PersonGroupDesciption' (GroupPerson, GroupPersonDescription) " +
                                        "VALUES (@GroupPerson, @GroupPersonDescription )", connection))
                {
                    foreach (string group in listGroups.Distinct())
                    {
                        command.Parameters.Add("@GroupPerson", DbType.String).Value = group;
                        command.Parameters.Add("@GroupPersonDescription", DbType.String).Value = group;
                        try { command.ExecuteNonQuery(); } catch { }
                    }
                }
                commandTransaction = new SQLiteCommand("end", connection);
                commandTransaction.ExecuteNonQuery();

                //import people
                commandTransaction = new SQLiteCommand("begin", connection);
                commandTransaction.ExecuteNonQuery();
                using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PersonGroup' (FIO, NAV, GroupPerson, HourControlling, MinuteControlling, Controlling, HourControllingOut, MinuteControllingOut, ControllingOut) " +
                                         "VALUES (@FIO, @NAV, @GroupPerson, @HourControlling, @MinuteControlling, @Controlling, @HourControllingOut, @MinuteControllingOut, @ControllingOut)", connection))
                {
                    foreach (DataRow row in dtPeople.Rows)
                    {
                        command.Parameters.Add("@FIO", DbType.String).Value = row[0]; //row[0]
                        command.Parameters.Add("@NAV", DbType.String).Value = row[1];
                        command.Parameters.Add("@GroupPerson", DbType.String).Value = row[2];
                        command.Parameters.Add("@HourControlling", DbType.String).Value = row[3];
                        command.Parameters.Add("@MinuteControlling", DbType.String).Value = row[4];
                        command.Parameters.Add("@Controlling", DbType.Decimal).Value = row[5];
                        command.Parameters.Add("@HourControllingOut", DbType.String).Value = row[6];
                        command.Parameters.Add("@MinuteControllingOut", DbType.String).Value = row[7];
                        command.Parameters.Add("@ControllingOut", DbType.Decimal).Value = row[8];
                        try { command.ExecuteNonQuery(); } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                    }
                }
                commandTransaction = new SQLiteCommand("end", connection);
                commandTransaction.ExecuteNonQuery();
            }
        }

        private void AddPersonToGroup() //Add the selected person into the named group
        {
            string sTextGroup = textBoxGroup.Text.Trim();
            string sTextFIOSelected = comboBoxFio.SelectedItem.ToString().Trim();
            decimal controlStartHours = numUpDownHour.Value + (numUpDownMinute.Value + 1) / 60 - (1 / 60);
            SelectDataFromCombobox();

            using (var connection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
            {
                connection.Open();
                if (sTextGroup.Length > 0)
                {
                    using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PersonGroupDesciption' (GroupPerson, GroupPersonDescription) " +
                                            "VALUES (@GroupPerson, @GroupPersonDescription )", connection))
                    {
                        command.Parameters.Add("@GroupPerson", DbType.String).Value = sTextGroup;
                        command.Parameters.Add("@GroupPersonDescription", DbType.String).Value = textBoxGroupDescription.Text.Trim();
                        try { command.ExecuteNonQuery(); } catch { }
                    }

                }
                if (sTextGroup.Length > 0 && textBoxNav.Text.Trim().Length > 0 && sTextFIOSelected.Length > 10)
                {
                    using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PersonGroup' (FIO, NAV, GroupPerson, HourControlling, MinuteControlling, Controlling) " +
                                            "VALUES (@FIO, @NAV, @GroupPerson, @HourControlling, @MinuteControlling, @Controlling)", connection))
                    {
                        command.Parameters.Add("@FIO", DbType.String).Value = Regex.Split(sTextFIOSelected, "[|]")[0].Trim();
                        command.Parameters.Add("@NAV", DbType.String).Value = Regex.Split(sTextFIOSelected, "[|]")[1].Trim();
                        command.Parameters.Add("@GroupPerson", DbType.String).Value = sTextGroup;
                        command.Parameters.Add("@HourControlling", DbType.String).Value = numUpDownHour.Value.ToString();
                        command.Parameters.Add("@MinuteControlling", DbType.String).Value = numUpDownMinute.Value.ToString();
                        command.Parameters.Add("@Controlling", DbType.Decimal).Value = controlStartHours;
                        try { command.ExecuteNonQuery(); } catch { }
                    }
                    nameOfLastTableFromDB = "PersonGroup";
                    try
                    {
                        StatusLabel2.Text = "\"" + ShortFIO(Regex.Split(sTextFIOSelected, "[|]")[0].Trim()) + "\"" + " добавлен в группу \"" + sTextGroup + "\"";
                        _toolStripStatusLabelBackColor(StatusLabel2, SystemColors.Control);
                        bErrorData = true;
                    }
                    catch { }
                }
                else if (sTextGroup.Length > 0 && textBoxNav.Text.Trim().Length == 0 && sTextFIOSelected.Length > 10)
                    try
                    {
                        StatusLabel2.Text = "Отсутствует NAV-код у:" + ShortFIO(Regex.Split(sTextFIOSelected, "[|]")[0].Trim());
                        _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                        bErrorData = true;
                    }
                    catch { }
                else if (sTextGroup.Length == 0 && textBoxNav.Text.Trim().Length > 0 && sTextFIOSelected.Length > 10)
                    try
                    {
                        StatusLabel2.Text = "Не указана группа, в которую нужно добавить!";
                        _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                        bErrorData = true;
                    }
                    catch { }
                else
                    try
                    {
                        StatusLabel2.Text = "Проверьте вводимые данные!";
                        _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                        bErrorData = true;
                    }
                    catch { }
            }

            MembersGroup();
            PersonOrGroupItem.Text = "Работа с одной персоной";
            controlStartHours = 0;
            sTextGroup = null; sTextFIOSelected = null;
            labelGroup.BackColor = SystemColors.Control;
        }

        private void GetNamePoints() //Get names of the pass by points
        {
            if (databasePerson.Exists)
            {
                listPoints.Clear();
                string stringConnection = @"Data Source=" + sServer1 + @"\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID=" + sServer1UserName + @";Password=" + sServer1UserPassword + @"; Connect Timeout=60";
                string sqlQuery;
                using (var sqlConnection = new SqlConnection(stringConnection))
                {
                    sqlConnection.Open();
                    sqlQuery = "Select id, name FROM OBJ_ABC_ARC_READER;";
                    using (var sqlCommand = new SqlCommand(sqlQuery, sqlConnection))
                    {
                        using (var sqlReader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in sqlReader)
                            {
                                try
                                {
                                    if (record != null && record["id"].ToString().Trim().Length > 0)
                                    { listPoints.Add(sServer1 + "|" + record["id"].ToString().Trim() + "|" + record["name"].ToString().Trim()); }
                                }
                                catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                            }
                        }
                    }
                }

                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    //Write all PointsInfo into the Table  "EquipmentSettings"
                    using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'EquipmentSettings' (EquipmentParameterName, EquipmentParameterValue, EquipmentParameterServer)" +
                        " VALUES (@EquipmentParameterName, @EquipmentParameterValue, @EquipmentParameterServer)", sqlConnection))
                    {
                        foreach (var rowData in listPoints.ToArray())
                        {
                            var cellData = Regex.Split(rowData, "[|]");

                            sqlCommand.Parameters.Add("@EquipmentParameterServer", DbType.String).Value = cellData[0];
                            sqlCommand.Parameters.Add("@EquipmentParameterName", DbType.String).Value = "NamePoint-" + cellData[2];
                            sqlCommand.Parameters.Add("@EquipmentParameterValue", DbType.String).Value = "ValuePoint-" + cellData[1];
                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }
                    }
                }
                stringConnection = null; sqlQuery = null;
            }
        }

        private async void GetDataItem_Click(object sender, EventArgs e)
        { await Task.Run(() => GetData()); }

        private void GetData()
        {
            _changeControlBackColor(groupBoxPeriod, SystemColors.Control);
            _changeControlBackColor(groupBoxTime, SystemColors.Control);
            _ChangeMenuItemBackColor(QuickLoadDataItem, SystemColors.Control);

            _MenuItemEnabled(QuickLoadDataItem, false);
            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(QuickSettingsItem, false);
            _MenuItemEnabled(AnualDatesMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            VisualItemsAll_Enable(false);
            CheckBoxesFiltersAll_Enable(false);
            _controlEnable(dataGridView1, true);
            _controlVisible(pictureBox1, false);

            Task.Run(() => _timer1Enabled(true));

            CheckAliveServer(sServer1, sServer1UserName, sServer1UserPassword);

            if (bServer1Exist)
            {
                //Clear work tables
                DeleteAllDataInTableQuery(databasePerson, "PersonTemp");
                DeleteAllDataInTableQuery(databasePerson, "PersonRegistered");
                DeleteAllDataInTableQuery(databasePerson, "PersonRegisteredFull");
                GetNamePoints();  //Get names of the points

                if ((nameOfLastTableFromDB == "PersonGroupDesciption" || nameOfLastTableFromDB == "PersonGroup") && _textBoxReturnText(textBoxGroup).Length > 0)
                {
                    Task.Run(() => _toolStripStatusLabelSetText(StatusLabel2, "Получаю данные по группе " + _textBoxReturnText(textBoxGroup)));
                    stimerPrev = "Получаю данные по группе " + _textBoxReturnText(textBoxGroup);
                }
                else
                {
                    Task.Run(() => _toolStripStatusLabelSetText(StatusLabel2, "Получаю данные по \"" + ShortFIO(_textBoxReturnText(textBoxFIO)) + "\" "));
                    stimerPrev = "Получаю данные по \"" + ShortFIO(_textBoxReturnText(textBoxFIO)) + "\"";
                    nameOfLastTableFromDB = "PersonRegistered";
                }

                _ProgressBar1Value0();

                GetRegistrations();

                _controlVisible(dataGridView1, true);
                _controlEnable(dataGridView1, true);
                _MenuItemEnabled(VisualItem, true);
                _MenuItemEnabled(ExportIntoExcelItem, true);
                _MenuItemVisible(VisualWorkedTimeItem, true);
                _MenuItemEnabled(QuickSettingsItem, true);
                VisualItemsAll_Enable(true);
                CheckBoxesFiltersAll_Enable(true);

                panelViewResize();
            }
            else
            {
                GetInfoSetup();
                _MenuItemEnabled(QuickSettingsItem, true);
            }
            _changeControlBackColor(groupBoxRemoveDays, Color.PaleGreen);
        }

        private void GetRegistrations()
        {
            _controlVisible(dataGridView1, false);

            dtPersonRegisteredFull.Clear();

            if ((nameOfLastTableFromDB == "PersonGroupDesciption" || nameOfLastTableFromDB == "PersonGroup") && _textBoxReturnText(textBoxGroup).Length > 1)
            {
                GetGroupInfoFromDB();

                string[] sCell;
                foreach (string sRow in lListFIOTemp.ToArray())
                {
                    sCell = Regex.Split(sRow, "[|]"); //FIO|NAV|H|M
                    if (sCell[0].Length > 1)
                    {
                        listFIO.Clear();
                        _textBoxSetText(textBoxFIO, sCell[0]);   //иммитируем выбор данных
                        _textBoxSetText(textBoxNav, sCell[1]);   //Select person                  
                        listFIO.Add(sCell[0]);      // add the person into the list  

                        dControlHourSelected = TryParseStringToDecimal(sCell[2]);
                        dControlMinuteSelected = TryParseStringToDecimal(sCell[3]);

                        Person person = new Person();

                        person.NAV = _textBoxReturnText(textBoxNav);
                        person.FIO = _textBoxReturnText(textBoxFIO);
                        person.HourControlling = dControlHourSelected.ToString();
                        person.MinuteControlling = dControlMinuteSelected.ToString();
                        person.Controlling = dControlHourSelected + (dControlMinuteSelected + 1) / 60 - (1 / 60);
                        
                        GetPersonRegistrationFromServer(dtPersonRegisteredFull, person);     //Search Registration at checkpoints of the selected person
                    }
                }

                nameOfLastTableFromDB = "PersonGroup";
                _timer1Enabled(false);
                _toolStripStatusLabelSetText(StatusLabel2, "Данные по группе \"" + _textBoxReturnText(textBoxGroup) + "\" получены");
            }
            else
            {
                dControlHourSelected = _numUpDownHourReturn(numUpDownHour);
                dControlMinuteSelected =_numUpDownHourReturn( numUpDownMinute);
                
                Person person = new Person();
                person.NAV = _textBoxReturnText(textBoxNav);
                person.FIO = _textBoxReturnText(textBoxFIO);
                person.HourControlling = dControlHourSelected.ToString();
                person.MinuteControlling = dControlMinuteSelected.ToString();
                person.Controlling = dControlHourSelected + (dControlMinuteSelected + 1) / 60 - (1 / 60);

                GetPersonRegistrationFromServer(dtPersonRegisteredFull, person);

                nameOfLastTableFromDB = "PersonRegistered";
                _timer1Enabled(false);
                _toolStripStatusLabelSetText(StatusLabel2, "Данные с СКД по \"" + ShortFIO(_textBoxReturnText(textBoxFIO)) + "\" получены!");
            }


            //CopyWholeDataFromOneTableIntoAnother(databasePerson, "PersonTemp", "PersonRegistered");
            /*
            dtPersonTemp = new DataTable();
            dtPersonTemp = dtPersonRegisteredFull.Copy();
            dtPersonTemp.Columns[27].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[26].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[25].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[24].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[23].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[22].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[20].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[19].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[18].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[15].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[10].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[9].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[6].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[5].ColumnMapping = MappingType.Hidden;
            dtPersonTemp.Columns[0].ColumnMapping = MappingType.Hidden;
            // copyDataTable.Columns.RemoveAt(0); //Remove column with index 0
            */

            int[] hideCollumns = { 27, 26, 25, 24, 23, 22, 20, 19, 18, 15, 10, 9, 6, 5, 0 };
            //заполнить данными из Персонрегистреред
            ShowDatatableOnDatagridview(dtPersonTemp, hideCollumns, dataGridView1);
            _controlVisible(dataGridView1, true);

            _SetMenuItemDefaultColor(QuickLoadDataItem);
            _ChangeMenuItemBackColor(ExportIntoExcelItem, Color.PaleGreen);

            _ProgressBar1Value100();
            _MenuItemEnabled(QuickLoadDataItem, true);
            _MenuItemEnabled(FunctionMenuItem, true);
            _MenuItemEnabled(QuickSettingsItem, true);
            _MenuItemEnabled(AnualDatesMenuItem, true);
            _MenuItemEnabled(GroupsMenuItem, true);
            _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);
            stimerPrev = "";
        }

        private void GetPersonRegistrationFromServer(DataTable dt, Person person)
        {
            DataRow rowPerson;

            //string personNAV = _textBoxReturnText(textBoxNav); 
            //string personFIO = _textBoxReturnText(textBoxFIO);
            string personNAV = person.NAV;
            string personFIO = person.FIO;

            decimal hourControlStart = dControlHourSelected;
            decimal minuteControlStart = dControlMinuteSelected;
            decimal controlStartHours = hourControlStart + (minuteControlStart + 1) / 60 - (1 / 60);
            string stringIdCardIntellect = "";
            string personNAVTemp = "";
            string[] stringSelectedFIO = new string[3];
            try { stringSelectedFIO[0] = Regex.Split(personFIO, "[ ]")[0]; } catch { stringSelectedFIO[0] = ""; }
            try { stringSelectedFIO[1] = Regex.Split(personFIO, "[ ]")[1]; } catch { stringSelectedFIO[1] = ""; }
            try { stringSelectedFIO[2] = Regex.Split(personFIO, "[ ]")[2]; } catch { stringSelectedFIO[2] = ""; }


            stimerPrev = "Получаю данные по \"" + ShortFIO(personFIO) + "\"";
            _toolStripStatusLabelSetText(StatusLabel2, "Получаю данные по \"" + ShortFIO(personFIO) + "\"");

            try
            {
                if (stringIdCardIntellect.Length == 0 && personNAV.Length == 6)
                {
                    string stringConnection = @"Data Source=" + sServer1 + @"\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID=" + sServer1UserName + @";Password=" + sServer1UserPassword + @";Connect Timeout=240";
                    using (var sqlConnection = new SqlConnection(stringConnection))
                    {
                        sqlConnection.Open();
                        using (var cmd = new SqlCommand("Select id, tabnum FROM OBJ_PERSON where tabnum like '%" + personNAV + "%';", sqlConnection))
                        {
                            using (var reader = cmd.ExecuteReader())
                            {
                                foreach (DbDataRecord record in reader)
                                {
                                    try
                                    {
                                        _ProgressWork1();

                                        if (record?["tabnum"].ToString().Trim() == personNAV)
                                        {
                                            stringIdCardIntellect = record["id"].ToString().Trim();
                                            break;
                                        }
                                    } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                                }
                            }
                        }
                    }
                }

                else if (personNAV.Length != 6)
                {
                    foreach (var strRowWithNav in listFIO.ToArray())
                    {
                        _ProgressWork1();
                        if (strRowWithNav.Contains(personNAV) && personNAV.Length > 0 && strRowWithNav.Contains(sServer1))
                            try
                            {
                                stringIdCardIntellect = Regex.Split(strRowWithNav, "[|]")[3].Trim();
                            } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                    }

                    if (stringIdCardIntellect.Length == 0)
                    {
                        foreach (var strRowWithNav in listFIO.ToArray())
                        {
                            if (strRowWithNav.ToLower().Contains(stringSelectedFIO[0].ToLower().Trim()) &&
                                strRowWithNav.ToLower().Contains(stringSelectedFIO[1].ToLower().Trim()) &&
                                strRowWithNav.ToLower().Contains(stringSelectedFIO[2].ToLower().Trim()) &&
                                strRowWithNav.Contains(sServer1))
                            {
                                try { stringIdCardIntellect = Regex.Split(strRowWithNav, "[|]")[3].Trim(); } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                                try { personNAVTemp = Regex.Split(strRowWithNav, "[|]")[4].Trim(); } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                                if (personNAV.Length < 1 && personNAVTemp.Length > 0)
                                { personNAV = personNAVTemp; _ProgressWork1(); break; }
                            }
                        }
                    }
                }
            }
            catch (Exception expt) { MessageBox.Show(expt.ToString()); }

            try
            {
                string stringConnection;
                string stringSqlWhere = "";

                int idCardIntellect = 0;

                decimal hourManaging = 0;
                decimal minuteManaging = 0;
                decimal managingHours = 0;

                try { idCardIntellect = Convert.ToInt32(stringIdCardIntellect); } catch { idCardIntellect = 0; }
                if (idCardIntellect > 0)
                {
                    stringSqlWhere = " where param1 like '" + stringIdCardIntellect + "' AND date >= '" + _dateTimePickerStart() + "' AND date <= '" + _dateTimePickerEnd() + "' ";
                    stringConnection = @"Data Source=" + sServer1 + @"\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID=" + sServer1UserName + ";Password=" + sServer1UserPassword + "; Connect Timeout=240";
                    using (var sqlConnection = new SqlConnection(stringConnection))
                    {
                        sqlConnection.Open();
                        string query = "SELECT param0, param1, objid, convert(varchar, date, 120) as date, convert(varchar, PROTOCOL.time, 114) as time FROM PROTOCOL" + stringSqlWhere + "order by date asc";
                        using (var cmd = new SqlCommand(query, sqlConnection))
                        {
                            using (var reader = cmd.ExecuteReader())
                            {
                                foreach (DbDataRecord record in reader)
                                {
                                    try
                                    {
                                        if (record != null && record["param0"].ToString().Trim().Length > 0)
                                        {
                                            string stringDataNew = Regex.Split(record["date"].ToString().Trim(), "[ ]")[0];
                                            hourManaging = Convert.ToDecimal(Regex.Split(record["time"].ToString().Trim(), "[:]")[0]);
                                            minuteManaging = Convert.ToDecimal(Regex.Split(record["time"].ToString().Trim(), "[:]")[1]);
                                            managingHours = hourManaging + (minuteManaging + 1) / 60;

                                            listRegistrations.Add(
                                                personFIO + "|" + personNAV + "|" + record["param1"].ToString().Trim() + "|" + stringDataNew + "|" +
                                                hourManaging + "|" + minuteManaging + "|" + managingHours.ToString("#.###") + "|" +
                                                hourControlStart + "|" + minuteControlStart + "|" + controlStartHours.ToString("#.###") + "|" +
                                                sServer1 + "|" + record["objid"].ToString().Trim()
                                                );

                                            _ProgressWork1();
                                        }
                                    }
                                    catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                                }
                            }
                        }
                    }
                }

                stringConnection = null;
                stringSqlWhere = null;
                _ProgressWork1();
            }
            catch (Exception Expt)
            { MessageBox.Show(Expt.ToString(), @"Сервер не доступен, или неправильная авторизация", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            iCounterLine = 0;
            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    //Write all registration of person into the Table  "PersonRegistered"
                    var sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                    using (var sqlCommand = new SQLiteCommand("INSERT INTO 'PersonRegistered' (FIO, NAV, iDCard, DateRegistered, HourComming, MinuteComming, Comming, " +
                        " HourControlling, MinuteControlling, Controlling, ServerOfRegistration, Reserv1, Reserv2)" +
                        " VALUES (@FIO, @NAV, @iDCard, @DateRegistered, @HourComming, @MinuteComming, @Comming, " +
                        " @HourControlling, @MinuteControlling, @Controlling, @ServerOfRegistration, @Reserv1, @Reserv2)", sqlConnection))
                    {
                        string[] cellData;
                        string namePoint = "";
                        string nameDirection = "";
                        foreach (var rowData in listRegistrations.ToArray())
                        {
                            cellData = Regex.Split(rowData, "[|]");
                            namePoint = "";
                            nameDirection = "";
                            //Only for the server1 of sServer1
                            foreach (var onePointData in listPoints.ToArray())
                            {
                                if (onePointData != null && cellData[10] != null && cellData[11] != null &&
                                    onePointData.Contains(cellData[10]) && onePointData.Contains(cellData[11]) &&
                                    onePointData.Contains("|") && cellData[11].Length ==
                                    Regex.Split(onePointData, "[|]")[1].Length)
                                    try
                                    {
                                        namePoint = Regex.Split(onePointData, "[|]")[2];
                                        if (namePoint.ToLower().Contains("выход"))
                                            nameDirection = "Выход";
                                        else if (namePoint.ToLower().Contains("вход"))
                                            nameDirection = "Вход";
                                        break;
                                    } catch { }
                            }

                            iCounterLine++;

                            rowPerson = dt.NewRow();
                            /*
                                  new DataColumn("№ п/п",typeof(int)),//0
                                  new DataColumn("Фамилия Имя Отчество",typeof(string)),//1
                                  new DataColumn("NAV-код",typeof(string)),
                                  new DataColumn("Группа",typeof(string)),
                                  new DataColumn("Время прихода,часы",typeof(string)),
                                  new DataColumn("Время прихода,минут",typeof(string)), //5
                                  new DataColumn("Время прихода",typeof(decimal)),
                                  new DataColumn("Время ухода,часы",typeof(string)),
                                  new DataColumn("Время ухода,минут",typeof(string)),
                                  new DataColumn("Время ухода",typeof(decimal)),
                                  new DataColumn("№ пропуска",typeof(int)), //10
                                  new DataColumn("Отдел",typeof(string)),
                                  new DataColumn("Дата регистрации",typeof(string)),
                                  new DataColumn("Время регистрации,часы",typeof(string)),
                                  new DataColumn("Время регистрации,минут",typeof(string)),
                                  new DataColumn("Время регистрации",typeof(decimal)), //15
                                  new DataColumn("Реальное время ухода,часы",typeof(string)),
                                  new DataColumn("Реальное время ухода,минут",typeof(string)),
                                  new DataColumn("Реальное время ухода",typeof(decimal)), //18
                                  new DataColumn("Сервер СКД",typeof(string)), //19
                                  new DataColumn("Имя точки прохода",typeof(string)), //20
                                  new DataColumn("Направление прохода",typeof(string)), //21
 */
                            rowPerson[1] = person.FIO;
                            rowPerson[2] = person.NAV;
                            rowPerson[3] = cellData[2];
                            rowPerson[12] = cellData[3];
                            rowPerson[13] = cellData[4];
                            rowPerson[14] = cellData[5];
                            rowPerson[15] = Convert.ToDecimal(cellData[6]);
                            rowPerson[4] = cellData[7];
                            rowPerson[5] = cellData[8];
                            rowPerson[6] = Convert.ToDecimal(cellData[9]);
                            rowPerson[7] = person.HourControllingOut;
                            rowPerson[8] = person.MinuteControllingOut;
                            rowPerson[9] = person.ControllingOut;
                            rowPerson[20] = namePoint;
                            rowPerson[21] = nameDirection;
                            dt.Rows.Add(rowPerson);

                            sqlCommand.Parameters.Add("@FIO", DbType.String).Value = cellData[0];
                            sqlCommand.Parameters.Add("@NAV", DbType.String).Value = cellData[1];
                            sqlCommand.Parameters.Add("@iDCard", DbType.String).Value = cellData[2];
                            sqlCommand.Parameters.Add("@DateRegistered", DbType.String).Value = cellData[3];
                            sqlCommand.Parameters.Add("@HourComming", DbType.String).Value = cellData[4];
                            sqlCommand.Parameters.Add("@MinuteComming", DbType.String).Value = cellData[5];
                            sqlCommand.Parameters.Add("@Comming", DbType.Decimal).Value = Convert.ToDecimal(cellData[6]);
                            sqlCommand.Parameters.Add("@HourControlling", DbType.String).Value = cellData[7];
                            sqlCommand.Parameters.Add("@MinuteControlling", DbType.String).Value = cellData[8];
                            sqlCommand.Parameters.Add("@Controlling", DbType.Decimal).Value = Convert.ToDecimal(cellData[9]);
                            sqlCommand.Parameters.Add("@ServerOfRegistration", DbType.String).Value = cellData[10];
                            sqlCommand.Parameters.Add("@Reserv1", DbType.String).Value = namePoint;
                            sqlCommand.Parameters.Add("@Reserv2", DbType.String).Value = nameDirection;
                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                            _ProgressWork1();
                        }
                        cellData = new string[1];
                        namePoint = null;
                        nameDirection = null;
                    }
                    listRegistrations.Clear();
                    sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                    sqlCommand1.Dispose();
                }
            }
            if (iCounterLine > 0)
            { bLoaded = true; }

            personNAV = null; personFIO = null; hourControlStart = 0; minuteControlStart = 0; controlStartHours = 0;
            stringIdCardIntellect = null; personNAVTemp = null; stringSelectedFIO = new string[1];
        }

        private void GetGroupInfoFromDB() //Get info the selected group from DB and make a few lists with these data
        {
            lListFIOTemp.Clear();

            using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
            {
                sqlConnection.Open();
                using (var sqlCommand = new SQLiteCommand(
                    "Select * FROM PersonGroup where GroupPerson like '" + _textBoxReturnText(textBoxGroup) + "';", sqlConnection))
                {
                    using (var sqlReader = sqlCommand.ExecuteReader())
                    {
                        string d1 = "", d2 = "", d3 = "", d4 = "";

                        foreach (DbDataRecord record in sqlReader)
                        {
                            d1 = ""; d2 = ""; d3 = ""; d4 = "";
                            try { d1 = record["FIO"].ToString().Trim(); } catch { d1 = ""; }

                            if (record != null && d1.Length > 1)
                            {
                                try { d2 = record["NAV"].ToString().Trim(); } catch { d2 = ""; }
                                try { d3 = record["HourControlling"].ToString().Trim(); } catch { d3 = ""; }
                                try { d4 = record["MinuteControlling"].ToString().Trim(); } catch { d4 = ""; }
                                lListFIOTemp.Add(d1 + "|" + d2 + "|" + d3 + "|" + d4);
                                //FIO|NAV|H|M
                            }
                        }
                        d1 = null; d2 = null; d3 = null; d4 = null;
                    }
                }
            }
        }

        private void DeletePersonFromGroupItem_Click(object sender, EventArgs e)
        {
            int IndexCurrentRow = _dataGridView1CurrentRowIndex();
            if (IndexCurrentRow > -1)
            {
                if (nameOfLastTableFromDB == "PersonGroup")
                {
                    int IndexColumn1 = 0;
                    int IndexColumn2 = 0;

                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        if (dataGridView1.Columns[i].HeaderText.Contains("NAV"))
                            IndexColumn1 = i;
                        if (dataGridView1.Columns[i].HeaderText == "Группа")
                            IndexColumn2 = i;
                    }
                    DeleteDataTableQuery(databasePerson, "PersonGroup", " where NAV like '%" + _textBoxReturnText(textBoxNav) + "%'", "NAV", dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString().Trim(),
                        "GroupPerson", dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn2].Value.ToString().Trim());
                }
                nameOfLastTableFromDB = "PersonGroup";
                PersonOrGroupItem.Text = "Работа с одной персоной";
                ShowDataTableQuery(databasePerson, "PersonGroup", "SELECT FIO, NAV,GroupPerson, HourControlling, MinuteControlling ", " Where GroupPerson like '" + textBoxGroup.Text.Trim() + "' ORDER BY FIO");
            }
        }

        private void DeleteGroupItem_Click(object sender, EventArgs e)
        {
            int IndexCurrentRow = _dataGridView1CurrentRowIndex();
            if (IndexCurrentRow > -1)
            {

                switch (nameOfLastTableFromDB)
                {
                    case "PersonGroupDesciption":
                        int IndexColumn1 = -1;           // индекс 1-й колонки в датагрид

                        for (int i = 0; i < dataGridView1.ColumnCount; i++)
                        {
                            if (dataGridView1.Columns[i].HeaderText == "GroupPerson")
                                IndexColumn1 = i;
                        }

                        if (IndexColumn1 > -1)
                        {
                            DeleteDataTableOneQuery(databasePerson, "PersonGroup", "GroupPerson", dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString().Trim());
                            DeleteDataTableOneQuery(databasePerson, "PersonGroupDesciption", "GroupPerson", dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString().Trim());
                        }
                        break;
                    case "PersonGroup" when textBoxGroup.Text.Trim().Length > 0:
                        DeleteDataTableOneQuery(databasePerson, "PersonGroup", "GroupPerson", textBoxGroup.Text.Trim());
                        DeleteDataTableOneQuery(databasePerson, "PersonGroupDesciption", "GroupPerson", textBoxGroup.Text.Trim());
                        textBoxGroup.BackColor = Color.White;
                        break;
                    default:
                        break;
                }
                PersonOrGroupItem.Text = "Работа с одной персоной";
                nameOfLastTableFromDB = "PersonGroup";
                ShowDataTableQuery(databasePerson, "PersonGroupDesciption", "SELECT GroupPerson, GroupPersonDescription ", " group by GroupPerson ORDER BY GroupPerson asc; ");
                MembersGroupItem.Enabled = true;
            }
            ListGroup();
        }


        private void infoItem_Click(object sender, EventArgs e)
        { ShowDataTableQuery(databasePerson, "TechnicalInfo", "SELECT PCName,POName,POVersion,LastDateStarted ", "ORDER BY LastDateStarted DESC"); }

        private void EnterEditAnualItem_Click(object sender, EventArgs e)
        {
            if (EnterEditAnualItem.Text.Contains(@"Войти в режим редактирования праздников"))
                EnterEditAnual();
            else if (EnterEditAnualItem.Text.Contains(@"Выйти из режим редактирования"))
                ExitEditAnual();
        }

        private void EnterEditAnual()
        {
            dataGridView1.Visible = false;
            panelView.Visible = false;
            FunctionMenuItem.Enabled = false;
            GroupsMenuItem.Enabled = false;
            VisualItemsAll_Enable(false);
            //ViewMenuItem.Enabled = false;
            QuickSettingsItem.Enabled = false;
            QuickLoadDataItem.Enabled = false;
            //FilterItem.Enabled = false;
            CheckBoxesFiltersAll_Enable(false);
            comboBoxFio.Enabled = false;

            StatusLabel2.Text = @"Режим работы с праздниками и выходными";

            AddAnualDateItem.Enabled = true;
            DeleteAnualDateItem.Enabled = true;
            EnterEditAnualItem.Text = @"Выйти из режим редактирования";
            timer1.Start();
        }

        private void ExitEditAnual()
        {
            panelView.Visible = true;
            dataGridView1.Visible = true;

            FunctionMenuItem.Enabled = true;
            GroupsMenuItem.Enabled = true;
            VisualItemsAll_Enable(true);
            //ViewMenuItem.Enabled = true;
            QuickSettingsItem.Enabled = true;
            QuickLoadDataItem.Enabled = true;
            CheckBoxesFiltersAll_Enable(true);
            comboBoxFio.Enabled = true;

            AddAnualDateItem.Enabled = false;
            DeleteAnualDateItem.Enabled = false;
            EnterEditAnualItem.Text = @"Войти в режим редактирования праздников";
            timer1.Stop();
            StatusLabel2.ForeColor = Color.Black;
            StatusLabel2.Text = "Начните работу с кнопки - \"Получить ФИО\"";
        }

        private void AddAnualDateItem_Click(object sender, EventArgs e)
        {
            monthCalendar.AddAnnuallyBoldedDate(monthCalendar.SelectionStart);
            monthCalendar.UpdateBoldedDates();

            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'BoldedDates' (BoldedDate, NAV, Groups, Reserv1, Reserv2) " +
                        " VALUES (@BoldedDate, @NAV, @Groups, @Reserv1, @Reserv2)", sqlConnection))
                    {
                        sqlCommand.Parameters.Add("@BoldedDate", DbType.String).Value = monthCalendar.SelectionStart.ToString();
                        sqlCommand.Parameters.Add("@NAV", DbType.String).Value = textBoxNav.Text;
                        sqlCommand.Parameters.Add("@Groups", DbType.String).Value = textBoxGroup.Text;
                        sqlCommand.Parameters.Add("@Reserv1", DbType.String).Value = "Add";
                        sqlCommand.Parameters.Add("@Reserv2", DbType.String).Value = "";
                        try { sqlCommand.ExecuteNonQuery(); } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                    }
                }
            }
        }

        private void DeleteAnualDateItem_Click(object sender, EventArgs e)
        {
            monthCalendar.RemoveBoldedDate(monthCalendar.SelectionStart);
            monthCalendar.RemoveAnnuallyBoldedDate(monthCalendar.SelectionStart);
            monthCalendar.UpdateBoldedDates();

            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'BoldedDates' (BoldedDate, NAV, Groups, Reserv1, Reserv2) " +
                        " VALUES (@BoldedDate, @NAV, @Groups, @Reserv1, @Reserv2)", sqlConnection))
                    {
                        sqlCommand.Parameters.Add("@BoldedDate", DbType.String).Value = monthCalendar.SelectionStart.ToString();
                        sqlCommand.Parameters.Add("@NAV", DbType.String).Value = textBoxNav.Text;
                        sqlCommand.Parameters.Add("@Groups", DbType.String).Value = textBoxGroup.Text;
                        sqlCommand.Parameters.Add("@Reserv1", DbType.String).Value = "Delete";
                        sqlCommand.Parameters.Add("@Reserv2", DbType.String).Value = "";
                        try { sqlCommand.ExecuteNonQuery(); } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                    }
                }
            }
        }



        //Start of Block. Access to Controls from other threads

        private int _dataGridView1ColumnCount() //add string into  from other threads
        {
            int iDgv = 0;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { try { iDgv = dataGridView1.ColumnCount; } catch { iDgv = 0; } }));
            else
                try { iDgv = dataGridView1.ColumnCount; } catch { iDgv = 0; }
            return iDgv;
        }

        private int _dataGridView1RowsCount() //add string into  from other threads
        {
            int iDgv = 0;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { try { iDgv = dataGridView1.Rows.Count; } catch { iDgv = 0; } }));
            else
                try { iDgv = dataGridView1.Rows.Count; } catch { iDgv = 0; }
            return iDgv;
        }

        private string _dataGridView1ColumnHeaderText(int i) //add string into  from other threads
        {
            string sDgv = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { try { sDgv = dataGridView1.Columns[i].HeaderText; } catch { sDgv = ""; } }));
            else
                try { sDgv = dataGridView1.Columns[i].HeaderText; } catch { sDgv = ""; }
            return sDgv;
        }

        private string _dataGridView1CellValue(int iRow, int iCells) //from other threads
        {
            string sDgv = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { try { sDgv = dataGridView1.Rows[iRow].Cells[iCells].Value.ToString(); } catch { sDgv = ""; } }));
            else
                try { sDgv = dataGridView1.Rows[iRow].Cells[iCells].Value.ToString(); } catch { sDgv = ""; }
            return sDgv;
        }

        private int _dataGridView1CurrentRowIndex() //add string into  from other threads
        {
            int iDgv = -1;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { try { iDgv = dataGridView1.CurrentRow.Index; } catch { iDgv = -1; } }));
            else
                try { iDgv = dataGridView1.CurrentRow.Index; } catch { iDgv = -1; }
            return iDgv;
        }

        private void _dataGridViewSource(DataTable dt)
        {

            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate
                {
                    if (dt != null && dt.Rows.Count > 0)
                    { dataGridView1.DataSource = dt; }
                    else
                    {
                        System.Collections.ArrayList Empty = new System.Collections.ArrayList();
                        dataGridView1.DataSource = Empty;
                    }
                }));
            }
            else
            {
                if (dt != null && dt.Rows.Count > 0)
                { dataGridView1.DataSource = dt; }
                else
                {
                    System.Collections.ArrayList Empty = new System.Collections.ArrayList();
                    dataGridView1.DataSource = Empty;
                }
            }
        }


        private string _textBoxReturnText(TextBox txtBox) //add string into  from other threads
        {
            string tBox = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { tBox = txtBox.Text.Trim(); }));
            else
                tBox = txtBox.Text.Trim();
            return tBox;
        }

        private void _textBoxSetText(TextBox txtBox, string s) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { txtBox.Text = s; }));
            else
                txtBox.Text = s;
        }

        private void _comboBoxFioAdd(string s) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { comboBoxFio.Items.Add(s); }));
            else
                comboBoxFio.Items.Add(s);
        }

        private void _comboBoxFioClr() //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { comboBoxFio.Items.Clear(); }));
            else
                comboBoxFio.Items.Clear();
        }

        private void _comboBoxFioIndex(int i) //add string into comboBoxTargedPC from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { comboBoxFio.SelectedIndex = i; }));
            else
                comboBoxFio.SelectedIndex = i;
        }


        private void _numUpDownHourSet(NumericUpDown numericUpDown, int i) //add string into comboBoxTargedPC from other threads
        {
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate { try { numericUpDown.Value = i; } catch { numUpDownHour.Value = 9; } }));
            }
            else
            {
                try { numericUpDown.Value = i; } catch { numericUpDown.Value = 9; }
            }
        }

        private decimal _numUpDownHourReturn(NumericUpDown numericUpDown)
        {
            decimal iCombo = 0;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { iCombo = (decimal)numericUpDown.Value; }));
            else
                iCombo = (decimal)numericUpDown.Value;
            return iCombo;
        }
               
        private string _dateTimePickerStart() //add string into  from other threads
        {
            string stringDT = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    stringDT = dateTimePickerStart.Value.Year.ToString("0000") + "-" +
                    dateTimePickerStart.Value.Month.ToString("00") +
                        "-" + dateTimePickerStart.Value.Day.ToString("00") + " 00:00:00";
                }));
            else
            {
                stringDT = dateTimePickerStart.Value.Year.ToString("0000") + "-" +
                    dateTimePickerStart.Value.Month.ToString("00") + "-" +
                    dateTimePickerStart.Value.Day.ToString("00") + " 00:00:00";
            }
            return stringDT;
        }

        private string _dateTimePickerEnd() //add string into  from other threads
        {
            string stringDT = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    stringDT = dateTimePickerEnd.Value.Year.ToString("0000") + "-" +
                    dateTimePickerEnd.Value.Month.ToString("00") +
                        "-" + dateTimePickerEnd.Value.Day.ToString("00") + " 23:59:59";
                }));
            else
            {
                stringDT = dateTimePickerEnd.Value.Year.ToString("0000") + "-" +
                    dateTimePickerEnd.Value.Month.ToString("00") + "-" +
                    dateTimePickerEnd.Value.Day.ToString("00") + " 23:59:59";
            }
            return stringDT;
        }

        private string _dateTimePickerReturn(DateTimePicker dateTimePicker) //add string into  from other threads
        {
            string result = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                    { result = dateTimePicker.Value.ToString(); }
                ));
            else
                result = dateTimePicker.Value.ToString();
            return result;
        }


        private void _toolStripStatusLabelSetText(ToolStripStatusLabel statusLabel, string s) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { StatusLabel2.Text = s; }));
            else
                StatusLabel2.Text = s;
        }

        private void _toolStripStatusLabelForeColor(ToolStripStatusLabel statusLabel, Color s)
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { statusLabel.ForeColor = s; }));
            else
                statusLabel.ForeColor = s;
        }

        private void _toolStripStatusLabelBackColor(ToolStripStatusLabel statusLabel, Color s) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { StatusLabel2.BackColor = s; }));
            else
                StatusLabel2.BackColor = s;
        }

        private void _ChangeMenuItemBackColor(ToolStripMenuItem tMenuItem, Color colorMenu) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { tMenuItem.BackColor = colorMenu; }));
            else
                tMenuItem.BackColor = colorMenu; ;
        }

        private void _MenuItemEnabled(ToolStripMenuItem tMenuItem, bool bEnabled) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { tMenuItem.Enabled = bEnabled; }));
            else
                tMenuItem.Enabled = bEnabled;
        }

        private void _MenuItemVisible(ToolStripMenuItem tMenuItem, bool bEnabled) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { tMenuItem.Visible = bEnabled; }));
            else
                tMenuItem.Visible = bEnabled;
        }

        private void _SetMenuItemDefaultColor(ToolStripMenuItem tMenuItem) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { tMenuItem.BackColor = SystemColors.Control; }));
            else
                tMenuItem.BackColor = SystemColors.Control;
        }


        private void _SetCheckboxChecked(CheckBox checkBox, bool checkboxChecked) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { checkBox.Checked = checkboxChecked; }));
            else
                checkBox.Checked = checkboxChecked;
        }

        private bool _CheckboxChecked(CheckBox checkBox) //add string into  from other threads
        {
            bool checkBoxChecked = false;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    checkBoxChecked = checkBox.Checked ? true : false;
                }));
            else
            {
                checkBoxChecked = checkBox.Checked ? true : false;
            }
            return checkBoxChecked;
        }


        private void _panelResume(Panel panel) //access from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    panelView.ResumeLayout();
                }));
            else
            {
                panelView.ResumeLayout();
            }
        }

        private void _panelSetAutoSizeMode(Panel panel, AutoSizeMode state) //access from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    panel.AutoSizeMode = state;
                }));
            else
            {
                panel.AutoSizeMode = state;
            }
        }

        private void _panelSetAutoScroll(Panel panel, bool state) //access from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    panel.AutoScroll = state;
                }));
            else
            {
                panel.AutoScroll = state;
            }
        }

        private void _panelSetAnchor(Panel panel, AnchorStyles anchorStyles) //access from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    panel.Anchor = anchorStyles;
                }));
            else
            {
                panel.Anchor = anchorStyles;
            }
        }

        private void _panelSetHeight(Panel panel, int height) //access from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    panel.Height = height;
                }));
            else
            {
                panel.Height = height;
            }
        }

        private int _panelParentHeightReturn(Panel panel) //access from other threads
        {
            int height = 0;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    height = panelView.Parent.Height;
                }));
            else
            {
                height = panelView.Parent.Height;
            }
            return height;
        }

        private int _panelHeightReturn(Panel panel) //access from other threads
        {
            int height = 0;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    height = panelView.Height;
                }));
            else
            {
                height = panelView.Height;
            }
            return height;
        }

        private int _panelWidthReturn(Panel panel) //access from other threads
        {
            int width = 0;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    width = panelView.Width;
                }));
            else
            {
                width = panelView.Width;
            }
            return width;
        }

        private int _panelControlsCountReturn(Panel panel) //access from other threads
        {
            int count = 0;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    count = panelView.Controls.Count;
                }));
            else
            {
                count = panelView.Controls.Count;
            }
            return count;
        }

        private void _RefreshPictureBox(PictureBox picBox, Bitmap picImage) // не работает
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    picBox.Image = RefreshBitmap(picImage, _panelWidthReturn(panelView) - 2, _panelHeightReturn(panelView) - 2); //сжатая картина
                    picBox.Refresh();
                }));
            else
            {
                picBox.Image = RefreshBitmap(picImage, _panelWidthReturn(panelView) - 2, _panelHeightReturn(panelView) - 2); //сжатая картина
                picBox.Refresh();
            }
        }


        private void _controlVisible(Control control, bool state) //access from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    control.Visible = state;
                }));
            else
            {
                control.Visible = state;
            }
        }

        private void _controlEnable(Control control, bool state) //access from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    control.Enabled = state;
                }));
            else
            {
                control.Enabled = state;
            }
        }

        private void _changeControlBackColor(Control control, Color color) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { control.BackColor = color; }));
            else
                control.BackColor = color; ;
        }


        private string stimerPrev = "";
        private string stimerCurr = "Ждите!";

        private void timer1_Tick(object sender, EventArgs e) //Change a Color of the Font on Status by the Timer
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    if (StatusLabel2.ForeColor == Color.DarkBlue)
                    { StatusLabel2.ForeColor = Color.DarkRed; StatusLabel2.Text = stimerCurr; }
                    else { StatusLabel2.ForeColor = Color.DarkBlue; StatusLabel2.Text = stimerPrev; }
                }));
            else
            {
                if (StatusLabel2.ForeColor == Color.DarkBlue)
                { StatusLabel2.ForeColor = Color.DarkRed; StatusLabel2.Text = stimerCurr; }
                else { StatusLabel2.ForeColor = Color.DarkBlue; StatusLabel2.Text = stimerPrev; }
            }
        }

        private void _timer1Enabled(bool state)
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    timer1.Enabled = state; if (state) timer1.Start(); else timer1.Stop();
                }));
            else
            { timer1.Enabled = state; if (state) timer1.Start(); else timer1.Stop(); }
        }

        private void timer1Start() //Set progressBar Value into 0 from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                { timer1.Start(); }));
            else
            { timer1.Start(); ; }
        }

        private void timer1Stop() //Set progressBar Value into 0 from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                { timer1.Stop(); StatusLabel1.ForeColor = Color.Black; StatusLabel2.ForeColor = Color.Black; }));
            else
            {
                timer1.Stop(); StatusLabel1.ForeColor = Color.Black; StatusLabel2.ForeColor = Color.Black;
            }
        }

        private void _ProgressWork1() //add into progressBar Value 2 from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    if (ProgressBar1.Value > 99)
                    { ProgressBar1.Value = 0; }
                    ProgressBar1.Maximum = 100;
                    ProgressBar1.Value += 1;
                }));
            else
            {
                if (ProgressBar1.Value > 99)
                { ProgressBar1.Value = 0; }
                ProgressBar1.Maximum = 100;
                ProgressBar1.Value += 1;
            }
        }

        private void _ProgressBar1Value0() //Set progressBar Value into 0 from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                { ProgressBar1.Value = 0; }));
            else
            { ProgressBar1.Value = 0; }
        }

        private void _ProgressBar1Value100() //Set progressBar Value into 100 from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    timer1.Stop(); ProgressBar1.Value = 100; StatusLabel1.ForeColor = Color.Black;
                }));
            else
            { timer1.Stop(); ProgressBar1.Value = 100; StatusLabel1.ForeColor = Color.Black; }
        }

        //End of Block. Access to Controls from other threads



        public void CheckBoxesFiltersAll_Enable(bool state)
        {
            _controlEnable(checkBoxStartWorkInTime, state);
            _controlEnable(checkBoxReEnter, state);
            _controlEnable(checkBoxCelebrate, state);
            _controlEnable(checkBoxWeekend, state);
        }

        public void VisualItemsAll_Enable(bool state)
        {
            _MenuItemEnabled(VisualItem, state);
            _MenuItemEnabled(VisualWorkedTimeItem, state);
            _MenuItemEnabled(ReportsItem, state);
        }

        private async void checkBox_CheckStateChanged(object sender, EventArgs e)
        {
            await Task.Run(() => checkBoxCheckStateChanged());
        }

        private void checkBoxCheckStateChanged()
        {
            CheckBoxesFiltersAll_Enable(false);

            _controlVisible(dataGridView1, true);
            _controlVisible(pictureBox1, false);

            DeleteAllDataInTableQuery(databasePerson, "PersonTemp");
            if (nameOfLastTableFromDB == "PersonGroupDesciption" || nameOfLastTableFromDB == "PersonGroup")
            {
                GetGroupInfoFromDB();

                if (_textBoxReturnText(textBoxGroup).Length > 0)
                {
                    string[] sCell;
                    foreach (string sRow in lListFIOTemp.ToArray())
                    {
                        sCell = Regex.Split(sRow, "[|]"); //FIO|NAV|H|M
                        if (sCell[0].Length > 1)
                        {
                            _textBoxSetText(textBoxFIO, sCell[0]);   //иммитируем выбор данных
                            _textBoxSetText(textBoxNav, sCell[1]);   //Select person   

                            dControlHourSelected = TryParseStringToDecimal(sCell[2]);
                            dControlMinuteSelected = TryParseStringToDecimal(sCell[3]);
                            FilterDataByNav();
                        }
                    }
                    sCell = null;
                }
                nameOfLastTableFromDB = "PersonGroup";
            }
            else
            {
                if (!_CheckboxChecked(checkBoxReEnter))
                { CopyWholeDataFromOneTableIntoAnother(databasePerson, "PersonTemp", "PersonRegistered"); }
                else
                { FilterDataByNav(); }
                nameOfLastTableFromDB = "PersonRegistered";
            }


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


            int[] hidecollumns = { 0 };
            ShowDatatableOnDatagridview(dtPeople, hidecollumns, dataGridView1);

            ShowDataTableQuery(databasePerson, "PersonTemp");

            panelViewResize();

            CheckBoxesFiltersAll_Enable(true);
            if (_CheckboxChecked(checkBoxStartWorkInTime))  // if (checkBoxStartWorkInTime.Checked)
            {
                _controlEnable(checkBoxReEnter, true);
                _controlEnable(checkBoxReEnter, false);
                _ChangeMenuItemBackColor(QuickLoadDataItem, SystemColors.Control);
            }
            else if (!_CheckboxChecked(checkBoxStartWorkInTime))  //else if (!checkBoxStartWorkInTime.Checked)
            {
                _controlEnable(checkBoxReEnter, true);
            }
        }


        private void ClearReportItem_Click(object sender, EventArgs e) //ReCreatePersonTables()
        { ReCreatePersonTables(); }

        private void ReCreatePersonTables() //Clear
        {
            DeleteTable(databasePerson, "PersonTemp");
            DeleteTable(databasePerson, "PersonRegistered");
            DeleteTable(databasePerson, "PersonRegisteredFull");
            textBoxFIO.Text = "";
            textBoxGroup.Text = "";
            textBoxGroupDescription.Text = "";
            textBoxNav.Text = "";
            GC.Collect();

            TryMakeDB();
            UpdateTableOfDB();

            ShowDataTableQuery(databasePerson, "PersonRegisteredFull");
            StatusLabel2.Text = @"Отчеты удалены";
        }

        private void ClearDataItem_Click(object sender, EventArgs e) //ReCreateAllPeopleTables()
        { ReCreateAllPeopleTables(); }

        private void ReCreateAllPeopleTables()
        {
            DeleteTable(databasePerson, "PersonRegisteredFull");
            DeleteTable(databasePerson, "PersonRegistered");
            DeleteTable(databasePerson, "PersonTemp");
            DeleteTable(databasePerson, "PersonsLastList");
            DeleteTable(databasePerson, "PersonsLastComboList");
            comboBoxFio.Items.Clear();
            comboBoxFio.SelectedText = "";
            try { comboBoxFio.Text = ""; } catch { }
            textBoxFIO.Text = "";
            textBoxGroup.Text = "";
            textBoxGroupDescription.Text = "";
            textBoxNav.Text = "";
            listFIO.Clear();
            iFIO = 0;
            StatusLabel2.Text = @"База очищена. Остались только созданные группы";
            GC.Collect();
            TryMakeDB();
            UpdateTableOfDB();
        }

        private void ClearAllItem_Click(object sender, EventArgs e) //ReCreate DB
        { ReCreateDB(); }

        private void ReCreateDB()
        {
            if (databasePerson.Exists)
            {
                DeleteTable(databasePerson, "PersonRegisteredFull");
                DeleteTable(databasePerson, "PersonRegistered");
                DeleteTable(databasePerson, "PersonTemp");
                DeleteTable(databasePerson, "PersonGroup");
                DeleteTable(databasePerson, "PersonGroupDesciption");
                DeleteTable(databasePerson, "TechnicalInfo");
                DeleteTable(databasePerson, "BoldedDates");
                DeleteTable(databasePerson, "MySettings");
                DeleteTable(databasePerson, "ProgramSettings");
                DeleteTable(databasePerson, "EquipmentSettings");
                DeleteTable(databasePerson, "ProgramSettings");
                DeleteTable(databasePerson, "PersonsLastList");
                DeleteTable(databasePerson, "PersonsLastComboList");
                GC.Collect();
                comboBoxFio.Items.Clear();
                comboBoxFio.SelectedText = "";
                try { comboBoxFio.Text = ""; } catch { }
                textBoxFIO.Text = "";
                textBoxGroup.Text = "";
                textBoxGroupDescription.Text = "";
                textBoxNav.Text = "";
                listFIO.Clear();
                iFIO = 0;
                TryMakeDB();
            }
            else
            { TryMakeDB(); }
            StatusLabel2.Text = "Все данные удалены. База пересоздана";
        }

        private void VisualItem_Click(object sender, EventArgs e)
        {
            if (bLoaded) { SelectPersonFromDataGrid(); }

            if (bLoaded && (nameOfLastTableFromDB == "PersonRegistered" || nameOfLastTableFromDB == "PersonGroup"))
            {
                dataGridView1.Visible = false;
                FindWorkDatesInSelected();
                DrawRegistration();
                ReportsItem.Visible = true;
                VisualWorkedTimeItem.Visible = true;
                VisualItem.Visible = false;
            }
            else { MessageBox.Show("Таблица с данными пустая.\nНет данных для визуализации!"); }

            if (nameOfLastTableFromDB == "PersonGroup")
            { MessageBox.Show("Визуализация выполняется только для одной выбранной персоны!"); }
        }

        private void CountDataInTheTableQuery(string myTable)
        {
            iCounterLine = 0;
            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();

                    using (var sqlCommand = new SQLiteCommand("Select * FROM '" + myTable + "'; ", sqlConnection))
                    {
                        using (var sqlReader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in sqlReader)
                            {
                                try
                                {
                                    if (record?["id"] != null)
                                    { iCounterLine++; }
                                }
                                catch { }
                            }
                        }
                    }
                }
            }
        }

        private void SelectPersonFromDataGrid()
        {
            int IndexCurrentRow = _dataGridView1CurrentRowIndex();
            if (IndexCurrentRow > -1)
            {
                if (nameOfLastTableFromDB == "PersonGroup")
                {
                    int IndexColumn1 = 0;           // индекс 1-й колонки в датагрид
                    int IndexColumn2 = 0;           // индекс 2-й колонки в датагрид
                    int IndexColumn3 = 0;           // индекс 3-й колонки в датагрид
                    int IndexColumn4 = 0;           // индекс 4-й колонки в датагрид

                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        if (dataGridView1.Columns[i].HeaderText.ToString() == "Фамилия Имя Отчество")
                            IndexColumn1 = i;
                        if (dataGridView1.Columns[i].HeaderText.ToString() == "NAV-код")
                            IndexColumn2 = i;
                        if (dataGridView1.Columns[i].HeaderText.ToString() == "Контрольное время, часы")
                            IndexColumn3 = i;
                        if (dataGridView1.Columns[i].HeaderText.ToString() == "Контрольное время, минуты")
                            IndexColumn4 = i;
                    }

                    textBoxFIO.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString();
                    textBoxNav.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn2].Value.ToString(); //Take the name of selected group
                    StatusLabel2.Text = @"Выбрана группа: " + textBoxGroup.Text + @" | Курсор на: " + textBoxFIO.Text;
                    groupBoxPeriod.BackColor = Color.PaleGreen;
                    numUpDownHour.Value = Convert.ToInt32(dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn3].Value.ToString());
                    numUpDownMinute.Value = Convert.ToInt32(dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn4].Value.ToString());
                }

                if (nameOfLastTableFromDB == "PersonRegistered")
                {
                    int IndexColumn1 = 0;           // индекс 1-й колонки в датагрид
                    int IndexColumn2 = 1;           // индекс 2-й колонки в датагрид

                    textBoxNav.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn2].Value.ToString(); //Take the name of selected group
                    textBoxFIO.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString();
                    StatusLabel2.Text = @"Выбран: " + textBoxFIO.Text + @" |  Всего ФИО: " + iFIO;
                    nameOfLastTableFromDB = "PersonRegistered";
                }
            }
        }

        private void FindWorkDatesInSelected() //Exclude Anual Dates from the table "PersonTemp" DB
        {
            boldeddDates.Clear();
            selectedDates.Clear();
            workSelectedDays = new string[1];
            myBoldedDates = new string[1];

            var oneDay = TimeSpan.FromDays(1);
            var twoDays = TimeSpan.FromDays(2);
            var fiftyDays = TimeSpan.FromDays(50);

            var mySelectedStartDay = new DateTime(dateTimePickerStart.Value.Year, dateTimePickerStart.Value.Month, dateTimePickerStart.Value.Day);
            var mySelectedEndDay = new DateTime(dateTimePickerEnd.Value.Year, dateTimePickerEnd.Value.Month, dateTimePickerEnd.Value.Day);
            DateTime myTempDate;
            int myYearNow = DateTime.Now.Year;

            monthCalendar.SelectionRange = new SelectionRange(mySelectedStartDay, mySelectedEndDay);
            monthCalendar.FirstDayOfWeek = Day.Monday;

            boldeddDates.Add(myYearNow + "-" + 01 + "-" + 01);
            boldeddDates.Add(myYearNow + "-" + 01 + "-" + 02);
            boldeddDates.Add(myYearNow + "-" + 01 + "-" + 07);
            boldeddDates.Add(myYearNow + "-" + 03 + "-" + 08);
            boldeddDates.Add(myYearNow + "-" + 05 + "-" + 01);
            boldeddDates.Add(myYearNow + "-" + 05 + "-" + 02);
            boldeddDates.Add(myYearNow + "-" + 05 + "-" + 09);
            boldeddDates.Add(myYearNow + "-" + 06 + "-" + 28);
            boldeddDates.Add(myYearNow + "-" + 08 + "-" + 24);
            boldeddDates.Add(myYearNow + "-" + 10 + "-" + 16);

            // Алгоритм для вычисления католической Пасхи    http://snippets.dzone.com/posts/show/765
            int Y = myYearNow;
            int a = Y % 19;
            int b = Y / 100;
            int c = Y % 100;
            int d = b / 4;
            int e = b % 4;
            int f = (b + 8) / 25;
            int g = (b - f + 1) / 3;
            int h = (19 * a + b - d - g + 15) % 30;
            int i = c / 4;
            int k = c % 4;
            int L = (32 + 2 * e + 2 * i - h - k) % 7;
            int m = (a + 11 * h + 22 * L) / 451;
            int monthEaster = (h + L - 7 * m + 114) / 31;
            int dayEaster = ((h + L - 7 * m + 114) % 31) + 1;

            //Independence day
            var dayBolded = new DateTime(myYearNow, 8, 24);
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Sunday:
                    myTempDate = new DateTime(myYearNow, 8, 24);
                    boldeddDates.Add(myTempDate.AddDays(1).ToString("yyyy-MM-dd"));
                    break;
                default:
                    break;
            }
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Saturday:
                    myTempDate = new DateTime(myYearNow, 8, 24);
                    boldeddDates.Add(myTempDate.AddDays(2).ToString("yyyy-MM-dd"));
                    break;
                default:
                    break;
            }

            //day of Ukraine Force
            dayBolded = new DateTime(myYearNow, 10, 16);
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Sunday:
                    myTempDate = new DateTime(myYearNow, 10, 16);
                    boldeddDates.Add(myTempDate.AddDays(1).ToString("yyyy-MM-dd"));
                    break;
                default:
                    break;
            }
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Saturday:
                    myTempDate = new DateTime(myYearNow, 10, 16);
                    boldeddDates.Add(myTempDate.AddDays(2).ToString("yyyy-MM-dd"));
                    break;
                default:
                    break;
            }

            //Cristmas day
            dayBolded = new DateTime(myYearNow, 7, 1);
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Sunday:
                    myTempDate = new DateTime(myYearNow, 7, 1);
                    boldeddDates.Add(myTempDate.AddDays(1).ToString("yyyy-MM-dd"));
                    break;
                case (int)Day.Saturday:
                    myTempDate = new DateTime(myYearNow, 7, 1);
                    boldeddDates.Add(myTempDate.AddDays(2).ToString("yyyy-MM-dd"));
                    break;
                default:
                    break;
            }

            //Troitsa
            dayBolded = new DateTime(myYearNow, monthEaster, dayEaster) + fiftyDays;
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Sunday:
                    myTempDate = new DateTime(myYearNow, monthEaster, dayEaster);
                    boldeddDates.Add(myTempDate.AddDays(51).ToString("yyyy-MM-dd"));
                    break;
                case (int)Day.Saturday:
                    myTempDate = new DateTime(myYearNow, monthEaster, dayEaster);
                    boldeddDates.Add(myTempDate.AddDays(52).ToString("yyyy-MM-dd"));
                    break;
                default:
                    break;
            }

            //incorrect for the days less 50 after and before every New Year
            for (var myDate = mySelectedStartDay; myDate <= mySelectedEndDay; myDate += oneDay)     // Sunday and Saturday
            {
                if (myDate.DayOfWeek == DayOfWeek.Saturday || myDate.DayOfWeek == DayOfWeek.Sunday)
                    boldeddDates.Add(myDate.ToString("yyyy-MM-dd"));
            }
            myBoldedDates = boldeddDates.ToArray();

            var aDateTime = mySelectedStartDay;
            bool bDateBolded = false;

            while (aDateTime <= mySelectedEndDay)
            {
                bDateBolded = false;
                foreach (string strBoldedDate in myBoldedDates)
                {
                    if (strBoldedDate.Contains(aDateTime.ToString("yyyy-MM-dd")))
                    {
                        bDateBolded = true;
                        break;
                    }
                }

                if (!bDateBolded)
                { selectedDates.Add(aDateTime.ToString("yyyy-MM-dd")); }
                aDateTime = aDateTime.AddDays(1);
            }
            workSelectedDays = selectedDates.ToArray();
        }

        private void DrawRegistration()  // Visualisation of registration
        {
            //  int iPanelBorder = 2;
            int iMinutesInHour = 60;
            int iShiftStart = 300;
            int iStringHeight = 19;
            int iShiftHeightText = 0;
            int iShiftHeightAll = 36;

            int iHourShouldStart = (int)numUpDownHour.Value * iMinutesInHour + (int)numUpDownMinute.Value;
            int iHourShouldEnd = 1080;
            //     int iHoursWorkDay = 540;
            panelView.Height = iShiftHeightAll + iStringHeight * workSelectedDays.Length;

            //   bmp?.Dispose();
            pictureBox1?.Dispose();
            if (panelView.Controls.Count > 1) panelView.Controls.RemoveAt(1);
            pictureBox1 = new PictureBox
            {
                //    Location = new Point(0, 0),
                SizeMode = PictureBoxSizeMode.AutoSize,
                Size = new Size(
                    iShiftStart + 24 * iMinutesInHour,
                    iShiftHeightAll + iStringHeight * workSelectedDays.Length
                // (iShiftStart + 24 * iMinutesInHour + 2) / 2 // 1740 на 870 - 24 часа и 43 строчки
                // 2 * (iShiftStart + 24 * iMinutesInHour + 2) / 5  //1740 на 696 - 24 часа и 34 строчки

                ),
                BorderStyle = BorderStyle.FixedSingle
            };
            //Disable it for PictureBox set at the Center
            bmp = new Bitmap(pictureBox1.Width, pictureBox1.Height);

            string d0 = "", d1 = "", d2 = "", d7 = "", d8 = "", d9 = "";
            int d3 = 0, d4 = 0, d5 = 0, d6 = 0;
            var font = new Font("Courier", 10, FontStyle.Regular);
            //constant for a person
            string sFIO = "";
            string sNAV = "";
            string sPoint = "";
            string sServer = "";
            int iTimeControling = 0;

            //variable for a person
            string sDatePrevious = "";      //дата в предыдущей выборке
            string sDateCurrent = "";       //дата в текущей выборке
            string sDirectionPrevious = "";
            string sDirectionCurrent = "";
            int iTimeComingPrevious = 0;
            int iTimeComingCurrent = 0;
            int iTimeComing;
            int irectsTempReal = 0;
            bool bTextFIOCorrect = true;

            string query = "";
            if (textBoxNav.TextLength == 6) { query = "Select * FROM PersonRegistered WHERE NAV LIKE '" + textBoxNav.Text + "' ORDER BY NAV, DateRegistered, Comming ASC;"; }
            else if (textBoxNav.TextLength != 6 && textBoxFIO.TextLength > 1) { query = "Select * FROM PersonRegistered WHERE FIO LIKE '" + textBoxFIO.Text + "' ORDER BY NAV, DateRegistered, Comming ASC;"; }
            else { bTextFIOCorrect = false; }

            //Start the Block of Draw 
            //-------------------------------
            //Draw the Axises and common Data
            if (databasePerson.Exists && bTextFIOCorrect)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    using (Graphics gr = Graphics.FromImage(bmp))
                    {
                        var myBrushWorkHour = new SolidBrush(Color.Gray);
                        var myBrushRealWorkHour = new SolidBrush(clrRealRegistration);

                        var axis = new Pen(Color.Black);
                        Rectangle[] rectsReal;
                        var rects = new Rectangle[workSelectedDays.Length];

                        int iLenghtRect = 0; //количество  вх-вых в рабочие дни

                        using (var sqlCommand = new SQLiteCommand(query, sqlConnection))
                        {
                            using (var sqlReader = sqlCommand.ExecuteReader())
                            {
                                foreach (DbDataRecord record in sqlReader)
                                {
                                    d0 = "";
                                    d2 = "";
                                    d8 = ""; d9 = "";
                                    try { d0 = record["FIO"].ToString().Trim(); } catch { }
                                    try { d2 = record["DateRegistered"].ToString().Trim(); } catch { }
                                    try { d8 = record["Reserv1"].ToString().Trim(); } catch { }
                                    try { d9 = record["ServerOfRegistration"].ToString().Trim(); } catch { }
                                    if (d0.Length > 3)    //учитываем проходы только через СтопНЕТ
                                        iLenghtRect += workSelectedDays.Count(t => t.Length == 10 && d2.Contains(t));
                                }
                            }
                            rectsReal = new Rectangle[iLenghtRect];

                            using (var sqlReader = sqlCommand.ExecuteReader())
                            {
                                foreach (DbDataRecord record in sqlReader)
                                {
                                    d0 = ""; d1 = ""; d2 = ""; d3 = 0; d4 = 0; d5 = 0; d6 = 0; d7 = ""; d8 = ""; d9 = ""; iTimeComing = 0;
                                    try { d0 = record["FIO"].ToString().Trim(); } catch { }
                                    try { d1 = record["NAV"].ToString().Trim(); } catch { }
                                    try { d2 = record["DateRegistered"].ToString().Trim(); } catch { }
                                    try { d3 = Convert.ToInt32(record["HourControlling"].ToString().Trim()); } catch { }
                                    try { d4 = Convert.ToInt32(record["MinuteControlling"].ToString().Trim()); } catch { }
                                    try { d5 = Convert.ToInt32(record["HourComming"].ToString().Trim()); } catch { }
                                    try { d6 = Convert.ToInt32(record["MinuteComming"].ToString().Trim()); } catch { }
                                    try { d7 = record["Reserv2"].ToString().Trim(); } catch { }
                                    try { d8 = record["Reserv1"].ToString().Trim(); } catch { }
                                    try { d9 = record["ServerOfRegistration"].ToString().Trim(); } catch { }

                                    //set parameters for persons's constant
                                    if (d0.Length > 1) sFIO = d0;
                                    if (d1.Length == 6) sNAV = d1;
                                    if (d3 * 60 + d4 > 0) iTimeControling = d3 * 60 + d4;
                                    if (d5 * 60 + d6 > 0) iTimeComing = d5 * 60 + d6;
                                    if (d9.Length > 1) sServer = d9;
                                    if (d8.Length > 0) sPoint = d8;

                                    if (d0.Length > 3)   //учитываем проходы только через СтопНЕТ
                                    {
                                        for (int k = 0; k < workSelectedDays.Length; k++)
                                        {
                                            if (workSelectedDays[k].Length == 10 && d2.Contains(workSelectedDays[k]))  //если приход пришел в рабочий день - считаем
                                            {
                                                //date
                                                sDatePrevious = sDateCurrent; sDateCurrent = d2;
                                                //direction
                                                sDirectionPrevious = sDirectionCurrent; sDirectionCurrent = d7;
                                                //TimeComming
                                                iTimeComingPrevious = iTimeComingCurrent; iTimeComingCurrent = iTimeComing;

                                                if (sDirectionCurrent.ToLower().Contains("вход") && sDirectionPrevious.ToLower().Contains("вход") && sDatePrevious.Contains(sDateCurrent))
                                                {
                                                    if (iTimeComingCurrent > iTimeComingPrevious)
                                                    {
                                                        rectsReal[irectsTempReal] = new Rectangle(iShiftStart + iTimeComingPrevious, 2 * iStringHeight + iShiftHeightText + k * iStringHeight + 1, iTimeComingCurrent - iTimeComingPrevious, 3 * iStringHeight / 4);
                                                        irectsTempReal++;
                                                    }
                                                }
                                                else if (sDirectionCurrent.ToLower().Contains("выход") && sDirectionPrevious.ToLower().Contains("вход") && sDatePrevious.Contains(sDateCurrent))
                                                {
                                                    if (iTimeComingCurrent > iTimeComingPrevious)
                                                    {
                                                        rectsReal[irectsTempReal] = new Rectangle(iShiftStart + iTimeComingPrevious, 2 * iStringHeight + iShiftHeightText + k * iStringHeight + 1, iTimeComingCurrent - iTimeComingPrevious, 3 * iStringHeight / 4);
                                                        irectsTempReal++;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            var myBrushAxis = new SolidBrush(Color.Black);
                            var pointForN = new PointF(iShiftStart - 300, iStringHeight + iShiftHeightText);
                            for (int k = 0; k < workSelectedDays.Length; k++)
                            {
                                pointForN.Y += iStringHeight;
                                gr.DrawLine(axis, new Point(0, iShiftHeightAll + k * iStringHeight), new Point(pictureBox1.Width, iShiftHeightAll + k * iStringHeight));
                                gr.DrawString(workSelectedDays[k] + " (" + ShortFIO(sFIO) + ")", font, myBrushAxis, pointForN); //Paint workdates and person's FIO
                            }
                            gr.DrawLine(axis, new Point(0, iShiftHeightAll + workSelectedDays.Length * iStringHeight), new Point(pictureBox1.Width, iShiftHeightAll + workSelectedDays.Length * iStringHeight));
                        }

                        iHourShouldStart = iTimeControling;
                        // наносим рисунки с рабочими часами
                        for (int k = 0; k < workSelectedDays.Length; k++)
                        {
                            rects[k] = new Rectangle(iShiftStart + iHourShouldStart, 2 * iStringHeight + iShiftHeightText + k * iStringHeight + iStringHeight / 4 + 1, iHourShouldEnd - iHourShouldStart, iStringHeight / 4);
                        }
                        //Fill RealWork
                        gr.FillRectangles(myBrushRealWorkHour, rectsReal);
                        // Fill WorkTime
                        gr.FillRectangles(myBrushWorkHour, rects);

                        axis.Dispose();
                        rectsReal = null;
                        rects = null;
                        myBrushRealWorkHour = null;
                        myBrushWorkHour = null;
                    }
                }
            }

            using (Graphics gr = Graphics.FromImage(bmp))
            {
                var myBrushAxis = new SolidBrush(Color.Black);
                var pointForN = new PointF(iShiftStart - 100, iStringHeight + iShiftHeightText);

                var axis = new Pen(Color.Black);
                //рисуем оси дат и делаем к ним подписи
                for (int k = 0; k < workSelectedDays.Length; k++)
                {
                    pointForN.Y += iStringHeight;
                    gr.DrawLine(axis, new Point(0, iShiftHeightAll + k * iStringHeight), new Point(pictureBox1.Width, iShiftHeightAll + k * iStringHeight));
                    //          gr.DrawString(workSelectedDays[k], font, myBrushAxis, pointForN);
                }

                //рисуем оси Часов и делаем к ним подписи
                gr.DrawString("Время, часы:", font, SystemBrushes.WindowText, new Point(iShiftStart - 110, iStringHeight / 4));
                gr.DrawString("Дата (ФИО)", font, SystemBrushes.WindowText, new Point(10, iStringHeight));
                gr.DrawLine(axis, new Point(0, 0), new Point(iShiftStart, iShiftHeightAll));
                gr.DrawLine(axis, new Point(iShiftStart, 0), new Point(iShiftStart, iShiftHeightAll));

                for (int k = 0; k <= 23; k++)
                {
                    gr.DrawLine(axis, new Point(iShiftStart + k * iMinutesInHour, iShiftHeightAll), new Point(iShiftStart + k * iMinutesInHour, Convert.ToInt32(pictureBox1.Height)));
                    gr.DrawString(Convert.ToString(k), font, SystemBrushes.WindowText, new Point(320 + k * iMinutesInHour, iStringHeight));
                }
                gr.DrawLine(axis, new Point(iShiftStart + 24 * iMinutesInHour, iShiftHeightAll), new Point(iShiftStart + 24 * iMinutesInHour, Convert.ToInt32(pictureBox1.Height)));

                axis.Dispose();
                myBrushAxis = null;
            }

            //-------------------------------
            //End the Block Draw


            d0 = null; d1 = null; d2 = null; d3 = 0; d4 = 0; d5 = 0; d6 = 0; d7 = null;
            pictureBox1.Image = bmp;
            pictureBox1.Refresh();
            panelView.Controls.Add(pictureBox1);
            font.Dispose();

            //---------------------------------------------------------------
            //указываем выбранный ФИО и устанавливаем на него фокус
            textBoxFIO.Text = sFIO;
            textBoxNav.Text = sNAV;
            StatusLabel2.Text = @"Выбран: " + ShortFIO(sFIO) + @" |  Всего ФИО: " + iFIO;
            if (comboBoxFio.FindString(sFIO) != -1) comboBoxFio.SelectedIndex = comboBoxFio.FindString(sFIO); //ищем в комбобокс выбранный ФИО и устанавливаем на него фокус

            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'ProgramSettings' (PoParameterName, PoParameterValue) " +
                        " VALUES (@PoParameterName, @PoParameterValue)", sqlConnection))
                    {
                        sqlCommand.Parameters.Add("@PoParameterName", DbType.String).Value = "clrRealRegistration";
                        sqlCommand.Parameters.Add("@PoParameterValue", DbType.String).Value = clrRealRegistration.Name;
                        try { sqlCommand.ExecuteNonQuery(); } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                    }
                }
            }
            sLastSelectedElement = "DrawRegistration";

            _RefreshPictureBox(pictureBox1, bmp);
            panelViewResize();
        }

        private string ShortFIO(string s) //Transform from full FIO into Short form FIO
        {
            var stmp = new string[1];
            try { stmp = Regex.Split(s, "[ ]"); } catch { }
            var sFullNameOnly = "";
            try { sFullNameOnly = stmp[0]; } catch { }
            try { sFullNameOnly += " " + stmp[1].Substring(0, 1) + @"."; } catch { }
            try { sFullNameOnly += " " + stmp[2].Substring(0, 1) + @"."; } catch { }
            stmp = new string[1];
            return sFullNameOnly;
        }

        private void VisualWorkedTimeItem_Click(object sender, EventArgs e)
        {
            if (bLoaded) { SelectPersonFromDataGrid(); }

            if (bLoaded && (nameOfLastTableFromDB == "PersonRegistered" || nameOfLastTableFromDB == "PersonGroup"))
            {
                dataGridView1.Visible = false;
                FindWorkDatesInSelected();
                DrawFullWorkedPeriodRegistration();
                ReportsItem.Visible = true;
                VisualItem.Visible = true;
                VisualWorkedTimeItem.Visible = false;
            }
            else { MessageBox.Show("Таблица с данными пустая.\nНет данных для визуализации!"); }
            if (nameOfLastTableFromDB == "PersonGroup")
            { MessageBox.Show("Визуализация выполняется только с одной выбранной персоны!"); }
        }

        private void DrawFullWorkedPeriodRegistration()  // Draw the whole period registration
        {
            //     int iPanelBorder = 2;
            int iMinutesInHour = 60;
            int iShiftStart = 300;
            int iStringHeight = 19;
            int iShiftHeightText = 0;
            int iShiftHeightAll = 36;

            int iHourShouldStart = (int)numUpDownHour.Value * iMinutesInHour + (int)numUpDownMinute.Value;
            int iHourShouldEnd = 1080;
            //   int iHoursWorkDay = 540;

            panelView.Height = iShiftHeightAll + iStringHeight * workSelectedDays.Length;
            // panelView.AutoScroll = false;
            // panelView.AutoSizeMode = AutoSizeMode.GrowAndShrink;

            // panelView.Anchor = AnchorStyles.Bottom;
            // panelView.Anchor = AnchorStyles.Left;
            // panelView.Dock = DockStyle.None;

            //  bmp?.Dispose();
            panelView.ResumeLayout();
            pictureBox1?.Dispose();
            if (panelView.Controls.Count > 1) panelView.Controls.RemoveAt(1);
            pictureBox1 = new PictureBox
            {
                //    Location = new Point(0, 0),
                SizeMode = PictureBoxSizeMode.AutoSize,
                Size = new Size(
                    iShiftStart + 24 * iMinutesInHour + 2,
                    iShiftHeightAll + iStringHeight * workSelectedDays.Length
                ),
                BorderStyle = BorderStyle.FixedSingle
            };

            //set to comment it if wante to set the PictureBox  at the Center place
            bmp = new Bitmap(pictureBox1.Width, pictureBox1.Height);

            string d0 = "", d1 = "", d2 = "";
            int d3 = 0, d4 = 0, d5 = 0, d6 = 0;
            decimal d7 = 0;

            //Start the Block of Draw 
            //Draw the Axises and common Data
            var font = new Font("Courier", 10, FontStyle.Regular);
            //constant for a person
            string sFIO = "";
            string sNAV = "";
            int iTimeControling = 0;

            //variable for a person
            string sDatePrevious = "";      //дата в предыдущей выборке
            string sDateCurrent = "";       //дата в текущей выборке
            int iTimeComingPrevious = 0;
            int iTimeComingCurrent = 0;
            decimal iTimeComingHourPrevious = 0;
            decimal iTimeComingHourCurrent = 0;

            int iTimeComing = 0;
            int irectsTempReal = 0;
            bool bTextFIOCorrect = true;

            string query = "";
            if (textBoxNav.TextLength == 6) { query = "Select * FROM PersonRegistered WHERE NAV LIKE '" + textBoxNav.Text + "' ORDER BY NAV, DateRegistered, Comming ASC;"; }
            else if (textBoxNav.TextLength != 6 && textBoxFIO.TextLength > 1) { query = "Select * FROM PersonRegistered WHERE FIO LIKE '" + textBoxFIO.Text + "' ORDER BY NAV, DateRegistered, Comming ASC;"; }
            else { bTextFIOCorrect = false; }

            //Start the Block of Draw 
            //-------------------------------

            //Draw the Axises and common Data
            if (databasePerson.Exists && bTextFIOCorrect)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    using (var gr = Graphics.FromImage(bmp))
                    {
                        var myBrushWorkHour = new SolidBrush(Color.Gray);
                        var myBrushRealWorkHour = new SolidBrush(clrRealRegistration);

                        Pen axis = new Pen(Color.Black);
                        Rectangle[] rectsReal;
                        Rectangle[] rectsRealMark;
                        var rects = new Rectangle[workSelectedDays.Length];

                        int iLenghtRect = 0; //количество  вх-вых в рабочие дни

                        using (var sqlCommand = new SQLiteCommand(query, sqlConnection))
                        {
                            using (var sqlReader = sqlCommand.ExecuteReader())
                            {
                                foreach (DbDataRecord record in sqlReader)
                                {
                                    d0 = ""; d2 = "";
                                    try { d0 = record["FIO"].ToString().Trim(); } catch { }
                                    try { d2 = record["DateRegistered"].ToString().Trim(); } catch { }
                                    if (d0.Length > 3)    //учитываем все проходы через все считыватели на всех серверах
                                        iLenghtRect += workSelectedDays.Count(t => t.Length == 10 && d2.Contains(t));
                                }
                            }
                            rectsRealMark = new Rectangle[iLenghtRect];
                            rectsReal = new Rectangle[iLenghtRect];

                            using (var sqlReader = sqlCommand.ExecuteReader())
                            {
                                foreach (DbDataRecord record in sqlReader)
                                {
                                    d0 = ""; d1 = ""; d2 = ""; d3 = 0; d4 = 0; d5 = 0; d6 = 0; d7 = 0; iTimeComing = 0;
                                    try { d0 = record["FIO"].ToString().Trim(); } catch { }
                                    try { d1 = record["NAV"].ToString().Trim(); } catch { }
                                    try { d2 = record["DateRegistered"].ToString().Trim(); } catch { }
                                    try { d3 = Convert.ToInt32(record["HourControlling"].ToString().Trim()); } catch { }
                                    try { d4 = Convert.ToInt32(record["MinuteControlling"].ToString().Trim()); } catch { }
                                    try { d5 = Convert.ToInt32(record["HourComming"].ToString().Trim()); } catch { }
                                    try { d6 = Convert.ToInt32(record["MinuteComming"].ToString().Trim()); } catch { }
                                    try { d7 = Convert.ToDecimal(record["Comming"].ToString().Trim()); } catch { }

                                    //set parameters for persons's constant
                                    if (d0.Length > 1) sFIO = d0;
                                    if (d1.Length == 6) sNAV = d1;
                                    if (d3 * 60 + d4 > 0) iTimeControling = d3 * 60 + d4;
                                    if (d5 * 60 + d6 > 0) iTimeComing = d5 * 60 + d6;

                                    if (d0.Length > 3)   //учитываем все проходы через все считыватели на всех серверах
                                    {
                                        for (int k = 0; k < workSelectedDays.Length; k++)
                                        {
                                            if (workSelectedDays[k].Length == 10 && d2.Contains(workSelectedDays[k])) //учитываем проходы ТОЛЬКО в рабочие дни
                                            {
                                                //Day
                                                sDatePrevious = sDateCurrent;
                                                sDateCurrent = d2;
                                                //TimeComming
                                                iTimeComingPrevious = iTimeComingCurrent;
                                                iTimeComingCurrent = iTimeComing;
                                                iTimeComingHourPrevious = iTimeComingHourCurrent;
                                                iTimeComingHourCurrent = d7;

                                                if (sDatePrevious == sDateCurrent && iTimeComingHourPrevious <= iTimeComingHourCurrent)
                                                { rectsReal[irectsTempReal] = new Rectangle(iShiftStart + iTimeComingPrevious, 2 * iStringHeight + iShiftHeightText + k * iStringHeight + 1, iTimeComingCurrent - iTimeComingPrevious, 3 * iStringHeight / 4); }
                                                irectsTempReal++;
                                            }
                                        }
                                    }
                                }
                            }

                            //Draw axis and Paint FIO, workdays
                            var myBrushAxis = new SolidBrush(Color.Black);
                            var pointForN = new PointF(iShiftStart - 300, iStringHeight + iShiftHeightText);
                            for (int k = 0; k < workSelectedDays.Length; k++)
                            {
                                pointForN.Y += iStringHeight;
                                gr.DrawLine(axis, new Point(0, iShiftHeightAll + k * iStringHeight), new Point(pictureBox1.Width, iShiftHeightAll + k * iStringHeight));
                                gr.DrawString(workSelectedDays[k] + " (" + ShortFIO(sFIO) + ")", font, myBrushAxis, pointForN); //Paint workdays and person's FIO
                            }
                            gr.DrawLine(axis, new Point(0, iShiftHeightAll + workSelectedDays.Length * iStringHeight), new Point(pictureBox1.Width, iShiftHeightAll + workSelectedDays.Length * iStringHeight));
                            myBrushAxis = null;
                        }

                        iHourShouldStart = iTimeControling;
                        //Paint Marks of Registration
                        using (var sqlCommand = new SQLiteCommand("Select * FROM PersonRegistered GROUP BY FIO, NAV, Comming, DateRegistered ORDER BY NAV, DateRegistered, Comming ASC;", sqlConnection))
                        {
                            irectsTempReal = 0;
                            using (var sqlReader = sqlCommand.ExecuteReader())
                            {
                                foreach (DbDataRecord record in sqlReader)
                                {
                                    d0 = ""; d2 = ""; d5 = 0; d6 = 0; iTimeComing = 0;
                                    try { d0 = record["FIO"].ToString().Trim(); } catch { }
                                    try { d2 = record["DateRegistered"].ToString().Trim(); } catch { }
                                    try { d5 = Convert.ToInt32(record["HourComming"].ToString().Trim()); } catch { }
                                    try { d6 = Convert.ToInt32(record["MinuteComming"].ToString().Trim()); } catch { }

                                    //set parameters for persons's constant
                                    if (d5 * 60 + d6 > 0) iTimeComing = d5 * 60 + d6;

                                    if (d0.Length > 3)   //учитываем все проходы через все считыватели на всех серверах
                                    {
                                        for (int k = 0; k < workSelectedDays.Length; k++)
                                        {
                                            if (workSelectedDays[k].Length == 10 && d2.Contains(workSelectedDays[k])) //учитываем проходы ТОЛЬКО в рабочие дни
                                            {
                                                rectsRealMark[irectsTempReal] = new Rectangle(iShiftStart + iTimeComing, 2 * iStringHeight + iShiftHeightText + k * iStringHeight + 1, 2, 3 * iStringHeight / 4);
                                                irectsTempReal++;
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // наносим рисунки с рабочими часами
                        for (int k = 0; k < workSelectedDays.Length; k++)
                        { rects[k] = new Rectangle(iShiftStart + iHourShouldStart, 2 * iStringHeight + iShiftHeightText + k * iStringHeight + iStringHeight / 4 + 1, iHourShouldEnd - iHourShouldStart, iStringHeight / 4); }
                        //Fill RealWork
                        gr.FillRectangles(myBrushRealWorkHour, rectsReal);
                        //Fill All Mark at Passthrow Points
                        gr.FillRectangles(myBrushRealWorkHour, rectsRealMark);
                        // Fill WorkTime
                        gr.FillRectangles(myBrushWorkHour, rects);

                        axis.Dispose();
                        rectsReal = null;
                        rects = null;
                        myBrushRealWorkHour = null;
                        myBrushWorkHour = null;
                    }
                }
            }

            using (Graphics gr = Graphics.FromImage(bmp))
            {
                var myBrushAxis = new SolidBrush(Color.Black);
                var pointForN = new PointF(iShiftStart - 100, iStringHeight + iShiftHeightText);

                var axis = new Pen(Color.Black);
                //рисуем оси дат и делаем к ним подписи
                for (int k = 0; k < workSelectedDays.Length; k++)
                {
                    pointForN.Y += iStringHeight;
                    gr.DrawLine(axis, new Point(0, iShiftHeightAll + k * iStringHeight), new Point(pictureBox1.Width, iShiftHeightAll + k * iStringHeight));
                }

                //рисуем оси Часов и делаем к ним подписи
                gr.DrawString("Время, часы:", font, SystemBrushes.WindowText, new Point(iShiftStart - 110, iStringHeight / 4));
                gr.DrawString("Дата (ФИО)", font, SystemBrushes.WindowText, new Point(10, iStringHeight));
                gr.DrawLine(axis, new Point(0, 0), new Point(iShiftStart, iShiftHeightAll));
                gr.DrawLine(axis, new Point(iShiftStart, 0), new Point(iShiftStart, iShiftHeightAll));

                for (int k = 0; k <= 23; k++)
                {
                    gr.DrawLine(axis, new Point(iShiftStart + k * iMinutesInHour, iShiftHeightAll), new Point(iShiftStart + k * iMinutesInHour, Convert.ToInt32(pictureBox1.Height)));
                    gr.DrawString(Convert.ToString(k), font, SystemBrushes.WindowText, new Point(320 + k * iMinutesInHour, iStringHeight));
                }
                gr.DrawLine(axis, new Point(iShiftStart + 24 * iMinutesInHour, iShiftHeightAll), new Point(iShiftStart + 24 * iMinutesInHour, Convert.ToInt32(pictureBox1.Height)));

                axis.Dispose();
                myBrushAxis = null;
            }

            pictureBox1.Image = bmp;
            pictureBox1.Refresh();
            panelView.Controls.Add(pictureBox1);

            d0 = null; d1 = null; d2 = null; d3 = 0; d4 = 0; d5 = 0; d6 = 0;
            font.Dispose();
            //Отмечаем ФИО как выборанный и устанавливаем на него фокус
            textBoxFIO.Text = sFIO;
            textBoxNav.Text = sNAV;
            StatusLabel2.Text = @"Выбран: " + ShortFIO(sFIO) + @" |  Всего ФИО: " + iFIO;
            if (comboBoxFio.FindString(sFIO) != -1) comboBoxFio.SelectedIndex = comboBoxFio.FindString(sFIO); //ищем в комбобокс выбранный ФИО и устанавливаем на него фокус

            //Write down the selected color
            if (!databasePerson.Exists) return;
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'ProgramSettings' (PoParameterName, PoParameterValue) " +
                                                              " VALUES (@PoParameterName, @PoParameterValue)", sqlConnection))
                    {
                        sqlCommand.Parameters.Add("@PoParameterName", DbType.String).Value = "clrRealRegistration";
                        sqlCommand.Parameters.Add("@PoParameterValue", DbType.String).Value = clrRealRegistration.Name;
                        try { sqlCommand.ExecuteNonQuery(); } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                    }
                }
            }
            sLastSelectedElement = "DrawFullWorkedPeriodRegistration";
            _RefreshPictureBox(pictureBox1, bmp);
            panelViewResize();
        }


        private Bitmap RefreshBitmap(Bitmap b, int nWidth, int nHeight)
        {
            Bitmap result = new Bitmap(nWidth, nHeight);
            using (Graphics g = Graphics.FromImage((Image)result))
                g.DrawImage(b, 0, 0, nWidth, nHeight);
            return result;
        }

        private void ReportsItem_Click(object sender, EventArgs e)
        {
            pictureBox1.Visible = false;
            if (panelView != null && panelView.Controls.Count > 1) panelView.Controls.RemoveAt(1);
            bmp?.Dispose();
            pictureBox1?.Dispose();
            dataGridView1.Visible = true;
            sLastSelectedElement = "dataGridView";
            panelViewResize();
            ReportsItem.Visible = false;
            VisualItem.Visible = true;
            VisualWorkedTimeItem.Visible = true;
        }

        private void panelView_SizeChanged(object sender, EventArgs e)
        { panelViewResize(); }

        private void panelViewResize() //Change PanelView
        {
            int iStringHeight = 19;
            int iShiftHeightAll = 36;
            switch (sLastSelectedElement)
            {
                case "DrawFullWorkedPeriodRegistration":
                    _panelSetHeight(panelView, iShiftHeightAll + iStringHeight * workSelectedDays.Length); //Fixed size of Picture. If need autosize - disable this row
                    break;
                case "DrawRegistration":
                    _panelSetHeight(panelView, iShiftHeightAll + iStringHeight * workSelectedDays.Length); //Fixed size of Picture. If need autosize - disable this row
                    break;
                default:
                    _panelSetAnchor(panelView, (AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top));
                    _panelSetHeight(panelView, _panelParentHeightReturn(panelView) - 120);
                    _panelSetAutoScroll(panelView, true);
                    _panelSetAutoSizeMode(panelView, AutoSizeMode.GrowAndShrink);
                    _panelResume(panelView);
                    break;
            }

            if (_panelControlsCountReturn(panelView) > 1)
            {
                _RefreshPictureBox(pictureBox1, bmp);
            }
        }

        private void dateTimePickerStart_CloseUp(object sender, EventArgs e)
        {
            QuickLoadDataItem.Enabled = true;
            QuickLoadDataItem.BackColor = Color.PaleGreen;
            dateTimePickerEnd.MinDate = DateTime.Parse(dateTimePickerStart.Value.Year + "-" + dateTimePickerStart.Value.Month + "-" + dateTimePickerStart.Value.Day);
        }

        private void dateTimePickerEnd_CloseUp(object sender, EventArgs e)
        { dateTimePickerStart.MaxDate = DateTime.Parse(dateTimePickerEnd.Value.Year + "-" + dateTimePickerEnd.Value.Month + "-" + dateTimePickerEnd.Value.Day); }

        private void BlueItem_Click(object sender, EventArgs e)
        {
            clrRealRegistration = Color.SteelBlue;
            switch (sLastSelectedElement)
            {
                case "DrawFullWorkedPeriodRegistration":
                    DrawFullWorkedPeriodRegistration();
                    break;
                case "DrawRegistration":
                    DrawRegistration();
                    break;
                default:
                    break;
            }
        }

        private void RedItem_Click(object sender, EventArgs e)
        {
            clrRealRegistration = Color.Crimson;
            switch (sLastSelectedElement)
            {
                case "DrawFullWorkedPeriodRegistration":
                    DrawFullWorkedPeriodRegistration();
                    break;
                case "DrawRegistration":
                    DrawRegistration();
                    break;
                default:
                    break;
            }
        }

        private void GreenItem_Click(object sender, EventArgs e)
        {
            clrRealRegistration = Color.ForestGreen;
            switch (sLastSelectedElement)
            {
                case "DrawFullWorkedPeriodRegistration":
                    DrawFullWorkedPeriodRegistration();
                    break;
                case "DrawRegistration":
                    DrawRegistration();
                    break;
                default:
                    break;
            }
        }

        private void YellowItem_Click(object sender, EventArgs e)
        {
            clrRealRegistration = Color.Gold;
            switch (sLastSelectedElement)
            {
                case "DrawFullWorkedPeriodRegistration":
                    DrawFullWorkedPeriodRegistration();
                    break;
                case "DrawRegistration":
                    DrawRegistration();
                    break;
                default:
                    break;
            }
        }



        private void SettingsProgrammItem_Click(object sender, EventArgs e)
        { SettingsProgramm(); }

        private void SettingsProgramm()
        {
            panelViewResize();
            panelView.Visible = false;
            _MenuItemEnabled(QuickLoadDataItem, false);
            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(QuickSettingsItem, false);
            _MenuItemEnabled(AnualDatesMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            VisualItemsAll_Enable(false);
            CheckBoxesFiltersAll_Enable(false);

            labelServer1 = new Label
            {
                Text = "Server",
                BackColor = Color.PaleGreen,
                Location = new Point(20, 60),
                Size = new Size(590, 22),
                BorderStyle = BorderStyle.None,
                TextAlign = ContentAlignment.MiddleLeft,
                Parent = groupBoxProperties
            };
            textBoxServer1 = new TextBox
            {
                Text = sServer1,
                Location = new Point(90, 61),
                Size = new Size(90, 20),
                BorderStyle = BorderStyle.FixedSingle,
                Parent = groupBoxProperties
            };
            toolTip1.SetToolTip(textBoxServer1, "Имя сервера \"Server\" с базой Intellect в виде - NameOfServer.Domain.Subdomain");

            labelServer1UserName = new Label
            {
                Text = "Имя администратора",
                BackColor = Color.PaleGreen,
                Location = new Point(220, 61),
                Size = new Size(70, 20),
                BorderStyle = BorderStyle.None,
                TextAlign = ContentAlignment.MiddleLeft,
                Parent = groupBoxProperties
            };
            textBoxServer1UserName = new TextBox
            {
                Text = sServer1UserName,
                PasswordChar = '*',
                Location = new Point(300, 61),
                Size = new Size(90, 20),
                BorderStyle = BorderStyle.FixedSingle,
                Parent = groupBoxProperties
            };
            toolTip1.SetToolTip(textBoxServer1UserName, "Имя администратора \"sa\" SQL-сервера");

            labelServer1UserPassword = new Label
            {
                Text = "Password",
                BackColor = Color.PaleGreen,
                Location = new Point(420, 61),
                Size = new Size(70, 20),
                BorderStyle = BorderStyle.None,
                TextAlign = ContentAlignment.MiddleLeft,
                Parent = groupBoxProperties
            };
            textBoxServer1UserPassword = new TextBox
            {
                Text = sServer1UserPassword,
                PasswordChar = '*',
                Location = new Point(500, 61),
                Size = new Size(90, 20),
                BorderStyle = BorderStyle.FixedSingle,
                Parent = groupBoxProperties
            };
            toolTip1.SetToolTip(textBoxServer1UserPassword, "Пароль администратора \"sa\" SQL-сервера \"Server\"");

            textBoxServer1.BringToFront();
            textBoxServer1UserName.BringToFront();
            textBoxServer1UserPassword.BringToFront();
            labelServer1UserName.BringToFront();
            labelServer1UserPassword.BringToFront();

            groupBoxProperties.Visible = true;
        }

        private void buttonPropertiesSave_Click(object sender, EventArgs e)
        { PropertiesSave(); }

        private void PropertiesSave()
        {
            sServer1 = textBoxServer1.Text.Trim();
            sServer1UserName = textBoxServer1UserName.Text.Trim();
            sServer1UserPassword = textBoxServer1UserPassword.Text.Trim();

            if (sServer1.Length > 0 && sServer1UserName.Length > 0 && sServer1UserPassword.Length > 0)
            {
                groupBoxProperties.Visible = false;

                if (databasePerson.Exists)
                {
                    using (SQLiteConnection sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                    {
                        sqlConnection.Open();

                        SQLiteCommand sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                        sqlCommand1.ExecuteNonQuery();

                        using (SQLiteCommand sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'EquipmentSettings' (EquipmentParameterName, EquipmentParameterValue, EquipmentParameterServer, Reserv1, Reserv2)" +
                                   " VALUES (@EquipmentParameterName, @EquipmentParameterValue, @EquipmentParameterServer, @Reserv1, @Reserv2)", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@EquipmentParameterName", DbType.String).Value = "Server1UserName";
                            sqlCommand.Parameters.Add("@EquipmentParameterValue", DbType.String).Value = "Server1";
                            sqlCommand.Parameters.Add("@EquipmentParameterServer", DbType.String).Value = sServer1;
                            sqlCommand.Parameters.Add("@Reserv1", DbType.String).Value = EncryptStringToBase64Text(sServer1UserName, btsMess1, btsMess2);
                            sqlCommand.Parameters.Add("@Reserv2", DbType.String).Value = EncryptStringToBase64Text(sServer1UserPassword, btsMess1, btsMess2);
                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }

                        sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                        sqlCommand1.ExecuteNonQuery();
                        sqlCommand1.Dispose();
                    }
                }

                labelServer1.Dispose();
                labelServer1UserName.Dispose();
                labelServer1UserPassword.Dispose();
                textBoxServer1.Dispose();
                textBoxServer1UserName.Dispose();
                textBoxServer1UserPassword.Dispose();

                _MenuItemEnabled(QuickLoadDataItem, true);
                _MenuItemEnabled(FunctionMenuItem, true);
                _MenuItemEnabled(QuickSettingsItem, true);
                _MenuItemEnabled(AnualDatesMenuItem, true);
                _MenuItemEnabled(GroupsMenuItem, true);
                //_MenuItemEnabled(ViewMenuItem, true);
                VisualItemsAll_Enable(true);

                _controlVisible(panelView, true);
            }
            else
            {
                GetInfoSetup();
                _MenuItemEnabled(QuickSettingsItem, true);
            }
        }

        private void buttonPropertiesCancel_Click(object sender, EventArgs e)
        {
            groupBoxProperties.Visible = false;

            labelServer1.Dispose();
            labelServer1UserName.Dispose();
            labelServer1UserPassword.Dispose();
            textBoxServer1.Dispose();
            textBoxServer1UserName.Dispose();
            textBoxServer1UserPassword.Dispose();

            _MenuItemEnabled(QuickLoadDataItem, true);
            _MenuItemEnabled(FunctionMenuItem, true);
            _MenuItemEnabled(QuickSettingsItem, true);
            _MenuItemEnabled(AnualDatesMenuItem, true);
            _MenuItemEnabled(GroupsMenuItem, true);
            //_MenuItemEnabled(FilterItem, true);
            //_MenuItemEnabled(ViewMenuItem, true);
            VisualItemsAll_Enable(true);
            CheckBoxesFiltersAll_Enable(true);

            panelView.Visible = true;
        }

        private void ClearRegistryItem_Click(object sender, EventArgs e)
        { ClearRegistryData(); }

        private void ClearRegistryData()
        {
            try
            {
                using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(myRegKey))
                {
                    EvUserKey.DeleteSubKey("SKDServer");
                    EvUserKey.DeleteSubKey("SKDUser");
                    EvUserKey.DeleteSubKey("SKDUserPassword");
                }
            }
            catch { MessageBox.Show("Ошибки с доступом у реестру на запись. Данные не удалены."); }
        }

        private void btnPropertiesSaveInRegistry_Click(object sender, EventArgs e)
        { SaveDataInRegistry(); }

        private async void SaveDataInRegistry() //Save Parameters into Registry and variables
        {
            string server = textBoxServer1.Text.Trim();
            string user = textBoxServer1UserName.Text.Trim();
            string password = textBoxServer1UserPassword.Text.Trim();
            await Task.Run(() => CheckAliveServer(server, user, password));

            if (bServer1Exist)
            {
                _controlVisible(groupBoxProperties, false);

                sServer1 = server;
                sServer1UserName = user;
                sServer1UserPassword = password;

                try
                {
                    using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(myRegKey))
                    {
                        EvUserKey.SetValue("SKDServer", server, Microsoft.Win32.RegistryValueKind.String);
                        EvUserKey.SetValue("SKDUser", EncryptStringToBase64Text(user, btsMess1, btsMess2), Microsoft.Win32.RegistryValueKind.String);
                        EvUserKey.SetValue("SKDUserPassword", EncryptStringToBase64Text(password, btsMess1, btsMess2), Microsoft.Win32.RegistryValueKind.String);
                    }
                }
                catch { MessageBox.Show("Ошибки с доступом на запись в реестр. Данные сохранены не корректно."); }
                labelServer1.Dispose();
                labelServer1UserName.Dispose();
                labelServer1UserPassword.Dispose();
                textBoxServer1.Dispose();
                textBoxServer1UserName.Dispose();
                textBoxServer1UserPassword.Dispose();

                _MenuItemEnabled(QuickLoadDataItem, true);
                _MenuItemEnabled(FunctionMenuItem, true);
                _MenuItemEnabled(QuickSettingsItem, true);
                _MenuItemEnabled(AnualDatesMenuItem, true);
                _MenuItemEnabled(GroupsMenuItem, true);
                VisualItemsAll_Enable(true);

                panelView.Visible = true;
            }
            else
            {
                GetInfoSetup();
                _MenuItemEnabled(QuickSettingsItem, true);
            }
        }

        private void CheckSavedDataInRegistry() //Read Parameters into Registry and variables
        {
            try
            {
                using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(myRegKey, Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree, System.Security.AccessControl.RegistryRights.ReadKey))
                {
                    sServer1Registry = EvUserKey.GetValue("SKDServer").ToString().Trim();
                    sServer1UserNameRegistry = DecryptBase64ToString(EvUserKey.GetValue("SKDUser").ToString(), btsMess1, btsMess2).Trim();
                    sServer1UserPasswordRegistry = DecryptBase64ToString(EvUserKey.GetValue("SKDUserPassword").ToString(), btsMess1, btsMess2).Trim();
                }
            }
            catch { }
        }

        private void PersonOrGroupItem_Click(object sender, EventArgs e)
        {
            if (PersonOrGroupItem.Text == "Работа с одной персоной")
            {
                PersonOrGroupItem.Text = "Работа с группой";
                nameOfLastTableFromDB = "PersonGroup";
                textBoxGroup.Text = "";
                textBoxGroupDescription.Text = "";
            }
            else
            {
                PersonOrGroupItem.Text = "Работа с одной персоной";
                nameOfLastTableFromDB = "PersonRegistered";
            }
        }

        private void SetupItem_Click(object sender, EventArgs e)
        { GetInfoSetup(); }

        private void GetInfoSetup()
        {
            DialogResult result = MessageBox.Show(
                @"Перед получением информации необходимо в Настройках:" + "\n\n" +
                 "1. Добавить имя сервера Интеллект (СКД - SERVER.DOMAIN), а также имя и пароль разрешенного пользователя для данного SQL-сервера СКД\n" +
                 "2. Сохранить данные параметры\n" +
                 "3. После этого можно получать списки пользователей проходивших пукты регистрации, " +
                 "просматривать данные по регистрациям и проводить анализ.\n\nДата и время локального ПК: " +
                _dateTimePickerReturn(dateTimePickerEnd),
                //dateTimePickerEnd.Value,
                @"Информация о программе",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1);
        }

        private void textBoxGroup_TextChanged(object sender, EventArgs e)
        {
            if (textBoxGroup.Text.Trim().Length > 0)
            {
                AddPersonToGroupItem.Enabled = true;
                CreateGroupItem.Enabled = true;
                if (textBoxGroupDescription.Text.Trim().Length > 0)
                {
                    StatusLabel2.Text = @"Создать группу: " + textBoxGroup.Text.Trim().ToString() + "(" + textBoxGroupDescription.Text.Trim() + ")";
                }
                else
                {
                    StatusLabel2.Text = @"Создать группу: " + textBoxGroup.Text.Trim().ToString();
                }
            }
            else
            {
                AddPersonToGroupItem.Enabled = false;
                CreateGroupItem.Enabled = false;
                StatusLabel2.Text = @"Всего ФИО: " + iFIO;
            }
        }


        /// //////////////// Start  DatagridView functions

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Rows.Count > 0 && dataGridView1.CurrentRow.Index < dataGridView1.Rows.Count)
                try
                {
                    int IndexCurrentRow = _dataGridView1CurrentRowIndex();
                    if (IndexCurrentRow > -1)
                    {
                        if (nameOfLastTableFromDB == "PersonGroupDesciption")
                        {
                            int IndexColumn1 = 0;           // индекс 1-й колонки в датагрид
                            int IndexColumn2 = 0;           // индекс 2-й колонки в датагрид

                            for (int i = 0; i < dataGridView1.ColumnCount; i++)
                            {
                                switch (dataGridView1.Columns[i].HeaderText)
                                {
                                    case "GroupPerson":
                                        IndexColumn1 = i;
                                        break;
                                    case "GroupPersonDescription":
                                        IndexColumn2 = i;
                                        break;
                                    default:
                                        break;
                                }
                            }
                            textBoxGroup.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString(); //Take the name of selected group
                            textBoxGroupDescription.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn2].Value.ToString(); //Take the name of selected group
                            groupBoxPeriod.BackColor = Color.PaleGreen;
                            groupBoxRemoveDays.BackColor = SystemColors.Control;
                            StatusLabel2.Text = @"Выбрана группа: " + textBoxGroup.Text + @" |  Всего ФИО: " + iFIO;
                            if (textBoxFIO.TextLength > 3)
                            {
                                comboBoxFio.SelectedIndex = comboBoxFio.FindString(textBoxFIO.Text);
                            }
                            nameOfLastTableFromDB = "PersonGroupDesciption";
                        }

                        else if (nameOfLastTableFromDB == "PersonGroup")
                        {
                            // comboBoxFio.Items.Clear();
                            int IndexColumn1 = -1;           // индекс 1-й колонки в датагрид
                            int IndexColumn2 = -1;           // индекс 2-й колонки в датагрид
                            int IndexColumn3 = -1;           // индекс 3-й колонки в датагрид
                            int IndexColumn4 = -1;           // индекс 4-й колонки в датагрид

                            for (int i = 0; i < dataGridView1.ColumnCount; i++)
                            {
                                if (dataGridView1.Columns[i].HeaderText.Trim() == "Фамилия Имя Отчество" ||
                                        dataGridView1.Columns[i].HeaderText.Trim() == "FIO")
                                    IndexColumn1 = i;
                                if (dataGridView1.Columns[i].HeaderText.Trim() == "NAV-код" ||
                                        dataGridView1.Columns[i].HeaderText.Trim() == "NAV")
                                    IndexColumn2 = i;
                                if (dataGridView1.Columns[i].HeaderText.Trim() == "Контрольное время, часы" ||
                                        dataGridView1.Columns[i].HeaderText.Trim() == "HourControlling")
                                    IndexColumn3 = i;
                                if (dataGridView1.Columns[i].HeaderText.Trim() == "Контрольное время, минуты" ||
                                        dataGridView1.Columns[i].HeaderText.Trim() == "MinuteControlling")
                                    IndexColumn4 = i;
                            }

                            textBoxFIO.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString();
                            textBoxNav.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn2].Value.ToString(); //Take the name of selected group
                            StatusLabel2.Text = @"Выбрана группа: " + textBoxGroup.Text + @" | Курсор на: " + textBoxFIO.Text;
                            groupBoxPeriod.BackColor = Color.PaleGreen;
                            numUpDownHour.Value = Convert.ToInt32(dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn3].Value.ToString());
                            numUpDownMinute.Value = Convert.ToInt32(dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn4].Value.ToString());
                            if (textBoxFIO.TextLength > 3)
                            {
                                comboBoxFio.SelectedIndex = comboBoxFio.FindString(textBoxFIO.Text);
                            }
                            nameOfLastTableFromDB = "PersonGroup";
                        }

                        if (nameOfLastTableFromDB == "PersonRegistered")
                        {
                            int IndexColumn1 = 0;           // индекс 1-й колонки в датагрид
                            int IndexColumn2 = 1;           // индекс 2-й колонки в датагрид

                            for (int i = 0; i < dataGridView1.ColumnCount; i++)
                            {
                                switch (dataGridView1.Columns[i].HeaderText)
                                {
                                    case "FIO":
                                        IndexColumn1 = i;
                                        break;
                                    case "Фамилия Имя Отчество":
                                        IndexColumn1 = i;
                                        break;
                                    case "NAV-код":
                                        IndexColumn2 = i;
                                        break;
                                    case "NAV":
                                        IndexColumn2 = i;
                                        break;
                                    default:
                                        break;
                                }
                            }

                            textBoxNav.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn2].Value.ToString(); //Take the name of selected group
                            textBoxFIO.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString();
                            StatusLabel2.Text = @"Выбран: " + textBoxFIO.Text + @" |  Всего ФИО: " + iFIO;
                            groupBoxPeriod.BackColor = Color.PaleGreen;
                            groupBoxTime.BackColor = Color.PaleGreen;
                            groupBoxRemoveDays.BackColor = SystemColors.Control;
                            numUpDownHour.Value = 9;
                            numUpDownMinute.Value = 0;
                            if (textBoxFIO.TextLength > 3)
                            {
                                comboBoxFio.SelectedIndex = comboBoxFio.FindString(textBoxFIO.Text);
                            }
                            //                    nameOfLastTableFromDB = "PersonRegistered";
                        }
                    }
                }
                catch (Exception expt)
                {
                    MessageBox.Show(expt.ToString());
                }
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0 && dataGridView1.CurrentRow.Index < dataGridView1.Rows.Count)
            {
                if (nameOfLastTableFromDB == "PersonGroupDesciption")
                {
                    bErrorData = false;
                    MembersGroup();
                }
                else if (nameOfLastTableFromDB == "PersonGroup")
                {
                    int IndexColumn1 = -1;
                    int IndexColumn2 = -1;
                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        if (dataGridView1.Columns[i].HeaderText.ToString() == "Контрольное время, часы")
                            IndexColumn1 = i;
                        else if (dataGridView1.Columns[i].HeaderText.ToString() == "Контрольное время, минуты")
                            IndexColumn2 = i;
                    }
                    if (IndexColumn1 > -1 || IndexColumn2 > -1)
                        UpdateControllingItem.Visible = true;
                    else
                        UpdateControllingItem.Visible = false;
                }
            }
        }

        private void UpdateControllingItem_Click(object sender, EventArgs e)
        {
            if (nameOfLastTableFromDB == "PersonGroup")
            {
                int IndexCurrentRow = _dataGridView1CurrentRowIndex();

                int IndexColumn1 = -1;
                int IndexColumn2 = -1;
                int IndexColumn3 = -1;
                int IndexColumn4 = -1;
                int IndexColumn5 = -1;

                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    if (dataGridView1.Columns[i].HeaderText.ToString() == "Фамилия Имя Отчество")
                        IndexColumn1 = i;
                    else if (dataGridView1.Columns[i].HeaderText.ToString() == "NAV-код")
                        IndexColumn2 = i;
                    else if (dataGridView1.Columns[i].HeaderText.ToString() == "Контрольное время, часы")
                        IndexColumn3 = i;
                    else if (dataGridView1.Columns[i].HeaderText.ToString() == "Контрольное время, минуты")
                        IndexColumn4 = i;
                    else if (dataGridView1.Columns[i].HeaderText.ToString() == "Группа")
                        IndexColumn5 = i;
                }

                textBoxFIO.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString();
                textBoxNav.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn2].Value.ToString(); //Take the name of selected group
                textBoxGroup.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn5].Value.ToString(); //Take the name of selected group

                decimal hourControlStart = Convert.ToDecimal(dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn3].Value.ToString());
                decimal minuteControlStart = Convert.ToDecimal(dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn4].Value.ToString());
                decimal controlStartHours = hourControlStart + (minuteControlStart + 1) / 60 - (1 / 60);
                MessageBox.Show(hourControlStart + "\n" + minuteControlStart + "\n" + textBoxFIO + "\n" + numUpDownHour.Value + "\n" + numUpDownMinute.Value);

                using (var connection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    connection.Open();
                    using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PersonGroup' (FIO, NAV, GroupPerson, HourControlling, MinuteControlling, Controlling) " +
                                            "VALUES (@FIO, @NAV, @GroupPerson, @HourControlling, @MinuteControlling, @Controlling)", connection))
                    {
                        command.Parameters.Add("@FIO", DbType.String).Value = textBoxFIO.Text;
                        command.Parameters.Add("@NAV", DbType.String).Value = textBoxNav.Text;
                        command.Parameters.Add("@GroupPerson", DbType.String).Value = textBoxGroup.Text;
                        command.Parameters.Add("@HourControlling", DbType.String).Value = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn3].Value.ToString();
                        command.Parameters.Add("@MinuteControlling", DbType.String).Value = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn4].Value.ToString();
                        command.Parameters.Add("@Controlling", DbType.Decimal).Value = controlStartHours;
                        try { command.ExecuteNonQuery(); } catch { }
                    }
                }

                ShowDataTableQuery(databasePerson, "PersonGroup",
                  "SELECT FIO AS 'Фамилия Имя Отчество', NAV AS 'NAV-код', GroupPerson AS 'Группа'," +
                  " HourControlling AS 'Контрольное время, часы', MinuteControlling AS 'Контрольное время, минуты' ",
                  " Where GroupPerson like '" + textBoxGroup.Text + "' ORDER BY FIO");
                StatusLabel2.Text = @"Обновлены данные " + textBoxFIO.Text + " в группе: " + textBoxGroup.Text;
            }
            //UpdateControllingItem.Visible = false;
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (nameOfLastTableFromDB == "PersonGroup")
            {
                int IndexCurrentRow = _dataGridView1CurrentRowIndex();

                int IndexColumn1 = -1;
                int IndexColumn2 = -1;
                int IndexColumn3 = -1;
                int IndexColumn4 = -1;
                int IndexColumn5 = -1;

                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    if (dataGridView1.Columns[i].HeaderText.ToString() == "Фамилия Имя Отчество")
                        IndexColumn1 = i;
                    else if (dataGridView1.Columns[i].HeaderText.ToString() == "NAV-код")
                        IndexColumn2 = i;
                    else if (dataGridView1.Columns[i].HeaderText.ToString() == "Контрольное время, часы")
                        IndexColumn3 = i;
                    else if (dataGridView1.Columns[i].HeaderText.ToString() == "Контрольное время, минуты")
                        IndexColumn4 = i;
                    else if (dataGridView1.Columns[i].HeaderText.ToString() == "Группа")
                        IndexColumn5 = i;
                }

                textBoxFIO.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString();
                textBoxNav.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn2].Value.ToString(); //Take the name of selected group
                textBoxGroup.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn5].Value.ToString(); //Take the name of selected group

                decimal hourControlStart = Convert.ToDecimal(dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn3].Value.ToString());
                decimal minuteControlStart = Convert.ToDecimal(dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn4].Value.ToString());
                decimal controlStartHours = hourControlStart + (minuteControlStart + 1) / 60 - (1 / 60);

                using (var connection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    connection.Open();
                    using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PersonGroup' (FIO, NAV, GroupPerson, HourControlling, MinuteControlling, Controlling) " +
                                            "VALUES (@FIO, @NAV, @GroupPerson, @HourControlling, @MinuteControlling, @Controlling)", connection))
                    {
                        command.Parameters.Add("@FIO", DbType.String).Value = textBoxFIO.Text;
                        command.Parameters.Add("@NAV", DbType.String).Value = textBoxNav.Text;
                        command.Parameters.Add("@GroupPerson", DbType.String).Value = textBoxGroup.Text;
                        command.Parameters.Add("@HourControlling", DbType.String).Value = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn3].Value.ToString();
                        command.Parameters.Add("@MinuteControlling", DbType.String).Value = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn4].Value.ToString();
                        command.Parameters.Add("@Controlling", DbType.Decimal).Value = controlStartHours;
                        try { command.ExecuteNonQuery(); } catch { }
                    }
                }

                ShowDataTableQuery(databasePerson, "PersonGroup",
                  "SELECT FIO AS 'Фамилия Имя Отчество', NAV AS 'NAV-код', GroupPerson AS 'Группа'," +
                  " HourControlling AS 'Контрольное время, часы', MinuteControlling AS 'Контрольное время, минуты' ",
                  " Where GroupPerson like '" + textBoxGroup.Text + "' ORDER BY FIO");
                StatusLabel2.Text = @"Обновлено время прихода " + ShortFIO(textBoxFIO.Text) + " в группе: " + textBoxGroup.Text;
            }
        }

        //Show help to Edit on some collumns DataGridView
        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (nameOfLastTableFromDB == "PersonGroup")
            {
                DataGridViewCell cell;
                if ((e.ColumnIndex == this.dataGridView1.Columns["Контрольное время, часы"].Index) && e.Value != null)
                {
                    cell = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    cell.ToolTipText = "Для установки нового значения нажмите F2,\nвнесите новое значение,\nа затем нажмите Enter";
                }
                else if ((e.ColumnIndex == this.dataGridView1.Columns["Контрольное время, минуты"].Index) && e.Value != null)
                {
                    cell = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    cell.ToolTipText = "Для установки нового значения нажмите F2,\nвнесите новое значение,\nа затем нажмите Enter";
                }
                else
                {
                    cell = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    cell.ToolTipText = cell.Value.ToString();
                }
            }
        }

        /// //////////////// End  DatagridView functions



        //Color Control elements of Person depending of the selected MenuItem  
        private void CreateGroupItem_MouseHover(object sender, EventArgs e)
        {  //Save previous color          
            labelGroupCurrentBackColor = labelGroup.BackColor;
            textBoxGroupCurrentBackColor = textBoxGroup.BackColor;
            labelGroupDescriptionCurrentBackColor = labelGroupDescription.BackColor;
            textBoxGroupDescriptionCurrentBackColor = textBoxGroupDescription.BackColor;

            //set Over Color
            labelGroup.BackColor = Color.PaleGreen;
            textBoxGroup.BackColor = Color.PaleGreen;
            labelGroupDescription.BackColor = Color.PaleGreen;
            textBoxGroupDescription.BackColor = Color.PaleGreen;
        }

        private void CreateGroupItem_MouseLeave(object sender, EventArgs e)
        {   //Restore saved color
            labelGroup.BackColor = labelGroupCurrentBackColor;
            textBoxGroup.BackColor = textBoxGroupCurrentBackColor;
            labelGroupDescription.BackColor = labelGroupDescriptionCurrentBackColor;
            textBoxGroupDescription.BackColor = textBoxGroupDescriptionCurrentBackColor;
        }

        private void PersonOrGroupItem_MouseEnter(object sender, EventArgs e)
        {
            if (PersonOrGroupItem.Text == "Работа с одной персоной")
            {  //Save previous color              
                comboBoxFioCurrentBackColor = comboBoxFio.BackColor;
                textBoxFIOCurrentBackColor = textBoxFIO.BackColor;
                textBoxNavCurrentBackColor = textBoxNav.BackColor;

                //set Over Color
                comboBoxFio.BackColor = Color.PaleGreen;
                textBoxFIO.BackColor = Color.PaleGreen;
                textBoxNav.BackColor = Color.PaleGreen;
            }
            else
            {  //Save previous color              
                labelGroupCurrentBackColor = labelGroup.BackColor;
                textBoxGroupCurrentBackColor = textBoxGroup.BackColor;

                //set Over Color
                labelGroup.BackColor = Color.PaleGreen;
                textBoxGroup.BackColor = Color.PaleGreen;
            }
        }

        private void PersonOrGroupItem_MouseLeave(object sender, EventArgs e)
        {
            if (PersonOrGroupItem.Text == "Работа с одной персоной")
            {   //Restore saved color             
                comboBoxFio.BackColor = comboBoxFioCurrentBackColor;
                textBoxFIO.BackColor = textBoxFIOCurrentBackColor;
                textBoxNav.BackColor = textBoxNavCurrentBackColor;
            }
            else
            {  //Restore saved color              
                labelGroup.BackColor = labelGroupCurrentBackColor;
                textBoxGroup.BackColor = textBoxGroupCurrentBackColor;
            }
        }

        private void textBoxGroupDescription_TextChanged(object sender, EventArgs e)
        {
            if (textBoxGroupDescription.Text.Trim().Length > 0)
            {
                StatusLabel2.Text = @"Создать группу: " + textBoxGroup.Text.Trim().ToString() + "(" + textBoxGroupDescription.Text.Trim() + ")";
            }
            else
            {
                StatusLabel2.Text = @"Создать группу: " + textBoxGroup.Text.Trim().ToString();
            }
        }



        //Start of the Block Encryption-Decryption
        private static string EncryptStringToBase64Text(string plainText, byte[] Key, byte[] IV) //Encrypt variables PlainText Data
        {
            string sBase64Test;
            sBase64Test = Convert.ToBase64String(EncryptStringToBytes(plainText, Key, IV));
            return sBase64Test;
        }

        private static byte[] EncryptStringToBytes(string plainText, byte[] Key, byte[] IV) //Encrypt variables PlainText Data
        {
            // Check arguments. 
            if (plainText == null || plainText.Length <= 0)
                throw new ArgumentNullException("plainText");
            if (Key == null || Key.Length <= 0)
                throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0)
                throw new ArgumentNullException("IV");
            byte[] encrypted;
            // Create an RijndaelManaged object with the specified key and IV. 
            using (RijndaelManaged rijAlg = new RijndaelManaged())
            {
                rijAlg.Key = Key;
                rijAlg.IV = IV;

                // Create a decrytor to perform the stream transform.
                ICryptoTransform encryptor = rijAlg.CreateEncryptor(rijAlg.Key, rijAlg.IV);

                // Create the streams used for encryption. 
                using (System.IO.MemoryStream msEncrypt = new System.IO.MemoryStream())
                {
                    using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                    {
                        using (System.IO.StreamWriter swEncrypt = new System.IO.StreamWriter(csEncrypt))
                        {

                            //Write all data to the stream.
                            swEncrypt.Write(plainText);
                        }
                        encrypted = msEncrypt.ToArray();
                    }
                }
            }
            // Return the encrypted bytes from the memory stream. 
            return encrypted;
        }

        private static string DecryptBase64ToString(string sBase64Text, byte[] Key, byte[] IV) //Encrypt variables PlainText Data
        {
            byte[] bBase64Test;
            bBase64Test = Convert.FromBase64String(sBase64Text);
            return DecryptStringFromBytes(bBase64Test, Key, IV);
        }

        private static string DecryptStringFromBytes(byte[] cipherText, byte[] Key, byte[] IV) //Decrypt PlainText Data to variables
        {
            // Check arguments. 
            if (cipherText == null || cipherText.Length <= 0)
                throw new ArgumentNullException("cipherText");
            if (Key == null || Key.Length <= 0)
                throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0)
                throw new ArgumentNullException("IV");

            // Declare the string used to hold the decrypted text.
            string plaintext = null;

            // Create an RijndaelManaged object  with the specified key and IV.
            using (RijndaelManaged rijAlg = new RijndaelManaged())
            {
                rijAlg.Key = Key;
                rijAlg.IV = IV;

                // Create a decrytor to perform the stream transform.
                ICryptoTransform decryptor = rijAlg.CreateDecryptor(rijAlg.Key, rijAlg.IV);

                // Create the streams used for decryption. 
                using (System.IO.MemoryStream msDecrypt = new System.IO.MemoryStream(cipherText))
                {
                    using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                    {
                        using (System.IO.StreamReader srDecrypt = new System.IO.StreamReader(csDecrypt))
                        {
                            // Read the decrypted bytes from the decrypting stream and place them in a string. 
                            plaintext = srDecrypt.ReadToEnd();
                        }
                    }
                }
            }

            return plaintext;
        }

        private void TestCryptionItem_Click(object sender, EventArgs e)
        {
            string original = "Here is some data to encrypt!";
            MessageBox.Show("Original:   " + original);

            using (RijndaelManaged myRijndael = new RijndaelManaged())
            {
                myRijndael.Key = btsMess1;
                myRijndael.IV = btsMess2;
                // Encrypt the string to an array of bytes.
                byte[] encrypted = EncryptStringToBytes(original, myRijndael.Key, myRijndael.IV);

                StringBuilder s = new StringBuilder();
                foreach (byte item in encrypted)
                {
                    s.Append(item.ToString("X2") + " ");
                }
                MessageBox.Show("Encrypted:   " + s);

                // Decrypt the bytes to a string.
                string decrypted = DecryptStringFromBytes(encrypted, btsMess1, btsMess2);

                //Display the original data and the decrypted data.
                MessageBox.Show("Decrypted:    " + decrypted);
            }
        }
        //End of the Block Encryption-Decryption

    }

    public class Person
    {
        public int id;
        public int idCard = 0;
        public string FIO;
        public string NAV = "";
        public string Department = "";
        public string GroupPerson = "Office";
        public string HourControlling = "9";
        public string MinuteControlling = "0";
        public decimal Controlling = 9;
        public string HourControllingOut = "18";
        public string MinuteControllingOut = "0";
        public decimal ControllingOut = 18;

        public string RealInHour = "9";
        public string RealInMinute = "0";
        public decimal RealInDecimal = 9;
        public string RealHourOut = "18";
        public string RealOutMinute = "0";
        public decimal RealOutDecimal = 18;
        public decimal RealWorkedDayHoursDecimal = 9;
        public string RealInHHMM = "09:00";
        public string RealOutHHMM = "18:00";
        public string RealWorkedDayHoursHHMM = "09:00";
        public string RealDate = "";

        public string serverSKD = "";
        public string namePassPoint = "";
        public string directionPass = "";

    }

    public class EncryptDecrypt
    {
        /*
            string plainText = "Hello, World!";            
            string TestKey = "gFlMfLZu4unBGfBh9weIuLwlcXHSa59vxyUQWM3yh1M=";
            string TestIV = "rRNgGypocle9VG9bth6kxg==";

            var ed = new EncDec(TestKey, TestIV);
            var cypherText = ed.Encrypt(plainText);  //4U8SPwju5aH9pHwmiYd1jQ==
            var plainText2 = ed.Decrypt(cypherText); //4U8SPwju5aH9pHwmiYd1jQ==
*/
        private byte[] key;
        private byte[] IV;

        public void EncDec(string keyText, string ivText)
        {
            key = Convert.FromBase64String(keyText);
            IV = Convert.FromBase64String(ivText);
        }

        public void EncDecKeys(byte[] key, byte[] IV)
        {
            this.key = key;
            this.IV = IV;
        }

        public string Encrypt(string plainText)
        {
            var bytes = EncryptStringToBytes(plainText, key, IV);
            return Convert.ToBase64String(bytes);
        }

        private static byte[] EncryptStringToBytes(string plainText, byte[] Key, byte[] IV)
        {
            // Check arguments. 
            if (plainText == null || plainText.Length <= 0)
                throw new ArgumentNullException("plainText");
            if (Key == null || Key.Length <= 0)
                throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0)
                throw new ArgumentNullException("IV");
            byte[] encrypted;
            // Create an RijndaelManaged object 
            // with the specified key and IV. 
            using (RijndaelManaged rijAlg = new RijndaelManaged())
            {
                rijAlg.Key = Key;
                rijAlg.IV = IV;

                // Create a decrytor to perform the stream transform.
                ICryptoTransform encryptor = rijAlg.CreateEncryptor(rijAlg.Key, rijAlg.IV);

                // Create the streams used for encryption. 
                using (System.IO.MemoryStream msEncrypt = new System.IO.MemoryStream())
                {
                    using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                    {
                        using (System.IO.StreamWriter swEncrypt = new System.IO.StreamWriter(csEncrypt))
                        {
                            swEncrypt.Write(plainText);  //Write all data to the stream.
                        }
                        encrypted = msEncrypt.ToArray();
                    }
                }
            }

            // Return the encrypted bytes from the memory stream. 
            return encrypted;
        }

        public string Decrypt(string cipherText)
        {
            var bytes = Convert.FromBase64String(cipherText);
            return DecryptStringFromBytes(bytes, key, IV);
        }

        private static string DecryptStringFromBytes(byte[] cipherText, byte[] Key, byte[] IV)
        {
            // Check arguments. 
            if (cipherText == null || cipherText.Length <= 0)
                throw new ArgumentNullException("cipherText");
            if (Key == null || Key.Length <= 0)
                throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0)
                throw new ArgumentNullException("IV");

            // Declare the string used to hold  the decrypted text. 
            string plaintext = null;

            // Create an RijndaelManaged object with the specified key and IV. 
            using (RijndaelManaged rijAlg = new RijndaelManaged())
            {
                rijAlg.Key = Key;
                rijAlg.IV = IV;

                // Create a decrytor to perform the stream transform.
                ICryptoTransform decryptor = rijAlg.CreateDecryptor(rijAlg.Key, rijAlg.IV);

                // Create the streams used for decryption. 
                using (System.IO.MemoryStream msDecrypt = new System.IO.MemoryStream(cipherText))
                {
                    using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                    {
                        using (System.IO.StreamReader srDecrypt = new System.IO.StreamReader(csDecrypt))
                        {

                            // Read the decrypted bytes from the decrypting stream 
                            // and place them in a string.
                            plaintext = srDecrypt.ReadToEnd();
                        }
                    }
                }
            }

            return plaintext;
        }
    }
}

