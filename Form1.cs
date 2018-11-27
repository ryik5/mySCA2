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
    public partial class FormPersonViewerSCA :Form
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
        private string sServer1 = "";
        private string sServer1UserName = "";
        private string sServer1UserPassword = "";
        private readonly byte[] btsMess1 = Convert.FromBase64String(@"OCvesunvXXsxtt381jr7vp3+UCwDbE4ebdiL1uinGi0="); //Key Encrypt
        private readonly byte[] btsMess2 = Convert.FromBase64String(@"NO6GC6Zjl934Eh8MAJWuKQ=="); //Key Decrypt

        public FormPersonViewerSCA()
        { InitializeComponent(); }

        private void Form1_Load(object sender, EventArgs e)
        { Form1Load(); }

        private void Form1Load()
        {
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
            dateTimePickerEnd.MaxDate = DateTime.Parse("2018-12-01");
            dateTimePickerStart.Value = DateTime.Parse(DateTime.Now.Year + "-" + DateTime.Now.Month + "-01");
            dateTimePickerEnd.Value = DateTime.Now;
            numUpDownHour.Value = 9;
            numUpDownMinute.Value = 0;
            PersonOrGroupItem.Text = "Работа с одной персоной";
            StatusLabel2.Text = "";

            MakeDB();
            SetTechInfoIntoDB();
            BoldAnualDates();
            WriteLastParametersIntoVariable();
            CheckForIllegalCrossThreadCalls = false;
            AddAnualDateItem.Enabled = false;
            DeleteAnualDateItem.Enabled = false;
            EnterEditAnualItem.Enabled = true;

            MembersGroupItem.Enabled = false;
            AddPersonToGroupItem.Enabled = false;
            CreateGroupItem.Visible = false;
            DeleteGroupItem.Visible = false;
            DeletePersonFromGroupItem.Visible = false;
            QuickFilterItem.Enabled = false;
            UpdateControllingItem.Visible = false;
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
            } catch { }
            try { textBoxNav.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn2].Value.ToString(); } catch { }

            numUpDownHour.Value = 9;
            numUpDownMinute.Value = 0;

            try
            {
                comboBoxFio.SelectedIndex = comboBoxFio.FindString(sFIO); //ищем в комбобокс выбранный ФИО и устанавливаем на него фокус
                if (comboBoxFio.FindString(sFIO) != -1 && ShortFIO(sFIO).Length > 3)
                {
                    StatusLabel2.Text = @"Выбран: " + ShortFIO(sFIO) + @" |  Всего ФИО: " + iFIO;
                }
                else if (ShortFIO(sFIO).Length < 3 && iFIO > 0)
                { StatusLabel2.Text = @"Всего ФИО: " + iFIO; }
            } catch { StatusLabel2.Text = " Начните работу с кнопки - \"Получить ФИО\""; }

        }


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

        private void MakeDB()
        {
            MakeDBAndExecuteSql("CREATE TABLE IF NOT EXISTS 'PersonRegisteredFull' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, FIO TEXT, NAV TEXT, iDCard TEXT, DateRegistered TEXT, " +
                    "HourComming TEXT, MinuteComming TEXT, Comming REAL, HourControlling TEXT, MinuteControlling TEXT, Controlling REAL, ServerOfRegistration TEXT, Reserv1 TEXT, Reserv2 TEXT);", databasePerson);
            MakeDBAndExecuteSql("CREATE TABLE IF NOT EXISTS 'PersonRegistered' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, FIO TEXT, NAV TEXT, iDCard TEXT, DateRegistered TEXT, " +
                    "HourComming TEXT, MinuteComming TEXT, Comming REAL, HourControlling TEXT, MinuteControlling TEXT, Controlling REAL, ServerOfRegistration TEXT, Reserv1 TEXT, Reserv2 TEXT);", databasePerson);
            MakeDBAndExecuteSql("CREATE TABLE IF NOT EXISTS 'PersonTemp' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, FIO TEXT, NAV TEXT, iDCard TEXT, DateRegistered TEXT, " +
                    "HourComming TEXT, MinuteComming TEXT, Comming REAL, HourControlling TEXT, MinuteControlling TEXT, Controlling REAL, ServerOfRegistration TEXT, Reserv1 TEXT, Reserv2 TEXT);", databasePerson);
            MakeDBAndExecuteSql("CREATE TABLE IF NOT EXISTS 'PersonGroup' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, FIO TEXT, NAV TEXT, GroupPerson TEXT, " +
                    "HourControlling TEXT, MinuteControlling TEXT, Controlling REAL, Reserv1 TEXT, Reserv2 TEXT, UNIQUE ('FIO', 'NAV', 'GroupPerson') ON CONFLICT REPLACE);", databasePerson);
            MakeDBAndExecuteSql("CREATE TABLE IF NOT EXISTS 'PersonGroupDesciption' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, GroupPerson TEXT, GroupPersonDescription TEXT, Reserv1 TEXT, Reserv2 TEXT, " +
                    "UNIQUE ('GroupPerson') ON CONFLICT REPLACE);", databasePerson);
            MakeDBAndExecuteSql("CREATE TABLE IF NOT EXISTS 'TechnicalInfo' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, PCName TEXT, POName TEXT, POVersion TEXT, LastDateStarted TEXT, Reserv1 TEXT, " +
                    "Reserv2 TEXT, Reserverd3 TEXT);", databasePerson);
            MakeDBAndExecuteSql("CREATE TABLE IF NOT EXISTS 'BoldedDates' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, BoldedDate TEXT, NAV TEXT, Groups TEXT, Reserv1 TEXT, Reserv2 TEXT);", databasePerson);
            MakeDBAndExecuteSql("CREATE TABLE IF NOT EXISTS 'MySettings' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, MyParameterName TEXT, MyParameterValue TEXT, Reserv1 TEXT, Reserv2 TEXT);", databasePerson);
            MakeDBAndExecuteSql("CREATE TABLE IF NOT EXISTS 'ProgramSettings' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, PoParameterName TEXT, PoParameterValue TEXT, Reserv1 TEXT, Reserv2 TEXT, " +
                   "UNIQUE (PoParameterName) ON CONFLICT REPLACE);", databasePerson);
            MakeDBAndExecuteSql("CREATE TABLE IF NOT EXISTS 'EquipmentSettings' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, EquipmentParameterName TEXT, EquipmentParameterValue TEXT, EquipmentParameterServer TEXT," +
                    "Reserv1, Reserv2, UNIQUE ('EquipmentParameterName', 'EquipmentParameterServer') ON CONFLICT REPLACE);", databasePerson);
            MakeDBAndExecuteSql("CREATE TABLE IF NOT EXISTS 'PersonsLastList' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, PersonsList TEXT, " +
                    "Reserv1 TEXT, Reserv2 TEXT, UNIQUE ('PersonsList', Reserv1) ON CONFLICT REPLACE);", databasePerson);
            MakeDBAndExecuteSql("CREATE TABLE IF NOT EXISTS 'PersonsLastComboList' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, ComboList TEXT, " +
                    "Reserv1 TEXT, Reserv2 TEXT, UNIQUE ('ComboList', Reserv1) ON CONFLICT REPLACE);", databasePerson);
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
                                } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
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
                                } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
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
                                } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
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
                                            sServer1 = record["EquipmentParameterServer"].ToString();
                                            sServer1UserName = DecryptBase64ToString(record["Reserv1"].ToString(), btsMess1, btsMess2);
                                            sServer1UserPassword = DecryptBase64ToString(record["Reserv2"].ToString(), btsMess1, btsMess2);
                                        }
                                    }
                                } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                            }
                        }
                    }

                }
            }
            iFIO = iCombo;
            try { comboBoxFio.SelectedIndex = 0; } catch { }
        }

        private void MakeDBAndExecuteSql(string SqlQuery, System.IO.FileInfo FileDB) //Prepare DB and execute of SQL Query
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
                } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
            }
            SqlQuery = null;
        }

        private void ShowDataTableQuery(System.IO.FileInfo databasePerson, string myTable, string mySqlQuery = "SELECT DISTINCT FIO AS 'Фамилия Имя Отчество', NAV AS 'NAV-код', " +
            " DateRegistered AS 'Дата регистрации', HourComming AS 'Время прихода, часы',  MinuteComming AS 'Время прихода, минуты', ServerOfRegistration AS 'Сервер', " +
            " HourControlling AS 'Контрольное время, часы', MinuteControlling AS 'Контрольное время, минуты', Reserv1 AS 'Точка прохода', Reserv2 AS 'Направление'",
            string mySqlWhere = "ORDER BY FIO, DateRegistered, Comming") //Query data from the Table of the DB
        {
            if (databasePerson.Exists)
            {
                iCounterLine = 0;
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    using (var sqlDA = new SQLiteDataAdapter(mySqlQuery + " FROM '" + myTable + "' " + mySqlWhere + "; ", sqlConnection))
                    {
                        var dt = new DataTable();
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
                                } catch { }
                            }
                        }
                    }
                }
            }
            sLastSelectedElement = "dataGridView";
        }

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
                        using (var sqlCommand = new SQLiteCommand("DELETE FROM '" + myTable + "' Where NAV like '" + _textBoxNavText() + "' AND " + mySqlParameter1 + "= @" + mySqlParameter1 + ";", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@" + mySqlParameter1, DbType.String).Value = mySqlData1;
                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }
                    }
                    if (mySqlParameter1.Trim().Length > 0 && mySqlParameter2.Trim().Length > 0)
                    {
                        using (var sqlCommand = new SQLiteCommand("DELETE FROM '" + myTable + "' Where NAV like '" + _textBoxNavText() + "' AND " + mySqlParameter1 + "= @" + mySqlParameter1 +
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
                        using (var sqlCommand = new SQLiteCommand("DELETE FROM '" + myTable + "' Where NAV like '" + _textBoxNavText() + "' AND " + mySqlParameter2 + " < @" + mySqlParameter2 + ";", sqlConnection))
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

        private async void buttonGetFio_Click(object sender, EventArgs e)
        {
            _MenuItemEnabled(QuickLoadDataItem, false);
            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(QuickSettingsItem, false);
            _MenuItemEnabled(AnualDatesMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            _MenuItemEnabled(GetFioItem, false);
            _MenuItemEnabled(QuickFilterItem, false);
            _MenuItemEnabled(ViewMenuItem, false);

            if (sServer1.Length > 0 && sServer1UserName.Length > 0 && sServer1UserPassword.Length > 0)
            {
                Task.Run(() => _timer1Enabled(true));
                _ProgressBar1Value0();
                dataGridView1.Visible = true;
                pictureBox1.Visible = false;

                await Task.Run(() => GetFioFromServers());
                DeleteAllDataInTableQuery(databasePerson, "PersonTemp");
                DeleteAllDataInTableQuery(databasePerson, "PersonRegistered");
                DeleteAllDataInTableQuery(databasePerson, "PersonRegisteredFull");
                ShowDataTableQuery(databasePerson, "PersonTemp");
                panelViewResize();
                _MenuItemEnabled(QuickSettingsItem, true);
                _MenuItemEnabled(ViewMenuItem, true);
            }
            else { GetInfoSetup(); _MenuItemEnabled(QuickSettingsItem, true); }
        }

        private bool bServer1Exist = true;

        private void GetFioFromServers() //Get the list of registered users
        {
            string stringConnection;
            List<string> ListFIOTemp = new List<string>();
            listFIO = new List<string>();
            try
            {
                _comboBoxFioClr();
                _toolStripStatusLabel2AddText("Запрашиваю списки персонала со всех серверов СКД. Ждите окончания процесса...");
                stimerPrev = "Запрашиваю списки персонала со всех серверов СКД. Ждите окончания процесса...";
                stringConnection = "Data Source=" + sServer1 + "\\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID=" + sServer1UserName + ";Password=" + sServer1UserPassword + "; Connect Timeout=240";
                using (var sqlConnection = new SqlConnection(stringConnection))
                {
                    sqlConnection.Open();
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
                                        listFIO.Add(record["name"].ToString().Trim() + "|" + record["surname"].ToString().Trim() + "|" + record["patronymic"].ToString().Trim() + "|" + record["id"].ToString().Trim() + "|" +
                                                    record["tabnum"].ToString().Trim() + "|" + sServer1);
                                        ListFIOTemp.Add(record["name"].ToString().Trim() + " " + record["surname"].ToString().Trim() + " " + record["patronymic"].ToString().Trim() + "|" + record["tabnum"].ToString().Trim());
                                    }
                                } catch { }
                                _ProgressWork1();
                            }
                        }
                    }
                }

                _toolStripStatusLabel2AddText("Все списки с ФИО с серверов СКД успешно получены");
                stimerPrev = "Все списки с ФИО с сервера СКД успешно получены";
            } catch (Exception Expt)
            {
                bServer1Exist = false;
                stimerPrev = "Сервер не доступен или неправильная авторизация";
                MessageBox.Show(Expt.Message, @"Сервер не доступен или неправильная авторизация", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                { _comboBoxFioIndex(0); } catch { };
                _timer1Enabled(false);
                _toolStripStatusLabel2AddText("Получено ФИО - " + iFIO + " ");
                _toolStripStatusLabel2ForeColor(Color.Black);
                QuickLoadDataItem.Enabled = true;
            }
            _toolStripStatusLabel1Color(Color.Black);
            stringConnection = null;
            ListFIOCombo = null;
            ListFIOTemp = null;
            if (!bServer1Exist)
            {
                _toolStripStatusLabel2AddText("Ошибка доступа к SQL БД СКД-серверов!");
                _toolStripStatusLabel2BackColor(Color.DarkOrange);
                QuickLoadDataItem.Enabled = false;
                MessageBox.Show("Проверьте правильность написания серверов,\nимя и пароль sa-администратора,\nа а также доступность серверов и их баз!");
            }
            _ProgressBar1Value100();
            _MenuItemEnabled(GetFioItem, true);
            _MenuItemEnabled(FunctionMenuItem, true);
            _MenuItemEnabled(QuickSettingsItem, true);
            _MenuItemEnabled(AnualDatesMenuItem, true);
            _MenuItemEnabled(GroupsMenuItem, true);
        }

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
                            _textBoxFIOAddText(sCell[0]);   //иммитируем выбор данных
                            _textBoxNavAddText(sCell[1]);   //Select person   

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

            QuickFilterItem.BackColor = SystemColors.Control;
            panelViewResize();
        }

        private void FilterDataByNav()    //Copy Data from PersonRegistered into PersonTemp by Filter(NAV and anual dates or minimalTime or dayoff)
        {
            if (checkBoxReEnter.Checked)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    HashSet<string> AllDateRegistration = new HashSet<string>();
                    using (var sqlCommand = new SQLiteCommand("SELECT  *, MIN(Comming) AS FirstRegistered FROM PersonRegistered  " +
                        " WHERE NAV like '" + textBoxNav.Text.Trim() + "' GROUP BY FIO, NAV, DateRegistered ORDER BY DateRegistered ASC;", sqlConnection))
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

                                        dControlHourSelected = Convert.ToDecimal(record["HourControlling"].ToString().Trim());
                                        dControlMinuteSelected = Convert.ToDecimal(record["MinuteControlling"].ToString().Trim());
                                    }
                                } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
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
                        " where NAV like '" + textBoxNav.Text.Trim() + "' order by FIO, DateRegistered, Comming ASC;", sqlConnection))
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
                                } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
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

            if (checkBoxWeekend.Checked)
            { DeleteAnualDates(databasePerson, "PersonTemp"); }

            if (checkBoxStartWorkInTime.Checked)
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

        private void _dataGridView1Enabled(bool bEnabled) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { dataGridView1.Enabled = bEnabled; }));
            else
                dataGridView1.Enabled = bEnabled;
        }

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

        private void ExportDatagridToExcel()  //Export to Excel from DataGridView
        {
            _MenuItemEnabled(QuickLoadDataItem, false);
            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(QuickSettingsItem, false);
            _MenuItemEnabled(AnualDatesMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            _MenuItemEnabled(QuickSettingsItem, false);
            _MenuItemEnabled(ViewMenuItem, false);
            _dataGridView1Enabled(false);

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
            _toolStripStatusLabel2ForeColor(Color.Black);
            sLastSelectedElement = "ExportExcel";
            iDGCollumns = 0; iDGRows = 0;
            _toolStripStatusLabel2AddText("Готово!");
            _MenuItemEnabled(QuickLoadDataItem, true);
            _MenuItemEnabled(FunctionMenuItem, true);
            _MenuItemEnabled(QuickSettingsItem, true);
            _MenuItemEnabled(AnualDatesMenuItem, true);
            _MenuItemEnabled(GroupsMenuItem, true);
            _MenuItemEnabled(QuickSettingsItem, true);
            _MenuItemEnabled(ViewMenuItem, true);
            _dataGridView1Enabled(true);
        }

        private void releaseObject(object obj) //for function - ExportDatagridToExcel()
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            } catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object of Excel \n" + ex);
            } finally
            { GC.Collect(); }
        }

        private async void buttonExport_Click(object sender, EventArgs e)
        {
            Task.Run(() => _timer1Enabled(true));
            Task.Run(() => _toolStripStatusLabel2AddText("Генерирую Excel-файл"));
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
            checkBoxReEnter.Checked = false;
            checkBoxStartWorkInTime.Checked = false;
            checkBoxCelebrate.Checked = false;
            checkBoxWeekend.Checked = false;
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
            } catch { }
            if (comboBoxFio.SelectedIndex > -1)
            {
                QuickLoadDataItem.BackColor = Color.PaleGreen;
                groupBoxPeriod.BackColor = System.Drawing.Color.PaleGreen;
                groupBoxTime.BackColor = System.Drawing.Color.PaleGreen;
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

                QuickLoadDataItem.BackColor = System.Drawing.Color.PaleGreen;
                // QuickFilterItem.BackColor = SystemColors.Control;
                groupBoxPeriod.BackColor = System.Drawing.Color.PaleGreen;
                groupBoxTime.BackColor = System.Drawing.Color.PaleGreen;
                groupBoxRemoveDays.BackColor = SystemColors.Control;
            } catch { }
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
                } catch { }

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
                        _toolStripStatusLabel2BackColor(SystemColors.Control);
                        bErrorData = true;
                    } catch { }
                }
                else if (sTextGroup.Length > 0 && textBoxNav.Text.Trim().Length == 0 && sTextFIOSelected.Length > 10)
                    try
                    {
                        StatusLabel2.Text = "У сотрудника отсутствует NAV-код:" + ShortFIO(Regex.Split(sTextFIOSelected, "[|]")[0].Trim());
                        _toolStripStatusLabel2BackColor(Color.DarkOrange);
                        bErrorData = true;
                    } catch { }
                else if (sTextGroup.Length == 0 && textBoxNav.Text.Trim().Length > 0 && sTextFIOSelected.Length > 10)
                    try
                    {
                        StatusLabel2.Text = "Не указана группа, в которую нужно добавить!";
                        _toolStripStatusLabel2BackColor(Color.DarkOrange);
                        bErrorData = true;
                    } catch { }
                else
                    try
                    {
                        StatusLabel2.Text = "Проверьте вводимые данные!";
                        _toolStripStatusLabel2BackColor(Color.DarkOrange);
                        bErrorData = true;
                    } catch { }
            }

            MembersGroup();
            PersonOrGroupItem.Text = "Работа с одной персоной";
            controlStartHours = 0;
            sTextGroup = null; sTextFIOSelected = null;
            labelGroup.BackColor = SystemColors.Control;
        }

        private void GetNamePoints()
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
                                } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
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

        private void GetDataItem_Click(object sender, EventArgs e)
        {
            GetDataItem();
        }

        private async void GetDataItem()
        {
            groupBoxPeriod.BackColor = SystemColors.Control;
            groupBoxTime.BackColor = SystemColors.Control;

            _ChangeMenuItemBackColor(QuickLoadDataItem, SystemColors.Control);
            checkBoxReEnter.Checked = false;
            checkBoxStartWorkInTime.Checked = false;
            checkBoxWeekend.Checked = false;
            checkBoxCelebrate.Checked = false;
            _MenuItemEnabled(QuickLoadDataItem, false);
            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(QuickSettingsItem, false);
            _MenuItemEnabled(AnualDatesMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            _MenuItemEnabled(QuickSettingsItem, false);
            _MenuItemEnabled(ViewMenuItem, false);
            _dataGridView1Enabled(false);

            if (sServer1.Length > 0 && sServer1UserName.Length > 0 && sServer1UserPassword.Length > 0)
            {
                Task.Run(() => _timer1Enabled(true));

                if ((nameOfLastTableFromDB == "PersonGroupDesciption" || nameOfLastTableFromDB == "PersonGroup") && _textBoxGroupText().Length > 0)
                {
                    Task.Run(() => _toolStripStatusLabel2AddText("Получаю данные по группе " + _textBoxGroupText()));
                    stimerPrev = "Получаю данные по группе " + _textBoxGroupText();
                }
                else
                {
                    Task.Run(() => _toolStripStatusLabel2AddText("Получаю данные по \"" + ShortFIO(_textBoxFIOText()) + "\" "));
                    stimerPrev = "Получаю данные по \"" + ShortFIO(_textBoxFIOText()) + "\" ";
                    nameOfLastTableFromDB = "PersonRegistered";
                }
                _ProgressBar1Value0();
                pictureBox1.Visible = false;
                await Task.Run(() => GetDataGroupItem());
                dataGridView1.Visible = true;
                _dataGridView1Enabled(true);
                VisualItem.Enabled = true;
                VisualWorkedTimeItem.Visible = true;
                ExportIntoExcelItem.Enabled = true;
                panelViewResize();
                _MenuItemEnabled(QuickSettingsItem, true);
                _MenuItemEnabled(ViewMenuItem, true);
            }
            else { GetInfoSetup(); _MenuItemEnabled(QuickSettingsItem, true); }
            groupBoxRemoveDays.BackColor = Color.PaleGreen;
        }

        private void GetDataGroupItem()
        {
            //Clear work tables
            DeleteAllDataInTableQuery(databasePerson, "PersonTemp");
            DeleteAllDataInTableQuery(databasePerson, "PersonRegistered");
            DeleteAllDataInTableQuery(databasePerson, "PersonRegisteredFull");
            GetNamePoints();  //Get names of the points

            if ((nameOfLastTableFromDB == "PersonGroupDesciption" || nameOfLastTableFromDB == "PersonGroup") && _textBoxGroupText().Length > 1)
            {
                _toolStripStatusLabel2AddText("Получаю данные по группе " + _textBoxGroupText());
                stimerPrev = "Получаю данные по группе " + _textBoxGroupText();
                GetGroupInfoFromDB();

                string[] sCell;
                foreach (string sRow in lListFIOTemp.ToArray())
                {
                    sCell = Regex.Split(sRow, "[|]"); //FIO|NAV|H|M
                    if (sCell[0].Length > 1)
                    {
                        listFIO.Clear();
                        _textBoxFIOAddText(sCell[0]);   //иммитируем выбор данных
                        _textBoxNavAddText(sCell[1]);   //Select person                  
                        listFIO.Add(sCell[0]);      // add the person into the list  

                        dControlHourSelected = Convert.ToDecimal(sCell[2]);
                        dControlMinuteSelected = Convert.ToDecimal(sCell[3]);

                        GetRegistrationFromServer();   //Search Registration at checkpoints of the selected person
                    }
                }

                nameOfLastTableFromDB = "PersonGroup";
                _timer1Enabled(false);
                _toolStripStatusLabel2AddText("Данные по группе \"" + _textBoxGroupText() + "\" получены");
            }
            else
            {
                dControlHourSelected = Convert.ToDecimal(numUpDownHour.Value);
                dControlMinuteSelected = Convert.ToDecimal(numUpDownMinute.Value);

                stimerPrev = "Получаю данные по \"" + ShortFIO(_textBoxFIOText()) + "\"";
                GetRegistrationFromServer();
                nameOfLastTableFromDB = "PersonRegistered";
                _timer1Enabled(false);
                _toolStripStatusLabel2AddText("Данные с СКД по \"" + ShortFIO(_textBoxFIOText()) + "\" получены!");
            }

            CopyWholeDataFromOneTableIntoAnother(databasePerson, "PersonTemp", "PersonRegistered");
            ShowDataTableQuery(databasePerson, "PersonTemp");

            _SetMenuItemDefaultColor(QuickLoadDataItem);
            _ChangeMenuItemBackColor(QuickFilterItem, Color.PaleGreen);
            _ChangeMenuItemBackColor(ExportIntoExcelItem, Color.PaleGreen);

            _ProgressBar1Value100();
            _MenuItemEnabled(QuickLoadDataItem, true);
            _MenuItemEnabled(FunctionMenuItem, true);
            _MenuItemEnabled(QuickSettingsItem, true);
            _MenuItemEnabled(AnualDatesMenuItem, true);
            _MenuItemEnabled(GroupsMenuItem, true);
            _toolStripStatusLabel2ForeColor(Color.Black);
            stimerPrev = "";
        }


        private void GetRegistrationFromServer()
        {
            string personNAV = _textBoxNavText(); ; string personFIO = _textBoxFIOText();
            stimerPrev = "Получаю данные по \"" + ShortFIO(personFIO) + "\"";
            _toolStripStatusLabel2AddText("Получаю данные по \"" + ShortFIO(personFIO) + "\"");

            decimal hourControlStart = dControlHourSelected;
            decimal minuteControlStart = dControlMinuteSelected;
            decimal controlStartHours = hourControlStart + (minuteControlStart + 1) / 60 - (1 / 60);
            string stringIdCardIntellect = "";
            string stringIdCardSKD = "";
            string personNAVTemp = "";
            string[] stringSelectedFIO = new string[3];
            try { stringSelectedFIO[0] = Regex.Split(personFIO, "[ ]")[0]; } catch { stringSelectedFIO[0] = ""; }
            try { stringSelectedFIO[1] = Regex.Split(personFIO, "[ ]")[1]; } catch { stringSelectedFIO[1] = ""; }
            try { stringSelectedFIO[2] = Regex.Split(personFIO, "[ ]")[2]; } catch { stringSelectedFIO[2] = ""; }

            try
            {
                if (personNAV.Length != 6)
                {
                    foreach (var strRowWithNav in listFIO.ToArray())
                    {
                        if (strRowWithNav.Contains(personNAV) && personNAV.Length > 0 && strRowWithNav.Contains(sServer1))
                            try { stringIdCardIntellect = Regex.Split(strRowWithNav, "[|]")[3].Trim(); } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                        _ProgressWork1();
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
                                        if (record?["tabnum"].ToString().Trim().Length > 0)
                                        { stringIdCardIntellect = record["id"].ToString().Trim(); _ProgressWork1(); break; }
                                    } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                                }
                            }
                        }
                    }
                }

                if (personNAV.Length < 1 && personNAV.Length != 6)
                { stringIdCardSKD = "0"; }
            } catch (Exception expt) { MessageBox.Show(expt.ToString()); }

            try
            {
                string stringConnection;
                string stringSqlWhere = "";

                decimal idCardIntellect = 0;
                decimal idCardSKD = 0;

                decimal hourManaging = 0;
                decimal minuteManaging = 0;
                decimal managingHours = 0;

                try { idCardIntellect = Convert.ToDecimal(stringIdCardIntellect); } catch { idCardIntellect = 0; }
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
                                            hourManaging = 0;
                                            minuteManaging = 0;
                                            managingHours = 0;

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
                                    } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                                }
                            }
                        }
                    }
                }

                try { idCardSKD = Convert.ToDecimal(stringIdCardSKD); } catch { idCardSKD = 0; }
                stringConnection = null;
                stringSqlWhere = null;
                _ProgressWork1();
            } catch (Exception Expt)
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
                                    } catch
                                    {
                                    }
                            }

                            iCounterLine++;
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
            stringIdCardIntellect = null; stringIdCardSKD = null; personNAVTemp = null; stringSelectedFIO = new string[1];
        }

        private void GetGroupInfoFromDB() //Get info the selected group from DB and make a few lists with these data
        {
            lListFIOTemp.Clear();

            using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
            {
                sqlConnection.Open();
                using (var sqlCommand = new SQLiteCommand(
                    "Select * FROM PersonGroup where GroupPerson like '" + _textBoxGroupText() + "';", sqlConnection))
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
                    DeleteDataTableQuery(databasePerson, "PersonGroup", " where NAV like '%" + _textBoxNavText() + "%'", "NAV", dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString().Trim(),
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
                            groupBoxPeriod.BackColor = System.Drawing.Color.PaleGreen;
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
                                if (dataGridView1.Columns[i].HeaderText.ToString() == "Фамилия Имя Отчество" ||
                                        dataGridView1.Columns[i].HeaderText.ToString() == "FIO")
                                    IndexColumn1 = i;
                                if (dataGridView1.Columns[i].HeaderText.ToString() == "NAV-код" ||
                                        dataGridView1.Columns[i].HeaderText.ToString() == "NAV")
                                    IndexColumn2 = i;
                                if (dataGridView1.Columns[i].HeaderText.ToString() == "Контрольное время, часы" ||
                                        dataGridView1.Columns[i].HeaderText.ToString() == "HourControlling")
                                    IndexColumn3 = i;
                                if (dataGridView1.Columns[i].HeaderText.ToString() == "Контрольное время, минуты" ||
                                        dataGridView1.Columns[i].HeaderText.ToString() == "MinuteControlling")
                                    IndexColumn4 = i;
                            }

                            textBoxFIO.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString();
                            textBoxNav.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn2].Value.ToString(); //Take the name of selected group
                            StatusLabel2.Text = @"Выбрана группа: " + textBoxGroup.Text + @" | Курсор на: " + textBoxFIO.Text;
                            groupBoxPeriod.BackColor = System.Drawing.Color.PaleGreen;
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

                            textBoxNav.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn2].Value.ToString(); //Take the name of selected group
                            textBoxFIO.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString();
                            StatusLabel2.Text = @"Выбран: " + textBoxFIO.Text + @" |  Всего ФИО: " + iFIO;
                            groupBoxPeriod.BackColor = System.Drawing.Color.PaleGreen;
                            groupBoxTime.BackColor = System.Drawing.Color.PaleGreen;
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
                } catch (Exception expt)
                {
                    MessageBox.Show(expt.ToString());
                }
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
            ViewMenuItem.Enabled = false;
            QuickSettingsItem.Enabled = false;
            QuickLoadDataItem.Enabled = false;
            QuickFilterItem.Enabled = false;
            comboBoxFio.Enabled = false;
            checkBoxReEnter.Enabled = false;
            checkBoxStartWorkInTime.Enabled = false;
            checkBoxCelebrate.Enabled = false;
            checkBoxWeekend.Enabled = false;
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
            ViewMenuItem.Enabled = true;
            QuickSettingsItem.Enabled = true;
            QuickLoadDataItem.Enabled = true;
            QuickFilterItem.Enabled = true;
            comboBoxFio.Enabled = true;
            checkBoxReEnter.Enabled = true;
            checkBoxStartWorkInTime.Enabled = true;
            checkBoxCelebrate.Enabled = true;
            checkBoxWeekend.Enabled = true;

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
        private string _textBoxGroupText() //add string into  from other threads
        {
            string tBox = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { tBox = textBoxGroup.Text.Trim(); }));
            else
                tBox = textBoxGroup.Text.Trim();
            return tBox;
        }

        private void _textBoxNavAddText(string s) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { textBoxNav.Text = s; }));
            else
                textBoxNav.Text = s;
        }

        private string _textBoxNavText() //add string into  from other threads
        {
            string tBox = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { tBox = textBoxNav.Text.Trim(); }));
            else
                tBox = textBoxNav.Text;
            return tBox;
        }

        private void _textBoxFIOAddText(string s) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { textBoxFIO.Text = s; }));
            else
                textBoxFIO.Text = s;
        }

        private string _textBoxFIOText() //add string into  from other threads
        {
            string tBox = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { tBox = textBoxFIO.Text.Trim(); }));
            else
                tBox = textBoxFIO.Text.Trim();
            return tBox;
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

        private void _numUpDownHourValue(int i) //add string into comboBoxTargedPC from other threads
        {
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate { try { numUpDownHour.Value = i; } catch { numUpDownHour.Value = 9; } }));
            }
            else
            {
                try { numUpDownHour.Value = i; } catch { numUpDownHour.Value = 9; }
            }
        }

        private decimal _numUpDownHour()
        {
            decimal iCombo = 0;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { iCombo = (decimal)numUpDownHour.Value; }));
            else
                iCombo = (decimal)numUpDownHour.Value;
            return iCombo;
        }

        private void _numUpDownMinuteValue(int i) //add string into comboBoxTargedPC from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { numUpDownMinute.Value = i; }));
            else
                numUpDownMinute.Value = i;
        }

        private decimal _numUpDownMinute()
        {
            decimal iCombo = 0;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { iCombo = (decimal)numUpDownMinute.Value; }));
            else
                iCombo = (decimal)numUpDownMinute.Value;
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

        private void _timer1Enabled(bool s) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    timer1.Enabled = s; if (s) timer1.Start(); else timer1.Stop();
                }));
            else
            { timer1.Enabled = s; if (s) timer1.Start(); else timer1.Stop(); }
        }

        private void _toolStripStatusLabel2AddText(string s) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { StatusLabel2.Text = s; }));
            else
                StatusLabel2.Text = s;
        }

        private void _toolStripStatusLabel1Color(Color s)
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { StatusLabel1.ForeColor = s; }));
            else
                StatusLabel1.ForeColor = s;
        }

        private void _toolStripStatusLabel2ForeColor(Color s)
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { StatusLabel2.ForeColor = s; }));
            else
                StatusLabel2.ForeColor = s;
        }

        private void _toolStripStatusLabel2BackColor(Color s) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { StatusLabel2.BackColor = s; }));
            else
                StatusLabel2.BackColor = s;
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

        private void _ChangeMenuItemBackColor(ToolStripMenuItem tMenuItem, System.Drawing.Color colorMenu) //add string into  from other threads
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

        private void _SetMenuItemDefaultColor(ToolStripMenuItem tMenuItem) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { tMenuItem.BackColor = SystemColors.Control; }));
            else
                tMenuItem.BackColor = SystemColors.Control;
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
        //End of Block. Access to Controls from other threads

        string stimerPrev = "";
        string stimerCurr = "Ждите!";
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

        private void checkBoxReEnter_CheckStateChanged(object sender, EventArgs e)
        {
            QuickFilterItem.BackColor = System.Drawing.Color.PaleGreen;
            if (checkBoxReEnter.Checked)
            {
                QuickFilterItem.Enabled = true;
                QuickFilterItem.BackColor = System.Drawing.Color.PaleGreen;
                QuickLoadDataItem.BackColor = SystemColors.Control;
            }
            else
            {
                QuickFilterItem.Enabled = true;
                QuickFilterItem.BackColor = SystemColors.Control;
            }
        }

        private void checkBoxStartWorkInTime_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBoxStartWorkInTime.Checked)
            {
                checkBoxReEnter.Checked = true;
                checkBoxReEnter.Enabled = false;
                QuickFilterItem.BackColor = System.Drawing.Color.PaleGreen;
                QuickLoadDataItem.BackColor = SystemColors.Control;
            }
            else
            {
                checkBoxReEnter.Enabled = true;
                QuickFilterItem.BackColor = SystemColors.Control;
            }
        }

        private void checkBoxWeekend_CheckStateChanged(object sender, EventArgs e)
        { QuickFilterItem.BackColor = System.Drawing.Color.PaleGreen; }

        private void checkBoxCelebrate_CheckStateChanged(object sender, EventArgs e)
        { QuickFilterItem.BackColor = System.Drawing.Color.PaleGreen; }

        private void ClearReportItem_Click(object sender, EventArgs e) //Clear
        {
            DeleteTable(databasePerson, "PersonTemp");
            DeleteTable(databasePerson, "PersonRegistered");
            DeleteTable(databasePerson, "PersonRegisteredFull");
            textBoxFIO.Text = "";
            textBoxGroup.Text = "";
            textBoxGroupDescription.Text = "";
            textBoxNav.Text = "";
            GC.Collect();
            MakeDB();
            ShowDataTableQuery(databasePerson, "PersonRegisteredFull");
            StatusLabel2.Text = @"Отчеты удалены";
        }

        private void ClearDataItem_Click(object sender, EventArgs e)
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
            MakeDB();
        }

        private void ClearAllItem_Click(object sender, EventArgs e)
        {
            DeleteDB();
        }

        private void DeleteDB()
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
                MakeDB();
            }
            else
            { MakeDB(); }
            StatusLabel2.Text = "Все данные удалены. База пересоздана";
        }

        private void VisualItem_Click(object sender, EventArgs e)
        {
            //   CountDataInTheTableQuery("PersonRegistered");
            //    if (iCounterLine > 0) bLoaded = true; else bLoaded = false;

            if (bLoaded) { SelectPersonFromDataGrid(); }

            if (bLoaded && (nameOfLastTableFromDB == "PersonRegistered" || nameOfLastTableFromDB == "PersonGroup"))
            {
                dataGridView1.Visible = false;
                FindWorkDatesInSelected();
                DrawRegistration();
                ReportsItem.Visible = true;
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
                                } catch { }
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
                    groupBoxPeriod.BackColor = System.Drawing.Color.PaleGreen;
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
                    // numUpDownHour.Value = 9;
                    //  numUpDownMinute.Value = 0;
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

        private void DrawRegistration()  // Draw registration
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

            RefreshPictureBox(pictureBox1, bmp);
            panelViewResize();
        }

        private string ShortFIO(string s) //Transform FULL FIO into Short form
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
            //Disable it for PictureBox set at the Center
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
            RefreshPictureBox(pictureBox1, bmp);
            panelViewResize();
        }

        private void RefreshPictureBox(PictureBox picBox, Bitmap picImage) // не работает
        {
            picBox.Image = RefreshBitmap(picImage, panelView.Width - 2, panelView.Height - 2); //сжатая картина
            picBox.Refresh();
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
        }

        private void panelView_SizeChanged(object sender, EventArgs e)
        {
            panelViewResize();
        }

        private void panelViewResize() //Change PanelView
        {
            int iStringHeight = 19;
            int iShiftHeightAll = 36;
            switch (sLastSelectedElement)
            {
                case "DrawFullWorkedPeriodRegistration":
                    panelView.Height = iShiftHeightAll + iStringHeight * workSelectedDays.Length; //Fixed size of Picture. If need autosize - disable this row
                    break;
                case "DrawRegistration":
                    panelView.Height = iShiftHeightAll + iStringHeight * workSelectedDays.Length; //Fixed size of Picture. If need autosize - disable this row
                    break;
                default:
                    panelView.Anchor = (AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top);
                    panelView.Height = panelView.Parent.Height - 120;
                    panelView.AutoScroll = true;
                    panelView.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                    panelView.ResumeLayout();
                    break;
            }

            if (panelView.Controls.Count > 1)
            {
                RefreshPictureBox(pictureBox1, bmp);
            }
        }

        private void dateTimePickerStart_CloseUp(object sender, EventArgs e)
        {
            QuickLoadDataItem.Enabled = true;
            QuickLoadDataItem.BackColor = System.Drawing.Color.PaleGreen;
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
        {
            panelViewResize();
            panelView.Visible = false;
            _MenuItemEnabled(QuickLoadDataItem, false);
            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(QuickSettingsItem, false);
            _MenuItemEnabled(AnualDatesMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            _MenuItemEnabled(ViewMenuItem, false);
            _MenuItemEnabled(QuickFilterItem, false);

            labelServer1 = new Label
            {
                Text = "Server1",
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
            toolTip1.SetToolTip(textBoxServer1, "Имя сервера \"Server1\" с базой Intellect в виде - NameOfServer.Domain.Subdomain");

            labelServer1UserName = new Label
            {
                Text = "UserName",
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
            toolTip1.SetToolTip(textBoxServer1UserName, "Имя администратора SQL-сервера \"Server1\"");

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
            toolTip1.SetToolTip(textBoxServer1UserPassword, "Пароль администратора SQL-сервера \"Server1\"");

            textBoxServer1.BringToFront();
            textBoxServer1UserName.BringToFront();
            textBoxServer1UserPassword.BringToFront();
            labelServer1UserName.BringToFront();
            labelServer1UserPassword.BringToFront();

            groupBoxProperties.Visible = true;
        }

        private void buttonPropertiesSave_Click(object sender, EventArgs e)
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
                _MenuItemEnabled(ViewMenuItem, true);

                panelView.Visible = true;
            }
            else
            {
                GetInfoSetup(); _MenuItemEnabled(QuickSettingsItem, true);
            }
        }

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
            _MenuItemEnabled(QuickFilterItem, true);
            _MenuItemEnabled(AnualDatesMenuItem, true);
            _MenuItemEnabled(GroupsMenuItem, true);
            _MenuItemEnabled(ViewMenuItem, true);

            panelView.Visible = true;
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
        {
            GetInfoSetup();
        }

        private void GetInfoSetup()
        {
            DialogResult result = MessageBox.Show(
                @"Перед получением информации необходимо внести:" + "\n\n" +
                 "1. Имя сервера Интеллект (СКД1 - SERVER1.DOMAIN) , а также имя и пароль пользователя SQL-сервера данного СКД\n" +
                 "2. Сохранить данные параметры\n" +
                 "3. После внесения этих данных можно получать списки пользователей проходивших пукты регистрации, " +
                 "просматривать данные по регистрациям и проводить анализ.\n\nДата и время локального ПК: " +
                dateTimePickerEnd.Value,
                @"Информация о программе",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }

        private void textBoxGroup_TextChanged(object sender, EventArgs e)
        {
            if (textBoxGroup.Text.Trim().Length > 0)
            {
                AddPersonToGroupItem.Enabled = true;
                CreateGroupItem.Visible = true;
            }
            else
            { StatusLabel2.Text = @"Всего ФИО: " + iFIO; }
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
            UpdateControllingItem.Visible = false;

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

                numUpDownHour.Value = Convert.ToInt32(dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn3].Value.ToString());
                numUpDownMinute.Value = Convert.ToInt32(dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn4].Value.ToString());
                decimal hourControlStart = dControlHourSelected;
                decimal minuteControlStart = dControlMinuteSelected;
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
                        command.Parameters.Add("@HourControlling", DbType.String).Value = numUpDownHour.Value.ToString();
                        command.Parameters.Add("@MinuteControlling", DbType.String).Value = numUpDownMinute.Value.ToString();
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
        }

        private void ViewMenuItem_Click(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// /////////////////////Registry  -   TODO
        /// </summary>

        public void ListsRegistryDataCheck() //Read previously Saved Parameters from Registry
        {
            List<string> lstSavedServices = new List<string>();
            List<string> lstSavedNumbers = new List<string>();
            bool foundSavedData;
            string[] getValue;

            try
            {
                using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                      myRegKey,
                      Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree,
                      System.Security.AccessControl.RegistryRights.ReadKey))
                {
                    getValue = (string[])EvUserKey.GetValue("ListServices");

                    try
                    {
                        foreach (string line in getValue)
                        {
                            lstSavedServices.Add(line);
                        }
                        foundSavedData = true;
                    } catch { }

                    getValue = (string[])EvUserKey.GetValue("ListNumbers");

                    try
                    {
                        foreach (string line in getValue)
                        {
                            lstSavedNumbers.Add(line);
                        }
                        foundSavedData = true;
                    } catch { }

                    //strSavedPathToInvoice = (string)EvUserKey.GetValue("PathToLastInvoice");
                }
            } catch (Exception exct)
            {
                // textBoxLog.AppendText("\n" + exct.ToString() + "\n");
            }
        }

        public void ListServicesRegistrySave() //Save Parameters into Registry and variables
        {
            List<string> listServices = new List<string>(); bool foundSavedData;

            try
            {
                using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(myRegKey))
                {
                    EvUserKey.SetValue("ListServices", listServices.ToArray(),
                        Microsoft.Win32.RegistryValueKind.MultiString);
                }
                foundSavedData = true;
            } catch { MessageBox.Show("Ошибки с доступом для записи списка сервисов в реестр. Данные сохранены не корректно."); }
        }

        public void ListNumbersRegistrySave() //Save inputed Credintials and Parameters into Registry and variables
        {
            List<string> listNumbers = new List<string>(); bool foundSavedData;

            try
            {
                using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(myRegKey))
                {
                    EvUserKey.SetValue("ListNumbers", listNumbers.ToArray(),
                        Microsoft.Win32.RegistryValueKind.MultiString);
                }
                foundSavedData = true;
            } catch { MessageBox.Show("Ошибки с доступом для записи списка номеров в реестр. Данные сохранены не корректно."); }
        }


        public void PathToLastInvoiceRegistrySave() //Save Parameters into Registry and variables
        {
            string filepathLoadedData = ""; bool foundSavedData;

            try
            {
                using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(myRegKey))
                {
                    EvUserKey.SetValue("PathToLastInvoice", filepathLoadedData,
                        Microsoft.Win32.RegistryValueKind.String);
                }
                foundSavedData = true;
            } catch { MessageBox.Show("Ошибки с доступом для записи пути к счету. Данные сохранены не корректно."); }

        }


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

