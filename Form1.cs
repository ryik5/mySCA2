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
        private string myRegKey = @"SOFTWARE\RYIK\PersonViewerSCA2";

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
        private int iFIO = 0;

        //Controls "NumUpDown"
        private decimal numUpHourStart = 9;
        private decimal numUpMinuteStart = 0;
        private decimal numUpHourEnd = 18;
        private decimal numUpMinuteEnd = 0;

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

        private Label labelMailServerName;
        private TextBox textBoxMailServerName;
        private Label labelMailServerUserName;
        private TextBox textBoxMailServerUserName;
        private Label labelMailServerUserPassword;
        private TextBox textBoxMailServerUserPassword;

        private Color clrRealRegistration = Color.PaleGreen;
        private string sLastSelectedElement = "MainForm";

        //Settings of Programm
        private bool bServer1Exist = false;
        private string sServer1 = "";
        private string sServer1Registry = "";
        private string sServer1DB = "";
        private string sServer1UserName = "";
        private string sServer1UserNameRegistry = "";
        private string sServer1UserNameDB = "";
        private string sServer1UserPassword = "";
        private string sServer1UserPasswordRegistry = "";
        private string sServer1UserPasswordDB = "";

        private string mailServer = "";
        private string mailServerRegistry = "";
        private string mailServerDB = "";
        private string mailServerUserName = "";
        private string mailServerUserNameRegistry = "";
        private string mailServerUserNameDB = "";
        private string mailServerUserPassword = "";
        private string mailServerUserPasswordRegistry = "";
        private string mailServerUserPasswordDB = "";

        private Label listComboLabel;
        private ComboBox listCombo = new ComboBox();

        private Label periodComboLabel;
        private ComboBox periodCombo = new ComboBox();


        private readonly byte[] btsMess1 = Convert.FromBase64String(@"OCvesunvXXsxtt381jr7vp3+UCwDbE4ebdiL1uinGi0="); //Key Encrypt
        private readonly byte[] btsMess2 = Convert.FromBase64String(@"NO6GC6Zjl934Eh8MAJWuKQ=="); //Key Decrypt

        private OpenFileDialog openFileDialog1 = new OpenFileDialog();
        private List<string> listGroups = new List<string>();

        //DataTables with people data
        private DataTable dtPeople = new DataTable("People");
        private DataColumn[] dcPeople =
            {
                                  new DataColumn(@"№ п/п",typeof(int)),//0
                                  new DataColumn(@"Фамилия Имя Отчество",typeof(string)),//1
                                  new DataColumn(@"NAV-код",typeof(string)),//2
                                  new DataColumn(@"Группа",typeof(string)),//3
                                  new DataColumn(@"Время прихода,часы",typeof(string)),//4
                                  new DataColumn(@"Время прихода,минут",typeof(string)), //5
                                  new DataColumn(@"Время прихода",typeof(decimal)),//6
                                  new DataColumn(@"Время ухода,часы",typeof(string)),//7
                                  new DataColumn(@"Время ухода,минут",typeof(string)),//8
                                  new DataColumn(@"Время ухода",typeof(decimal)),//9
                                  new DataColumn(@"№ пропуска",typeof(int)), //10
                                  new DataColumn(@"Отдел",typeof(string)),//11
                                  new DataColumn(@"Дата регистрации",typeof(string)),//12
                                  new DataColumn(@"Время регистрации,часы",typeof(string)),//13
                                  new DataColumn(@"Время регистрации,минут",typeof(string)),//14
                                  new DataColumn(@"Время регистрации",typeof(decimal)), //15
                                  new DataColumn(@"Реальное время ухода,часы",typeof(string)),//16
                                  new DataColumn(@"Реальное время ухода,минут",typeof(string)),//17
                                  new DataColumn(@"Реальное время ухода",typeof(decimal)), //18
                                  new DataColumn(@"Сервер СКД",typeof(string)), //19
                                  new DataColumn(@"Имя точки прохода",typeof(string)), //20
                                  new DataColumn(@"Направление прохода",typeof(string)), //21
                                  new DataColumn(@"Время прихода ЧЧ:ММ",typeof(string)),//22
                                  new DataColumn(@"Время ухода ЧЧ:ММ",typeof(string)),//23
                                  new DataColumn(@"Реальное время прихода ЧЧ:ММ",typeof(string)),//24
                                  new DataColumn(@"Реальное время ухода ЧЧ:ММ",typeof(string)), //25
                                  new DataColumn(@"Реальное отработанное время",typeof(decimal)), //26
                                  new DataColumn(@"Реальное отработанное время ЧЧ:ММ",typeof(string)), //27
                                  new DataColumn(@"Опоздание",typeof(string)),                    //28
                                  new DataColumn(@"Ранний уход",typeof(string)),                 //29
                                  new DataColumn(@"Отпуск (отгул)",typeof(string)),                 //30
                                  new DataColumn(@"Коммандировка",typeof(string)),                 //31
                                  new DataColumn(@"День недели",typeof(string)),                 //32
                                  new DataColumn(@"Больничный",typeof(string)),                 //33
                                  new DataColumn(@"Согласованное отсутствие",typeof(string)),   //34
                                  new DataColumn(@"Код",typeof(string)),                     //35
                                  new DataColumn(@"Вышестоящая группа",typeof(string)),//36
                                  new DataColumn(@"Описание группы",typeof(string)),   //37
                };
        private string[] arrayAllCollumnsDataTablePeople =
            {
                                  @"№ п/п",//0
                                  @"Фамилия Имя Отчество",//1
                                  @"NAV-код",//2
                                  @"Группа",//3
                                  @"Время прихода,часы",//4
                                  @"Время прихода,минут", //5
                                  @"Время прихода",//6
                                  @"Время ухода,часы",//7
                                  @"Время ухода,минут",//8
                                  @"Время ухода",//9
                                  @"№ пропуска", //10
                                  @"Отдел",//11
                                  @"Дата регистрации",//12
                                  @"Время регистрации,часы",//13
                                  @"Время регистрации,минут",//14
                                  @"Время регистрации", //15
                                  @"Реальное время ухода,часы",//16
                                  @"Реальное время ухода,минут",//17
                                  @"Реальное время ухода", //18
                                  @"Сервер СКД", //19
                                  @"Имя точки прохода", //20
                                  @"Направление прохода", //21
                                  @"Время прихода ЧЧ:ММ",//22
                                  @"Время ухода ЧЧ:ММ",//23
                                  @"Реальное время прихода ЧЧ:ММ",//24
                                  @"Реальное время ухода ЧЧ:ММ", //25
                                  @"Реальное отработанное время", //26
                                  @"Реальное отработанное время ЧЧ:ММ", //27
                                  @"Опоздание",                    //28
                                  @"Ранний уход",                 //29
                                  @"Отпуск (отгул)",                 //30
                                  @"Коммандировка",                 //31
                                  @"День недели",                    //32
                                  @"Больничный",                    //33
                                  @"Согласованное отсутствие",      //34
                                  @"Код",                           //35
                                  @"Вышестоящая группа",            //36
                                  @"Описание группы"                //37
        };

        private DataTable dtPersonTemp = new DataTable("PersonTemp");
        private DataTable dtPersonTempAllCollumns = new DataTable("PersonTempAllCollumns");
        private DataTable dtPersonRegisteredFull = new DataTable("PersonRegisteredFull");
        private DataTable dtPersonRegistered = new DataTable("PersonRegistered");
        private DataTable dtPersonGroup = new DataTable("PersonGroup");
        private DataTable dtPersonsLastList = new DataTable("PersonsLastList");
        private DataTable dtPersonsLastComboList = new DataTable("PersonsLastComboList");

        private DataTable dtGroup = new DataTable("Group");
        private DataColumn[] dcGroup =
                            {
                                  new DataColumn(@"№ п/п",typeof(int)),              //0
                                  new DataColumn(@"Код",typeof(string)),             //1
                                  new DataColumn(@"Группа",typeof(string)),          //2
                                  new DataColumn(@"Вышестоящая группа",typeof(string)),//3
                                  new DataColumn(@"Описание группы",typeof(string)),   //4
                            };

        //Color of Person's Control elements which depend on the selected MenuItem  
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

            numUpDownHourStart.Value = 9;
            numUpDownMinuteStart.Value = 0;
            numUpDownHourEnd.Value = 18;
            numUpDownMinuteEnd.Value = 0;

            PersonOrGroupItem.Text = "Работа с одной персоной";
            toolTip1.SetToolTip(textBoxGroup, "Создать группу");
            toolTip1.SetToolTip(textBoxGroupDescription, "Изменить описание группы");
            StatusLabel2.Text = "";

            TryMakeDB();
            UpdateTableOfDB();
            SetTechInfoIntoDB();
            BoldAnualDates();

            //read last saved parameters from db and Registry and set their into variables
            LoadPrevioslySavedParameters();
            sServer1 = sServer1Registry.Length > 0 ? sServer1Registry : sServer1DB;
            sServer1UserName = sServer1UserNameRegistry.Length > 0 ? sServer1UserNameRegistry : sServer1UserNameDB;
            sServer1UserPassword = sServer1UserPasswordRegistry.Length > 0 ? sServer1UserPasswordRegistry : sServer1UserPasswordDB;
            mailServer = mailServerRegistry.Length > 0 ? mailServerRegistry : mailServerDB;
            mailServerUserName = mailServerUserNameRegistry.Length > 0 ? mailServerUserNameRegistry : mailServerUserNameDB;
            mailServerUserPassword = mailServerUserPasswordRegistry.Length > 0 ? mailServerUserPasswordRegistry : mailServerUserPasswordDB;


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
            TableModeItem.Visible = false;
            VisualModeItem.Visible = false;
            VisualSelectColorMenuItem.Visible = false;
            TableExportToExcelItem.Visible = false;
            listFioItem.Visible = false;

            string sFIO = "";
            try
            {
                comboBoxFio.SelectedIndex = comboBoxFio.FindString(sFIO); //ищем в комбобокс-e выбранный ФИО и устанавливаем на него фокус
                if (comboBoxFio.FindString(sFIO) != -1 && ShortFIO(sFIO).Length > 3)
                { StatusLabel2.Text = @"Выбран: " + ShortFIO(sFIO) + @" |  Всего ФИО: " + iFIO; }
                else if (ShortFIO(sFIO).Length < 3 && iFIO > 0)
                { StatusLabel2.Text = @"Всего ФИО: " + iFIO; }
            } catch { StatusLabel2.Text = " Начните работу с кнопки - \"Получить ФИО\""; }


            //Prepare Datatables
            dtPeople.Columns.AddRange(dcPeople);
            dtPeople.DefaultView.Sort = "[Группа] ASC, [Фамилия Имя Отчество] ASC, [Дата регистрации] ASC, [Время регистрации] ASC, [Реальное время прихода ЧЧ:ММ] ASC, [Реальное время прихода ЧЧ:ММ] ASC";


            //Clone default collumn name and structure from 'dtPeople' to other DataTables
            dtPersonTemp = dtPeople.Clone();  //Copy only structure(Name of collumns)
            dtPersonRegisteredFull = dtPeople.Clone();  //Copy only structure(Name of collumns)
            dtPersonRegistered = dtPeople.Clone();  //Copy only structure(Name of collumns)
            dtPersonGroup = dtPeople.Clone();  //Copy only structure(Name of collumns)
            dtPersonsLastList = dtPeople.Clone();  //Copy only structure(Name of collumns)
            dtPersonsLastComboList = dtPeople.Clone();  //Copy only structure(Name of collumns)

            dataGridView1.ShowCellToolTips = true;

            dtGroup.Columns.AddRange(dcGroup);
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

        private void TryMakeDB()
        {
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'PersonRegisteredFull' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, FIO TEXT, NAV TEXT, iDCard TEXT, DateRegistered TEXT, " +
                    "HourComming TEXT, MinuteComming TEXT, Comming REAL, HourControlling TEXT, MinuteControlling TEXT, Controlling REAL," +
                    " HourControllingOut TEXT, MinuteControllingOut TEXT, ControllingOut REAL,  WorkedOut REAL, ServerOfRegistration TEXT, Late TEXT, Early TEXT,  Vacancy TEXT, BusinesTrip TEXT, Reserv1 TEXT, Reserv2 TEXT);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'PersonRegistered' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, FIO TEXT, NAV TEXT, iDCard TEXT, DateRegistered TEXT, " +
                    "HourComming TEXT, MinuteComming TEXT, Comming REAL, HourControlling TEXT, MinuteControlling TEXT, Controlling REAL, " +
                    "HourControllingOut TEXT, MinuteControllingOut TEXT, ControllingOut REAL,  WorkedOut REAL, ServerOfRegistration TEXT, Late TEXT, Early TEXT,  Vacancy TEXT, BusinesTrip TEXT, Reserv1 TEXT, Reserv2 TEXT);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'PersonTemp' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, FIO TEXT, NAV TEXT, iDCard TEXT, DateRegistered TEXT, " +
                    "HourComming TEXT, MinuteComming TEXT, Comming REAL, HourControlling TEXT, MinuteControlling TEXT, Controlling REAL, HourControllingOut TEXT, " +
                    "MinuteControllingOut TEXT, ControllingOut REAL,  WorkedOut REAL, ServerOfRegistration TEXT, Late TEXT, Early TEXT,  Vacancy TEXT, BusinesTrip TEXT, Reserv1 TEXT, Reserv2 TEXT);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'PersonGroup' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, FIO TEXT, NAV TEXT, GroupPerson TEXT, " +
                    "HourControlling TEXT, MinuteControlling TEXT, Controlling REAL, ControllingHHMM TEXT, HourControllingOut TEXT, MinuteControllingOut TEXT, ControllingOut REAL, ControllingOUTHHMM TEXT, Late TEXT, Early TEXT,  Vacancy TEXT, BusinesTrip TEXT, " +
                    "Reserv1 TEXT, Reserv2 TEXT, UNIQUE ('FIO', 'NAV', 'GroupPerson') ON CONFLICT REPLACE);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'PersonGroupDesciption' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, GroupPerson TEXT, GroupPersonDescription TEXT, Late TEXT, Early TEXT,  Vacancy TEXT, BusinesTrip TEXT, Reserv1 TEXT, Reserv2 TEXT, " +
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

            ExecuteSql("CREATE TABLE IF NOT EXISTS 'Mailing' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, SenderEmail TEXT, " +
                    "RecipientEmail TEXT, Schedule TEXT, TypeReport TEXT, Description TEXT, DateCreated TEXT, Reserv1 TEXT, Reserv2 TEXT);", databasePerson);
        }

        private void UpdateTableOfDB()
        {
            TryUpdateStructureSqlDB("PersonRegisteredFull", "FIO TEXT, NAV TEXT, iDCard TEXT, DateRegistered TEXT, " +
                    "HourComming TEXT, MinuteComming TEXT, Comming REAL, HourControlling TEXT, MinuteControlling TEXT, Controlling REAL," +
                    "HourControllingOut TEXT, MinuteControllingOut TEXT, ControllingOut REAL,  WorkedOut REAL, ServerOfRegistration TEXT, " +
                    "Late TEXT, Early TEXT,  Vacancy TEXT, BusinesTrip TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);

            TryUpdateStructureSqlDB("PersonRegistered", "FIO TEXT, NAV TEXT, iDCard TEXT, DateRegistered TEXT, " +
                    "HourComming TEXT, MinuteComming TEXT, Comming REAL, HourControlling TEXT, MinuteControlling TEXT, Controlling REAL, " +
                    "HourControllingOut TEXT, MinuteControllingOut TEXT, ControllingOut REAL,  WorkedOut REAL, ServerOfRegistration TEXT, " +
                    "Late TEXT, Early TEXT,  Vacancy TEXT, BusinesTrip TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
            
            TryUpdateStructureSqlDB("PersonTemp", "FIO TEXT, NAV TEXT, iDCard TEXT, DateRegistered TEXT, " +
                    "HourComming TEXT, MinuteComming TEXT, Comming REAL, HourControlling TEXT, MinuteControlling TEXT, Controlling REAL, " +
                    "HourControllingOut TEXT, MinuteControllingOut TEXT, ControllingOut REAL,  WorkedOut REAL, ServerOfRegistration TEXT, " +
                    "Late TEXT, Early TEXT,  Vacancy TEXT, BusinesTrip TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);

            TryUpdateStructureSqlDB("PersonGroup", "FIO TEXT, NAV TEXT, GroupPerson TEXT, HourControlling TEXT, MinuteControlling TEXT, Controlling REAL, " +
                    "HourControllingOut TEXT, MinuteControllingOut TEXT, ControllingOut REAL, ControllingHHMM TEXT, ControllingOUTHHMM TEXT, Late TEXT, Early TEXT,  Vacancy TEXT, BusinesTrip TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);

            TryUpdateStructureSqlDB("PersonGroupDesciption", "GroupPerson TEXT, GroupPersonDescription TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
            TryUpdateStructureSqlDB("TechnicalInfo", "PCName TEXT, POName TEXT, POVersion TEXT, LastDateStarted TEXT, Reserv1 TEXT, Reserv2 TEXT, Reserverd3 TEXT", databasePerson);
            TryUpdateStructureSqlDB("BoldedDates", "BoldedDate TEXT, NAV TEXT, Groups TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
            TryUpdateStructureSqlDB("MySettings", "MyParameterName TEXT, MyParameterValue TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
            TryUpdateStructureSqlDB("ProgramSettings", " PoParameterName TEXT, PoParameterValue TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
            TryUpdateStructureSqlDB("EquipmentSettings", "EquipmentParameterName TEXT, EquipmentParameterValue TEXT, EquipmentParameterServer TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
            TryUpdateStructureSqlDB("PersonsLastList", "PersonsList TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
            TryUpdateStructureSqlDB("PersonsLastComboList", "ComboList TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
            TryUpdateStructureSqlDB("Mailing", "SenderEmail TEXT, RecipientEmail TEXT, Schedule TEXT, TypeReport TEXT, Description TEXT, DateCreated TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
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

        private void LoadPrevioslySavedParameters()   //Select Previous Data from DB and write it into the combobox and Parameters
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
                                        if (record["EquipmentParameterValue"].ToString().Trim() == "SKDServer" && record["EquipmentParameterName"].ToString().Trim() == "SKDUser")
                                        {
                                            sServer1DB = record["EquipmentParameterServer"].ToString();
                                            try { sServer1UserNameDB = DecryptBase64ToString(record["Reserv1"].ToString(), btsMess1, btsMess2); } catch { }
                                            try { sServer1UserPasswordDB = DecryptBase64ToString(record["Reserv2"].ToString(), btsMess1, btsMess2); } catch { }
                                        }
                                        if (record["EquipmentParameterValue"].ToString().Trim() == "MailServer" && record["EquipmentParameterName"].ToString().Trim() == "MailUser")
                                        {
                                            mailServerDB = record["EquipmentParameterServer"].ToString();
                                            mailServerUserNameDB = record["Reserv1"].ToString();
                                            try { mailServerUserPasswordDB = DecryptBase64ToString(record["Reserv2"].ToString(), btsMess1, btsMess2); } catch { }
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

            using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(myRegKey, Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree, System.Security.AccessControl.RegistryRights.ReadKey))
            {
                try { sServer1Registry = EvUserKey.GetValue("SKDServer").ToString().Trim(); } catch { }
                try { sServer1UserNameRegistry = DecryptBase64ToString(EvUserKey.GetValue("SKDUser").ToString(), btsMess1, btsMess2).ToString().Trim(); } catch { }
                try { sServer1UserPasswordRegistry = DecryptBase64ToString(EvUserKey.GetValue("SKDUserPassword").ToString(), btsMess1, btsMess2).ToString().Trim(); } catch { }

                try { mailServerRegistry = EvUserKey.GetValue("MailServer").ToString().Trim(); } catch { }
                try { mailServerUserNameRegistry = EvUserKey.GetValue("MailUser").ToString().Trim(); } catch { }
                try { mailServerUserPasswordRegistry = DecryptBase64ToString(EvUserKey.GetValue("MailUserPassword").ToString(), btsMess1, btsMess2).ToString().Trim(); } catch { }
            }
        }

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
                } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
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
            DataTable dt = new DataTable();
            if (databasePerson.Exists)
            {
                // dtPeople.Clear();
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    using (var sqlDA = new SQLiteDataAdapter(mySqlQuery + " FROM '" + myTable + "' " + mySqlWhere + "; ", sqlConnection))
                    {
                        dt = new DataTable();
                        //dtPeople 
                        sqlDA.Fill(dt);
                    }
                }
            }
            _dataGridViewSource(dt);
            iCounterLine = _dataGridView1RowsCount();
            sLastSelectedElement = "dataGridView";
        }

        private void ShowDatatableOnDatagridview(DataTable dt, string[] nameHidenCollumnsArray) //Query data from the Table of the DB
        {
            DataTable dataTable = dt.Copy();
            for (int i = 0; i < nameHidenCollumnsArray.Length; i++)
            {
                if (nameHidenCollumnsArray[i] != null && nameHidenCollumnsArray[i].Length > 0)
                    try { dataTable.Columns[nameHidenCollumnsArray[i]].ColumnMapping = MappingType.Hidden; } catch { }
            }

            _dataGridViewSource(dataTable);
            _toolStripStatusLabelSetText(StatusLabel2, "Всего записей: " + _dataGridView1RowsCount());
            
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

        private void DeleteDataTableQueryParameters(System.IO.FileInfo databasePerson, string myTable, string mySqlParameter1, string mySqlData1, string mySqlParameter2 = "", string mySqlData2 = "") //Delete data from the Table of the DB by NAV (both parameters are string)
        {
            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();

                    if (mySqlParameter1.Length > 0 && mySqlParameter2.Length > 0)
                    {
                        using (var sqlCommand = new SQLiteCommand("DELETE FROM '" + myTable + "' Where " + mySqlParameter1 + "= @" + mySqlParameter1 + " AND " + mySqlParameter2 + "= @" + mySqlParameter2 + ";", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@" + mySqlParameter1, DbType.String).Value = mySqlData1;
                            sqlCommand.Parameters.Add("@" + mySqlParameter2, DbType.String).Value = mySqlData2;

                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }
                    }
                    else if (mySqlParameter1.Length > 0)
                    {
                        using (var sqlCommand = new SQLiteCommand("DELETE FROM '" + myTable + "' Where " + mySqlParameter1 + "= @" + mySqlParameter1 + ";", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@" + mySqlParameter1, DbType.String).Value = mySqlData1;
                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }
                    }
                    using (var sqlCommand = new SQLiteCommand("vacuum;", sqlConnection))   //vacuum DB
                    { try { sqlCommand.ExecuteNonQuery(); } catch { } }
                    sqlConnection.Close();
                }
            }
            myTable = null; mySqlParameter1 = null; mySqlData1 = null;
        }

        private void CheckAliveServer(string serverName, string userName, string userPasswords) //Get the list of registered users
        {
            bServer1Exist = false;
            string stringConnection;
            // _toolStripStatusLabelSetText(StatusLabel2, "Проверка доступности " + serverName + ". Ждите окончания процесса...");
            stimerPrev = "Проверка доступности " + serverName + ". Ждите окончания процесса...";
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
            } catch
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
                _MenuItemEnabled(VisualModeItem, false);
            }

            stringConnection = null;
        }


        private async void GetFio_Click(object sender, EventArgs e)
        {
            CheckBoxesFiltersAll_CheckedState(false);
            CheckBoxesFiltersAll_Enable(false);
            _MenuItemEnabled(QuickLoadDataItem, false);
            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(SettingsMenuItem, false);
            _MenuItemEnabled(AnualDatesMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            _MenuItemEnabled(GetFioItem, false);

            DataTable dtTempIntermediate = dtPeople.Clone();

            await Task.Run(() => CheckAliveServer(sServer1, sServer1UserName, sServer1UserPassword));

            if (bServer1Exist)
            {
                _ProgressBar1Value0();
                _timer1Enabled(true);
                dataGridView1.Visible = true;
                pictureBox1.Visible = false;

                await Task.Run(() => GetFioFromServers(dtTempIntermediate));

                string[] arrayHiddenCollumns =
                    {
                            //@"Группа",                  //3
                            @"Время прихода,часы",       //4
                            @"Время прихода,минут",      //5
                            @"Время прихода",            //6
                            @"Время ухода,часы",         //7
                            @"Время ухода,минут",        //8
                            @"Время ухода",              //9
                            @"№ пропуска",               //10
                            @"Дата регистрации",         //12
                            @"Время регистрации,часы",   //13
                            @"Время регистрации,минут",  //14
                            @"Время регистрации",        //15
                            @"Реальное время ухода,часы",//16
                            @"Реальное время ухода,минут",//17
                            @"Реальное время ухода",     //18
                            @"Сервер СКД",               //19
                            @"Имя точки прохода",        //20
                            @"Направление прохода",      //21
                           // @"Время прихода ЧЧ:ММ",      //22
                           // @"Время ухода ЧЧ:ММ",        //23
                            @"Реальное время прихода ЧЧ:ММ",//24
                            @"Реальное время ухода ЧЧ:ММ", //25
                            @"Реальное отработанное время", //26
                            @"Реальное отработанное время ЧЧ:ММ", //27
                            @"Опоздание",                   //28
                            @"Ранний уход",              //29
                            @"Отпуск (отгул)",           //30
                                  @"Коммандировка",                 //31
                                  @"День недели",                    //32
                                  @"Больничный",                    //33
                                  @"Согласованное отсутствие",      //34
                                  @"Код",                           //35
                                  @"Вышестоящая группа",            //36
                                  @"Описание группы"                //37
                      };

                await Task.Run(() => ImportTablePeopleToTableGroupsInLocalDB(databasePerson.ToString(), dtTempIntermediate));
                
                //show selected data     
                //distinct Records                
                var namesDistinctCollumnsArray = arrayAllCollumnsDataTablePeople.Except(arrayHiddenCollumns).ToArray(); //take distinct data
                dtPeople = GetDistinctRecords(dtTempIntermediate, namesDistinctCollumnsArray);


                await Task.Run(() => ShowDatatableOnDatagridview(dtPeople, arrayHiddenCollumns));

                await Task.Run(() => panelViewResize(numberPeopleInLoading));
                listFioItem.Visible = true;
            }
            else { GetInfoSetup(); }

            _MenuItemEnabled(SettingsMenuItem, true);
        }

        private void GetFioFromServers(DataTable dataTablePeopple) //Get the list of registered users
        {
            Person personFromServer = new Person();
            //  dataTablePeopple.Dispose();
            //  dtGroup.Dispose();
            dataTablePeopple.Clear();
            dtGroup.Clear();
            iFIO = 0;

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

                    string query = "SELECT id, name, surname, patronymic, post, tabnum, parent_id FROM OBJ_PERSON ";
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
                                        iFIO++;
                                        DataRow row = dataTablePeopple.NewRow();

                                        personFromServer = new Person();
                                        personFromServer.FIO = record["name"].ToString().Trim() + " " + record["surname"].ToString().Trim() + " " + record["patronymic"].ToString().Trim();
                                        personFromServer.NAV = record["tabnum"].ToString().Trim();
                                        personFromServer.idCard = Convert.ToInt32(record["id"].ToString().Trim());
                                        personFromServer.PositionInDepartment = record["post"].ToString().Trim();
                                        personFromServer.Department = record["parent_id"].ToString().Trim();
                                        personFromServer.GroupPerson = record["parent_id"].ToString().Trim();

                                        personFromServer.ControlInHour = _numUpDownReturn(numUpDownHourStart).ToString();
                                        personFromServer.ControlInHourDecimal = _numUpDownReturn(numUpDownHourStart);
                                        personFromServer.ControlInMinute = _numUpDownReturn(numUpDownMinuteStart).ToString();
                                        personFromServer.ControlInMinuteDecimal = _numUpDownReturn(numUpDownMinuteStart);
                                        personFromServer.ControlInHHMM = ConvertStringsTimeToStringHHMM(personFromServer.ControlInHour, personFromServer.ControlInMinute);

                                        personFromServer.ControlOutHour = _numUpDownReturn(numUpDownHourEnd).ToString();
                                        personFromServer.ControlOutHourDecimal = _numUpDownReturn(numUpDownHourEnd);
                                        personFromServer.ControlOutMinute = _numUpDownReturn(numUpDownMinuteEnd).ToString();
                                        personFromServer.ControlOutMinuteDecimal = _numUpDownReturn(numUpDownMinuteEnd);
                                        personFromServer.ControlOutHHMM = ConvertStringsTimeToStringHHMM(personFromServer.ControlOutHour, personFromServer.ControlOutMinute);

                                        row[0] = iFIO;
                                        row[1] = personFromServer.FIO;
                                        row[2] = personFromServer.NAV;
                                        row[3] = personFromServer.Department;
                                        row[4] = personFromServer.ControlInHour;
                                        row[5] = personFromServer.ControlInMinute;
                                        row[6] = ConvertStringsTimeToDecimal(personFromServer.ControlInHour, personFromServer.ControlInMinute);
                                        row[7] = personFromServer.ControlOutHour;
                                        row[8] = personFromServer.ControlOutMinute;
                                        row[6] = ConvertStringsTimeToDecimal(personFromServer.ControlOutHour, personFromServer.ControlOutMinute);
                                        row[22] = ConvertStringsTimeToStringHHMM(personFromServer.ControlInHour, personFromServer.ControlInMinute);
                                        row[23] = ConvertStringsTimeToStringHHMM(personFromServer.ControlOutHour, personFromServer.ControlOutMinute);
                                        row[10] = personFromServer.idCard;

                                        dataTablePeopple.Rows.Add(row);

                                        listFIO.Add(record["name"].ToString().Trim() + "|" + record["surname"].ToString().Trim() + "|" + record["patronymic"].ToString().Trim() + "|" + record["id"].ToString().Trim() + "|" +
                                                    record["tabnum"].ToString().Trim() + "|" + sServer1);
                                        ListFIOTemp.Add(record["name"].ToString().Trim() + " " + record["surname"].ToString().Trim() + " " + record["patronymic"].ToString().Trim() + "|" + record["tabnum"].ToString().Trim());
                                    }
                                } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                                _ProgressWork1();
                            }
                        }
                    }

                    //получить список департаментов с сервера
                    query = "SELECT id,level_id,name,owner_id,parent_id,region_id,schedule_id  FROM OBJ_DEPARTMENT";
                    using (var sqlCommand = new SqlCommand(query, sqlConnection))
                    {
                        using (var sqlReader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in sqlReader)
                            {
                                _ProgressWork1();
                                try
                                {
                                    if (record?["name"].ToString().Trim().Length > 0)
                                    {
                                        DataRow row = dtGroup.NewRow();
                                        row[2] = record["id"].ToString().Trim();
                                        row[4] = record["name"].ToString().Trim();
                                        row[3] = record["parent_id"].ToString().Trim();

                                        dtGroup.Rows.Add(row);
                                    }
                                } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                                _ProgressWork1();
                            }
                        }
                    }
                }

                _toolStripStatusLabelSetText(StatusLabel2, "Список ФИО успешно получен");
                stimerPrev = "Все списки с ФИО с сервера СКД успешно получены";
            } catch (Exception Expt)
            {
                bServer1Exist = false;
                stimerPrev = "Сервер не доступен или неправильная авторизация";
                MessageBox.Show(Expt.ToString(), @"Сервер не доступен или неправильная авторизация", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //Remove dublicat and output Result into combobox1
            IEnumerable<string> ListFIOCombo = ListFIOTemp.Distinct();

            if (databasePerson.Exists && bServer1Exist)
            {
                DeleteAllDataInTableQuery(databasePerson, "PersonsLastList");
                DeleteAllDataInTableQuery(databasePerson, "PersonsLastComboList");
                foreach (var dr in dtGroup.AsEnumerable())
                {
                    DeleteDataTableQueryParameters(databasePerson, "PersonGroup", "GroupPerson", dr[2].ToString());
                    DeleteDataTableQueryParameters(databasePerson, "PersonGroupDesciption", "GroupPerson", dr[2].ToString());
                }
                foreach (var dr in dtGroup.AsEnumerable())
                {
                  CreateGroupInDB  (databasePerson, dr[2].ToString(), dr[4].ToString());
                }


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
                    sqlConnection.Close();
                }


                foreach (string str in ListFIOCombo.ToArray())
                { _comboBoxFioAdd(str); }
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
                _toolStripStatusLabelSetText(StatusLabel2, "Ошибка доступа к БД СКД-сервера");
                _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                _MenuItemEnabled(QuickLoadDataItem, false);
                MessageBox.Show("Проверьте правильность написания серверов,\nимя и пароль sa-администратора,\nа а также доступность серверов и их баз!");
            }
            _ProgressBar1Value100();
            _MenuItemEnabled(GetFioItem, true);
            _MenuItemEnabled(FunctionMenuItem, true);
            _MenuItemEnabled(SettingsMenuItem, true);
            _MenuItemEnabled(AnualDatesMenuItem, true);
            _MenuItemEnabled(GroupsMenuItem, true);
        }

        private void listFioItem_Click(object sender, EventArgs e) //ListFioReturn()
        { ListFioReturn(); }

        private async void ListFioReturn()
        {
            string[] arrayHiddenCollumns =
                    {
                            //@"Группа",                  //3
                            @"Время прихода,часы",       //4
                            @"Время прихода,минут",      //5
                            @"Время прихода",            //6
                            @"Время ухода,часы",         //7
                            @"Время ухода,минут",        //8
                            @"Время ухода",              //9
                            @"№ пропуска",               //10
                            @"Дата регистрации",         //12
                            @"Время регистрации,часы",   //13
                            @"Время регистрации,минут",  //14
                            @"Время регистрации",        //15
                            @"Реальное время ухода,часы",//16
                            @"Реальное время ухода,минут",//17
                            @"Реальное время ухода",     //18
                            @"Сервер СКД",               //19
                            @"Имя точки прохода",        //20
                            @"Направление прохода",      //21
                           // @"Время прихода ЧЧ:ММ",      //22
                           // @"Время ухода ЧЧ:ММ",        //23
                            @"Реальное время прихода ЧЧ:ММ",//24
                            @"Реальное время ухода ЧЧ:ММ", //25
                            @"Реальное отработанное время", //26
                            @"Реальное отработанное время ЧЧ:ММ", //27
                            @"Опоздание",                   //28
                            @"Ранний уход",              //29
                            @"Отпуск (отгул)",           //30
                                  @"Коммандировка",                 //31
                                  @"День недели",                    //32
                                  @"Больничный",                    //33
                                  @"Согласованное отсутствие",      //34
                                  @"Код",                           //35
                                  @"Вышестоящая группа",            //36
                                  @"Описание группы"                //37
                      };

            await Task.Run(() => ShowDatatableOnDatagridview(dtPeople, arrayHiddenCollumns));

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
            _MenuItemEnabled(SettingsMenuItem, false);
            _MenuItemEnabled(AnualDatesMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            _MenuItemEnabled(SettingsMenuItem, false);
            //_MenuItemEnabled(ViewMenuItem, false);
            _MenuItemEnabled(VisualModeItem, false);
            _MenuItemEnabled(TableModeItem, false);
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
            _timer1Enabled(false);

            _ChangeMenuItemBackColor(TableExportToExcelItem, SystemColors.Control);
            stimerPrev = "";
            _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);
            sLastSelectedElement = "ExportExcel";
            iDGCollumns = 0; iDGRows = 0;
            _toolStripStatusLabelSetText(StatusLabel2, "Готово!");
            _MenuItemEnabled(QuickLoadDataItem, true);
            _MenuItemEnabled(FunctionMenuItem, true);
            _MenuItemEnabled(SettingsMenuItem, true);
            _MenuItemEnabled(AnualDatesMenuItem, true);
            _MenuItemEnabled(GroupsMenuItem, true);
            _MenuItemEnabled(SettingsMenuItem, true);
            _MenuItemEnabled(VisualModeItem, true);
            _MenuItemEnabled(TableModeItem, true);
            _controlEnable(dataGridView1, true);
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

        private async void Export_Click(object sender, EventArgs e)
        {
            _timer1Enabled(true);
            _toolStripStatusLabelSetText(StatusLabel2, "Генерирую Excel-файл");
            stimerPrev = "Наполняю файл данными из текущей таблицы";

            await Task.Run(() => ExportDatagridToExcel());
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        { SelectFioAndNavFromCombobox(); }

        private void SelectFioAndNavFromCombobox()
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
                textBoxFIO.Text = Regex.Split(sComboboxFIO, "[|]")[0].Trim();
                textBoxNav.Text = Regex.Split(sComboboxFIO, "[|]")[1].Trim();
                StatusLabel2.Text = @"Выбран: " + ShortFIO(textBoxFIO.Text) + @" |  Всего ФИО: " + iFIO;
            } catch { }
            if (comboBoxFio.SelectedIndex > -1)
            {
                QuickLoadDataItem.BackColor = Color.PaleGreen;
                groupBoxPeriod.BackColor = Color.PaleGreen;
                groupBoxTimeStart.BackColor = Color.PaleGreen;
                groupBoxTimeEnd.BackColor = Color.PaleGreen;
                groupBoxRemoveDays.BackColor = SystemColors.Control;
            }
            sComboboxFIO = null;
            nameOfLastTableFromDB = "PersonRegistered";
        }

        private void CreateGroupItem_Click(object sender, EventArgs e)
        {
            if (textBoxGroup.Text.Trim().Length > 0)
            {
                CreateGroupInDB(databasePerson, textBoxGroup.Text.Trim(), textBoxGroupDescription.Text.Trim());
            }

            PersonOrGroupItem.Text = "Работа с одной персоной";
            nameOfLastTableFromDB = "PersonGroup";
            ListGroups();
        }

        private void CreateGroupInDB(System.IO.FileInfo fileInfo,  string nameGroup, string descriptionGroup)
        {
            if (nameGroup.Length > 0)
            {
                using (var connection = new SQLiteConnection($"Data Source={fileInfo};Version=3;"))
                {
                    connection.Open();
                    using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PersonGroupDesciption' (GroupPerson, GroupPersonDescription) " +
                                            "VALUES (@GroupPerson, @GroupPersonDescription )", connection))
                    {
                        command.Parameters.Add("@GroupPerson", DbType.String).Value = nameGroup;
                        command.Parameters.Add("@GroupPersonDescription", DbType.String).Value = descriptionGroup;
                        try { command.ExecuteNonQuery(); } catch { }
                    }
                    connection.Close();
                }
                _toolStripStatusLabelSetText( StatusLabel2, "Группа - \"" + nameGroup + "\" создана"); 
            }

        }

        private void ListGroupsItem_Click(object sender, EventArgs e)
        { ListGroups(); }

        private void ListGroups()
        {
            groupBoxProperties.Visible = false;
            dataGridView1.Visible = false;

            ShowDataTableQuery(databasePerson, "PersonGroupDesciption", "SELECT GroupPerson AS 'Группа', GroupPersonDescription AS 'Описание группы', Reserv1 AS 'Колличество в группе' ", " group by GroupPerson ORDER BY GroupPerson asc; ");

            try
            {
                textBoxGroup.Text = dataGridView1.Rows[0].Cells[0].Value.ToString();
                textBoxGroupDescription.Text = dataGridView1.Rows[0].Cells[1].Value.ToString();
                StatusLabel2.Text = @"Выбрана группа: " + textBoxGroup.Text + @" |  Всего групп: " + _dataGridView1RowsCount();

                QuickLoadDataItem.BackColor = Color.PaleGreen;
                groupBoxPeriod.BackColor = Color.PaleGreen;
                groupBoxTimeStart.BackColor = Color.PaleGreen;
                groupBoxTimeEnd.BackColor = Color.PaleGreen;
                groupBoxRemoveDays.BackColor = SystemColors.Control;
            } catch { }

            DeleteGroupItem.Visible = true;
            dataGridView1.Visible = true;
            MembersGroupItem.Enabled = true;
            nameOfLastTableFromDB = "PersonGroupDesciption";
            PersonOrGroupItem.Text = "Работа с одной персоной";
        }

        private void MembersGroupItem_Click(object sender, EventArgs e)//SearchMembersSelectedGroup()
        { SearchMembersSelectedGroup(); }

        private void SearchMembersSelectedGroup()
        {
            if (nameOfLastTableFromDB == "PersonGroup" || nameOfLastTableFromDB == "PersonGroupDesciption")
            {
                int IndexCurrentRow = _dataGridView1CurrentRowIndex();
                string nameGroup = DefinyGroupNameByIndexRowDatagridview(IndexCurrentRow);
                SeekAndShowMembersOfGroup(nameGroup);
            }
        }

        private string DefinyGroupNameByIndexRowDatagridview(int IndexCurrentRow)
        {
            string nameFoundGroup = @"%%";
            try
            {
                if (0 < dataGridView1.Rows.Count && IndexCurrentRow < dataGridView1.Rows.Count)
                {
                    if (nameOfLastTableFromDB == "PersonGroup" || nameOfLastTableFromDB == "PersonGroupDesciption")
                    {
                        int IndexColumn1 = -1;
                        int IndexColumn2 = -1;

                        for (int i = 0; i < dataGridView1.ColumnCount; i++)
                        {
                            if (dataGridView1.Columns[i].HeaderText.ToString() == "GroupPerson" || dataGridView1.Columns[i].HeaderText.ToString() == "Группа")
                            { IndexColumn1 = i; }
                            else if (dataGridView1.Columns[i].HeaderText.ToString() == "GroupPersonDescription" || dataGridView1.Columns[i].HeaderText.ToString() == "Описание группа")
                            { IndexColumn2 = i; }
                        }
                        if (IndexColumn1 > -1 || IndexColumn2 > -1)
                        {
                            nameFoundGroup = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString();
                        }
                    }
                }
            } catch { }
            return nameFoundGroup;
        }

        private void SeekAndShowMembersOfGroup(string nameGroup)
        {
            dtPersonTemp.Dispose();
            dtPersonTemp = dtPeople.Clone();
            var dtTemp = dtPeople.Clone();

            numberPeopleInLoading = 0;
            DataRow dataRow;
            using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
            {
                sqlConnection.Open();
                using (var sqlCommand = new SQLiteCommand(
                    "Select * FROM PersonGroup where GroupPerson like '" + nameGroup + "';", sqlConnection))
                {
                    using (var sqlReader = sqlCommand.ExecuteReader())
                    {
                        string d1 = "", d2 = "", d3 = "9", d4 = "0", d13 = "18", d14 = "0";

                        foreach (DbDataRecord record in sqlReader)
                        {
                            try { d1 = record["FIO"].ToString().Trim(); } catch { }

                            if (record != null && d1.Length > 0)
                            {
                                dataRow = dtTemp.NewRow();

                                try { d2 = record["NAV"].ToString().Trim(); } catch { }
                                try { d3 = record["HourControlling"].ToString().Trim(); } catch { }
                                try { d4 = record["MinuteControlling"].ToString().Trim(); } catch { }
                                try { d13 = record["HourControllingOut"].ToString().Trim(); } catch { }
                                try { d14 = record["MinuteControllingOut"].ToString().Trim(); } catch { }
                                //HourControllingOut TEXT, MinuteControllingOut TEXT
                                //FIO|NAV|H|M

                                dataRow[@"Фамилия Имя Отчество"] = d1;
                                dataRow[@"NAV-код"] = d2;
                                dataRow[@"Группа"] = nameGroup;
                                dataRow[@"Время прихода,часы"] = d3;
                                dataRow[@"Время прихода,минут"] = d4;
                                dataRow[@"Время прихода"] = ConvertStringsTimeToDecimal(d3, d4);
                                dataRow[@"Время ухода,часы"] = d13;
                                dataRow[@"Время ухода,минут"] = d14;
                                dataRow[@"Время ухода"] = ConvertStringsTimeToDecimal(d13, d14);
                                dataRow[@"Отдел"] = nameGroup;
                                dataRow[@"Время прихода ЧЧ:ММ"] = ConvertStringsTimeToStringHHMM(d3, d4);
                                dataRow[@"Время ухода ЧЧ:ММ"] = ConvertStringsTimeToStringHHMM(d13, d14);

                                dtTemp.Rows.Add(dataRow);
                                numberPeopleInLoading++;
                            }
                        }
                        d1 = null; d2 = null; d3 = null; d4 = null; d13 = null; d14 = null;
                    }
                }
            }

            string[] nameHidenCollumnsArray =
            {
                                  @"№ п/п",//0
                                  @"Время прихода,часы",//4
                                  @"Время прихода,минут", //5
                                  @"Время прихода",//6
                                  @"Время ухода,часы",//7
                                  @"Время ухода,минут",//8
                                  @"Время ухода",//9
                                  @"№ пропуска", //10
                                  //@"Отдел",//11
                                  @"Дата регистрации",//12
                                  @"Время регистрации,часы",//13
                                  @"Время регистрации,минут",//14
                                  @"Время регистрации", //15
                                  @"Реальное время ухода,часы",//16
                                  @"Реальное время ухода,минут",//17
                                  @"Реальное время ухода", //18
                                  @"Сервер СКД", //19
                                  @"Имя точки прохода", //20
                                  @"Направление прохода", //21
                                  @"Реальное время прихода ЧЧ:ММ",//24
                                  @"Реальное время ухода ЧЧ:ММ", //25
                                  @"Реальное отработанное время", //26
                                  @"Реальное отработанное время ЧЧ:ММ", //27
                                  @"Опоздание",                    //28
                                  @"Ранний уход",                 //29
                                  @"Отпуск (отгул)",                 //30
                                  @"Коммандировка",                 //31
                                  @"День недели",                    //32
                                  @"Больничный",                    //33
                                  @"Согласованное отсутствие",      //34
                                  @"Код",                           //35
                                  @"Вышестоящая группа",            //36
                                  @"Описание группы"                //37
                        };

            var namesDistinctCollumnsArray = arrayAllCollumnsDataTablePeople.Except(nameHidenCollumnsArray).ToArray(); //take distinct data
            dtPersonTemp = GetDistinctRecords(dtTemp, namesDistinctCollumnsArray);

            ShowDatatableOnDatagridview(dtPersonTemp, nameHidenCollumnsArray);
            _MenuItemVisible(DeletePersonFromGroupItem, true);
            nameOfLastTableFromDB = "PersonGroup";
            dtTemp.Dispose();
        }


        private void importPeopleInLocalDBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ImportTextToTable(dtPersonGroup);
            ImportTablePeopleToTableGroupsInLocalDB(databasePerson.ToString(), dtPersonGroup);
            ImportListGroupsDescriptionInLocalDB(databasePerson.ToString(), listGroups);
        }

        private void ImportTextToTable(DataTable dt) //Fill dtPeople
        {
            List<string> listRows = LoadDataIntoList();
            listGroups = new List<string>();
            string checkHourS;
            string checkHourE;

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
                    DataRow row = dt.NewRow();

                    foreach (string strRow in listRows)
                    {

                        string[] cell = strRow.Split('\t');
                        if (cell.Length == 7)
                        {
                            row[0] = cell[0];
                            row[1] = cell[1];
                            row[2] = cell[2];
                            listGroups.Add(cell[2]);

                            checkHourS = cell[3];
                            if (TryParseStringToDecimal(checkHourS) == 0)
                            { checkHourS = numUpHourStart.ToString(); }
                            row[3] = checkHourS;
                            row[4] = cell[4];
                            row[5] = ConvertStringsTimeToDecimal(checkHourS, cell[4]);
                            row[22] = ConvertStringsTimeToStringHHMM(checkHourS, cell[4]);


                            checkHourE = cell[5];
                            if (TryParseStringToDecimal(checkHourE) == 0)
                            { checkHourE = numUpDownHourEnd.ToString(); }
                            row[6] = checkHourE;
                            row[7] = cell[6];
                            row[8] = ConvertStringsTimeToDecimal(checkHourE, cell[6]);
                            row[23] = ConvertStringsTimeToStringHHMM(checkHourE, cell[6]);

                            dt.Rows.Add(row);
                            row = dt.NewRow();
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("выбранный файл пустой, или \nне подходит для импорта.");
            }
        }

        private void ImportTablePeopleToTableGroupsInLocalDB(string pathToPersonDB, DataTable dtSource) //use listGroups
        {
            using (var connection = new SQLiteConnection($"Data Source={pathToPersonDB};Version=3;"))
            {
                connection.Open();

                //import groups
                SQLiteCommand commandTransaction = new SQLiteCommand("begin", connection);
                commandTransaction.ExecuteNonQuery();                
                using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PersonGroup' (FIO, NAV, GroupPerson, HourControlling, MinuteControlling, Controlling, HourControllingOut, MinuteControllingOut, ControllingOut, ControllingHHMM, ControllingOUTHHMM) " +
                                         "VALUES (@FIO, @NAV, @GroupPerson, @HourControlling, @MinuteControlling, @Controlling, @HourControllingOut, @MinuteControllingOut, @ControllingOut, @ControllingHHMM, @ControllingOUTHHMM)", connection))
                {
                    foreach (DataRow row in dtSource.Rows)
                    {
                        command.Parameters.Add("@FIO", DbType.String).Value = row[@"Фамилия Имя Отчество"].ToString(); //row[0]
                        command.Parameters.Add("@NAV", DbType.String).Value = row[@"NAV-код"].ToString();
                        command.Parameters.Add("@GroupPerson", DbType.String).Value = row[@"Группа"].ToString();
                        command.Parameters.Add("@HourControlling", DbType.String).Value = row[@"Время прихода,часы"].ToString();
                        command.Parameters.Add("@MinuteControlling", DbType.String).Value = row[@"Время прихода,минут"].ToString();
                        command.Parameters.Add("@Controlling", DbType.Decimal).Value = TryParseStringToDecimal(row[@"Время прихода"].ToString());
                        command.Parameters.Add("@HourControllingOut", DbType.String).Value = row[@"Время ухода,часы"].ToString();
                        command.Parameters.Add("@MinuteControllingOut", DbType.String).Value = row[@"Время ухода,минут"].ToString();
                        command.Parameters.Add("@ControllingOut", DbType.Decimal).Value = TryParseStringToDecimal(row[@"Время ухода"].ToString());
                        command.Parameters.Add("@ControllingHHMM", DbType.String).Value = row[@"Время прихода ЧЧ:ММ"].ToString();
                        command.Parameters.Add("@ControllingOUTHHMM", DbType.String).Value = row[@"Время ухода ЧЧ:ММ"].ToString();
                        try { command.ExecuteNonQuery(); } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                    }
                }
                commandTransaction = new SQLiteCommand("end", connection);
                commandTransaction.ExecuteNonQuery();
            }
        }

        private void ImportListGroupsDescriptionInLocalDB(string pathToPersonDB, List<string> groups) //use listGroups
        {
            using (var connection = new SQLiteConnection($"Data Source={pathToPersonDB};Version=3;"))
            {
                connection.Open();
                SQLiteCommand commandTransaction = new SQLiteCommand("begin", connection);
                commandTransaction.ExecuteNonQuery();
                using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PersonGroupDesciption' (GroupPerson, GroupPersonDescription) " +
                                        "VALUES (@GroupPerson, @GroupPersonDescription )", connection))
                {
                    foreach (string group in groups.Distinct())
                    {
                        if (group.Contains('|'))
                        {
                            command.Parameters.Add("@GroupPerson", DbType.String).Value = group.Split('|')[0];
                            command.Parameters.Add("@GroupPersonDescription", DbType.String).Value = group.Split('|')[1];
                        }
                        else
                        {
                            command.Parameters.Add("@GroupPerson", DbType.String).Value = group;
                            command.Parameters.Add("@GroupPersonDescription", DbType.String).Value = group;
                        }
                        try { command.ExecuteNonQuery(); } catch { }
                    }
                }
                commandTransaction = new SQLiteCommand("end", connection);
                commandTransaction.ExecuteNonQuery();
            }
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
                } catch (Exception expt) { MessageBox.Show("Error was happened on " + i + " row\n" + expt.ToString()); }
                if (i > listMaxLength - 10 || i == 0)
                {
                    MessageBox.Show("Error was happened on " + i + " row\n You've been chosen the long file!");
                }
            }
            return listValue;
        }

        private void AddPersonToGroupItem_Click(object sender, EventArgs e) //AddPersonToGroup() //Add the selected person into the named group
        { AddPersonToGroup(); }

        private void AddPersonToGroup() //Add the selected person into the named group
        {
            nameOfLastTableFromDB = "PersonGroup";

            SelectFioAndNavFromCombobox();

            string group = textBoxGroup.Text.Trim();
            string fio = textBoxFIO.Text;
            string nav = textBoxNav.Text;
            string[] timeIn = { "09", "00", "09:00" };
            string[] timeOut = { "18", "00", "18:00" };
            decimal[] timeInDecimal = { 9, 0, 09 };
            decimal[] timeOutDecimal = { 18, 0, 18 };

            int IndexCurrentRow = _dataGridView1CurrentRowIndex();
            int IndexColumn6 = -1;
            int IndexColumn7 = -1;

            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                if (dataGridView1.Columns[i].HeaderText.ToString() == "Время прихода ЧЧ:ММ" ||
                    dataGridView1.Columns[i].HeaderText.ToString() == "Контрольное время")
                    IndexColumn6 = i;
                else if (dataGridView1.Columns[i].HeaderText.ToString() == "Время ухода ЧЧ:ММ" ||
                    dataGridView1.Columns[i].HeaderText.ToString() == "Уход с работы")
                    IndexColumn7 = i;
            }
            if (IndexCurrentRow > -1)
            {
                timeIn = ConvertDecimalTimeToStringHHMMArray(ConvertStringTimeHHMMToDecimal(dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn6].Value.ToString()));
                timeOut = ConvertDecimalTimeToStringHHMMArray(ConvertStringTimeHHMMToDecimal(dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn7].Value.ToString()));
            }

            timeInDecimal = ConvertStringTimeHHMMToDecimalArray(ConvertStringsTimeToStringHHMM(timeIn[0], timeIn[1]));
            timeOutDecimal = ConvertStringTimeHHMMToDecimalArray(ConvertStringsTimeToStringHHMM(timeOut[0], timeOut[1]));

            using (var connection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
            {
                connection.Open();
                if (group.Length > 0)
                {
                    using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PersonGroupDesciption' (GroupPerson, GroupPersonDescription) " +
                                            "VALUES (@GroupPerson, @GroupPersonDescription )", connection))
                    {
                        command.Parameters.Add("@GroupPerson", DbType.String).Value = group;
                        command.Parameters.Add("@GroupPersonDescription", DbType.String).Value = textBoxGroupDescription.Text.Trim();
                        try { command.ExecuteNonQuery(); } catch { }
                    }
                }

                if (group.Length > 0 && textBoxNav.Text.Trim().Length > 0)
                {

                    using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PersonGroup' (FIO, NAV, GroupPerson, HourControlling, MinuteControlling, Controlling, HourControllingOut, MinuteControllingOut, ControllingOut, ControllingHHMM, ControllingOUTHHMM) " +
                                                "VALUES (@FIO, @NAV, @GroupPerson, @HourControlling, @MinuteControlling, @Controlling, @HourControllingOut, @MinuteControllingOut, @ControllingOut, @ControllingHHMM, @ControllingOUTHHMM)", connection))
                    {
                        command.Parameters.Add("@FIO", DbType.String).Value = fio;
                        command.Parameters.Add("@NAV", DbType.String).Value = nav;
                        command.Parameters.Add("@GroupPerson", DbType.String).Value = group;
                        command.Parameters.Add("@HourControlling", DbType.String).Value = timeIn[0];
                        command.Parameters.Add("@MinuteControlling", DbType.String).Value = timeIn[1];
                        command.Parameters.Add("@Controlling", DbType.Decimal).Value = timeInDecimal[2];

                        command.Parameters.Add("@HourControllingOut", DbType.String).Value = timeOut[0];
                        command.Parameters.Add("@MinuteControllingOut", DbType.String).Value = timeOut[1];
                        command.Parameters.Add("@ControllingOut", DbType.Decimal).Value = timeOutDecimal[2];

                        command.Parameters.Add("@ControllingHHMM", DbType.String).Value = timeIn[2];
                        command.Parameters.Add("@ControllingOUTHHMM", DbType.String).Value = timeOut[2];
                        try { command.ExecuteNonQuery(); } catch { }
                    }
                    StatusLabel2.Text = "\"" + ShortFIO(textBoxFIO.Text) + "\"" + " добавлен в группу \"" + group + "\"";
                    _toolStripStatusLabelBackColor(StatusLabel2, SystemColors.Control);
                }
                else if (group.Length > 0 && textBoxNav.Text.Trim().Length == 0)
                    try
                    {
                        StatusLabel2.Text = "Отсутствует NAV-код у:" + ShortFIO(textBoxFIO.Text);
                        _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                    } catch { }
                else if (group.Length == 0 && textBoxNav.Text.Trim().Length > 0)
                    try
                    {
                        StatusLabel2.Text = "Не указана группа, в которую нужно добавить!";
                        _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                    } catch { }
            }
            SeekAndShowMembersOfGroup(group);

            PersonOrGroupItem.Text = "Работа с одной персоной";

            group = null;
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

        private async void GetDataItem_Click(object sender, EventArgs e) //GetData()
        { await Task.Run(() => GetData()); }

        private async void GetData()
        {
            _changeControlBackColor(groupBoxPeriod, SystemColors.Control);
            _changeControlBackColor(groupBoxTimeStart, SystemColors.Control);
            _changeControlBackColor(groupBoxTimeEnd, SystemColors.Control);
            _ChangeMenuItemBackColor(QuickLoadDataItem, SystemColors.Control);

            _MenuItemEnabled(QuickLoadDataItem, false);
            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(SettingsMenuItem, false);
            _MenuItemEnabled(AnualDatesMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            CheckBoxesFiltersAll_Enable(false);
            _controlEnable(dataGridView1, false);
            _controlVisible(pictureBox1, false);

            _ProgressBar1Value0();
            _timer1Enabled(true);

            await Task.Run(() => CheckAliveServer(sServer1, sServer1UserName, sServer1UserPassword));

            if (bServer1Exist)
            {
                //Clear work tables
                DeleteAllDataInTableQuery(databasePerson, "PersonTemp");
                DeleteAllDataInTableQuery(databasePerson, "PersonRegistered");
                DeleteAllDataInTableQuery(databasePerson, "PersonRegisteredFull");
                GetNamePoints();  //Get names of the points

                if ((nameOfLastTableFromDB == "PersonGroupDesciption" || nameOfLastTableFromDB == "PersonGroup") && _textBoxReturnText(textBoxGroup).Length > 0)
                {
                    _toolStripStatusLabelSetText(StatusLabel2, "Получаю данные по группе " + _textBoxReturnText(textBoxGroup));
                    stimerPrev = "Получаю данные по группе " + _textBoxReturnText(textBoxGroup);
                }
                else
                {
                    _toolStripStatusLabelSetText(StatusLabel2, "Получаю данные по \"" + ShortFIO(_textBoxReturnText(textBoxFIO)) + "\" ");
                    stimerPrev = "Получаю данные по \"" + ShortFIO(_textBoxReturnText(textBoxFIO)) + "\"";
                    nameOfLastTableFromDB = "PersonRegistered";
                }

                _timer1Enabled(true);
                GetRegistrations();

                CheckBoxesFiltersAll_CheckedState(false);

                _controlVisible(dataGridView1, true);
                _controlEnable(dataGridView1, true);
                _controlEnable(checkBoxReEnter, true);
                _controlEnable(checkBoxTimeViolations, false);
                _controlEnable(checkBoxWeekend, false);
                _controlEnable(checkBoxCelebrate, false);

                _MenuItemEnabled(VisualModeItem, true);
                _MenuItemVisible(VisualModeItem, true);

                _MenuItemEnabled(VisualSelectColorMenuItem, true);
                _MenuItemVisible(VisualSelectColorMenuItem, true);

                _MenuItemEnabled(TableModeItem, true);
                _MenuItemVisible(TableModeItem, false);

                _MenuItemEnabled(TableExportToExcelItem, true);
                _MenuItemVisible(TableExportToExcelItem, true);

                _MenuItemEnabled(TableExportToExcelItem, true);
                _MenuItemEnabled(SettingsMenuItem, true);
                panelViewResize(numberPeopleInLoading);
            }
            else
            {
                GetInfoSetup();
                _MenuItemEnabled(SettingsMenuItem, true);
            }
            _changeControlBackColor(groupBoxRemoveDays, Color.PaleGreen);
        }

        private void GetRegistrations()
        {
            _controlVisible(dataGridView1, false);

            Person person = new Person();
            string selectedGroup = _textBoxReturnText(textBoxGroup);
            decimal dControlHourIn = _numUpDownReturn(numUpDownHourStart);
            decimal dControlMinuteIn = _numUpDownReturn(numUpDownMinuteStart);
            decimal dControlHourOut = _numUpDownReturn(numUpDownHourEnd);
            decimal dControlMinuteOut = _numUpDownReturn(numUpDownMinuteEnd);

            dtPersonRegisteredFull.Clear();

            if ((nameOfLastTableFromDB == "PersonGroupDesciption" || nameOfLastTableFromDB == "PersonGroup") && selectedGroup.Length > 0)
            {
                LoadGroupMembersFromDbToDataTable(dtPersonGroup, selectedGroup); //result will be in dtPersonGroup

                foreach (DataRow row in dtPersonGroup.Rows)
                {
                    if (row[1].ToString().Length > 0 && row[3].ToString() == selectedGroup)
                    {
                        person = new Person();

                        _textBoxSetText(textBoxFIO, row[1].ToString());   //иммитируем выбор данных
                        _textBoxSetText(textBoxNav, row[2].ToString());   //Select person                  

                        dControlHourIn = TryParseStringToDecimal(row[4].ToString());
                        dControlMinuteIn = TryParseStringToDecimal(row[5].ToString());
                        dControlHourOut = TryParseStringToDecimal(row[7].ToString());
                        dControlMinuteOut = TryParseStringToDecimal(row[6].ToString());

                        person.FIO = row[1].ToString();
                        person.NAV = row[2].ToString();
                        person.GroupPerson = selectedGroup;
                        person.ControlInHour = row[4].ToString();
                        person.ControlInHourDecimal = dControlHourIn;
                        person.ControlInMinute = row[5].ToString();
                        person.ControlInMinuteDecimal = dControlMinuteIn;
                        person.ControlInDecimal = ConvertDecimalSeparatedTimeToDecimal(dControlHourIn, dControlMinuteIn);
                        person.ControlInHHMM = ConvertStringsTimeToStringHHMM(row[4].ToString(), row[5].ToString());

                        person.ControlOutHour = row[7].ToString();
                        person.ControlOutHourDecimal = dControlHourOut;
                        person.ControlOutMinute = row[8].ToString();
                        person.ControlOutMinuteDecimal = dControlMinuteOut;
                        person.ControlOutDecimal = ConvertDecimalSeparatedTimeToDecimal(dControlHourOut, dControlMinuteOut);
                        person.ControlOutHHMM = ConvertStringsTimeToStringHHMM(row[7].ToString(), row[8].ToString());

                        GetPersonRegistrationFromServer(dtPersonRegisteredFull, person);     //Search Registration at checkpoints of the selected person
                    }
                }

                nameOfLastTableFromDB = "PersonGroup";
                _timer1Enabled(false);
                _toolStripStatusLabelSetText(StatusLabel2, "Данные по группе \"" + selectedGroup + "\" получены");
            }
            else
            {
                person = new Person();
                person.NAV = _textBoxReturnText(textBoxNav);
                person.FIO = _textBoxReturnText(textBoxFIO);
                person.GroupPerson = selectedGroup;

                person.ControlInHour = dControlHourIn.ToString();
                person.ControlInHourDecimal = dControlHourIn;
                person.ControlInMinute = dControlMinuteIn.ToString();
                person.ControlInMinuteDecimal = dControlMinuteIn;
                person.ControlInDecimal = ConvertDecimalSeparatedTimeToDecimal(dControlHourIn, dControlMinuteIn);
                person.ControlInHHMM = ConvertStringsTimeToStringHHMM(dControlHourIn.ToString(), dControlMinuteIn.ToString());

                person.ControlOutHour = dControlHourOut.ToString();
                person.ControlOutHourDecimal = dControlHourOut;
                person.ControlOutMinute = dControlMinuteOut.ToString();
                person.ControlOutMinuteDecimal = dControlMinuteOut;
                person.ControlOutDecimal = ConvertDecimalSeparatedTimeToDecimal(dControlHourOut, dControlMinuteOut);
                person.ControlOutHHMM = ConvertStringsTimeToStringHHMM(dControlHourOut.ToString(), dControlMinuteOut.ToString());

                GetPersonRegistrationFromServer(dtPersonRegisteredFull, person);

                nameOfLastTableFromDB = "PersonRegistered";
                _timer1Enabled(false);
                _toolStripStatusLabelSetText(StatusLabel2, "Данные с СКД по \"" + ShortFIO(_textBoxReturnText(textBoxFIO)) + "\" получены!");
            }

            person = null;

            dtPersonTemp = dtPersonRegisteredFull.Copy();
            //store all collumns
            dtPersonTempAllCollumns = dtPersonRegisteredFull.Copy();

            string[] nameHidenCollumnsArray =
            {
                @"№ п/п",//0
                @"Время прихода,часы",//4
                @"Время прихода,минут", //5
                @"Время прихода",//6
                @"Время ухода,часы",//7
                @"Время ухода,минут",//8
                @"Время ухода",//9
                @"Время регистрации,часы",//13
                @"Время регистрации,минут",//14
                @"Время регистрации", //15
                @"Реальное время ухода,часы",//16
                @"Реальное время ухода,минут",//17
                @"Реальное время ухода", //18
                //@"Время ухода ЧЧ:ММ",       //23
                @"Реальное отработанное время", //26
                @"Реальное отработанное время ЧЧ:ММ", //27
                @"Опоздание",                    //28
                @"Ранний уход",                 //29
                @"Отпуск (отгул)",              //30
                                  @"Коммандировка",                 //31
                                  @"День недели",                    //32
                                  @"Больничный",                    //33
                                  @"Согласованное отсутствие",      //34
                                  @"Код",                           //35
                                  @"Вышестоящая группа",            //36
                                  @"Описание группы"                //37
            };


            //show selected data     
            //distinct Records                
            var namesDistinctCollumnsArray = arrayAllCollumnsDataTablePeople.Except(nameHidenCollumnsArray).ToArray(); //take distinct data
            dtPersonTemp = GetDistinctRecords(dtPersonTempAllCollumns, namesDistinctCollumnsArray);
            ShowDatatableOnDatagridview(dtPersonTemp, nameHidenCollumnsArray);


            _ProgressBar1Value100();
            stimerPrev = "";
            _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);

            _controlVisible(dataGridView1, true);
            _MenuItemDefaultColorSet(QuickLoadDataItem);
            _ChangeMenuItemBackColor(TableExportToExcelItem, Color.PaleGreen);
            _MenuItemEnabled(QuickLoadDataItem, true);
            _MenuItemEnabled(FunctionMenuItem, true);
            _MenuItemEnabled(SettingsMenuItem, true);
            _MenuItemEnabled(AnualDatesMenuItem, true);
            _MenuItemEnabled(GroupsMenuItem, true);
        }

        private void GetPersonRegistrationFromServer(DataTable dt, Person person)
        {
            DataRow rowPerson;

            decimal hourControlStart = person.ControlInHourDecimal;
            decimal minuteControlStart = person.ControlInMinuteDecimal;
            decimal controlStart = ConvertDecimalSeparatedTimeToDecimal(hourControlStart, minuteControlStart);
            decimal hourControlEnd = person.ControlOutHourDecimal;
            decimal minuteControlEnd = person.ControlOutMinuteDecimal;
            decimal controlEnd = ConvertDecimalSeparatedTimeToDecimal(hourControlEnd, minuteControlEnd);

            string stringIdCardIntellect = "";
            string personNAVTemp = "";
            string[] stringSelectedFIO = { "", "", "" };
            try { stringSelectedFIO[0] = Regex.Split(person.FIO, "[ ]")[0]; } catch { }
            try { stringSelectedFIO[1] = Regex.Split(person.FIO, "[ ]")[1]; } catch { }
            try { stringSelectedFIO[2] = Regex.Split(person.FIO, "[ ]")[2]; } catch { }


            stimerPrev = "Получаю данные по \"" + ShortFIO(person.FIO) + "\"";
            _toolStripStatusLabelSetText(StatusLabel2, "Получаю данные по \"" + ShortFIO(person.FIO) + "\"");

            try
            {
                if (person.NAV.Length == 6)
                {
                    string stringConnection = @"Data Source=" + sServer1 + @"\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID=" + sServer1UserName + @";Password=" + sServer1UserPassword + @";Connect Timeout=240";
                    using (var sqlConnection = new SqlConnection(stringConnection))
                    {
                        sqlConnection.Open();
                        using (var cmd = new SqlCommand("Select id, tabnum FROM OBJ_PERSON where tabnum like '%" + person.NAV + "%';", sqlConnection))
                        {
                            using (var reader = cmd.ExecuteReader())
                            {
                                foreach (DbDataRecord record in reader)
                                {
                                    try
                                    {
                                        _ProgressWork1();

                                        if (record?["tabnum"].ToString().Trim() == person.NAV)
                                        {
                                            stringIdCardIntellect = record["id"].ToString().Trim();
                                            person.idCard = Convert.ToInt32(record["id"].ToString().Trim());
                                            break;
                                        }
                                    } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                                }
                            }
                        }
                    }
                }
                else if (person.NAV.Length != 6)
                {
                    foreach (var strRowWithNav in listFIO.ToArray())
                    {
                        _ProgressWork1();
                        if (strRowWithNav.Contains(person.NAV) && person.NAV.Length > 0 && strRowWithNav.Contains(sServer1))
                            try
                            {
                                stringIdCardIntellect = Regex.Split(strRowWithNav, "[|]")[3].Trim();
                                person.idCard = Convert.ToInt32(stringIdCardIntellect);
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
                                try
                                {
                                    stringIdCardIntellect = Regex.Split(strRowWithNav, "[|]")[3].Trim();
                                    person.idCard = Convert.ToInt32(stringIdCardIntellect);
                                } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                                try
                                {
                                    personNAVTemp = Regex.Split(strRowWithNav, "[|]")[4].Trim();
                                } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                                if (person.NAV.Length < 1 && personNAVTemp.Length > 0)
                                {
                                    person.NAV = personNAVTemp;
                                    _ProgressWork1(); break;
                                }
                            }
                        }
                    }
                }
            } catch (Exception expt) { MessageBox.Show(expt.ToString()); }

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
                    stringSqlWhere = " where action like 'ACCESS_IN' AND param1 like '" + stringIdCardIntellect + "' AND date >= '" + _dateTimePickerStart() + "' AND date <= '" + _dateTimePickerEnd() + "' "; //ORDER BY date DESC
                    stringConnection = @"Data Source=" + sServer1 + @"\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID=" + sServer1UserName + ";Password=" + sServer1UserPassword + "; Connect Timeout=120";
                    using (var sqlConnection = new SqlConnection(stringConnection))
                    {
                        sqlConnection.Open();
                        string query = "SELECT param0, param1, objid, CONVERT(varchar, date, 120) AS date, CONVERT(varchar, PROTOCOL.time, 114) AS time FROM protocol " + stringSqlWhere + " ORDER BY date ASC";
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
                                            managingHours = ConvertDecimalSeparatedTimeToDecimal(hourManaging, minuteManaging);

                                            listRegistrations.Add(
                                                person.FIO + "|" + person.NAV + "|" + record["param1"].ToString().Trim() + "|" + stringDataNew + "|" +
                                                hourManaging + "|" + minuteManaging + "|" + managingHours.ToString("#.###") + "|" +
                                                hourControlStart + "|" + minuteControlStart + "|" + controlStart.ToString("#.###") + "|" +
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

                stringConnection = null;
                stringSqlWhere = null;
                _ProgressWork1();
            } catch (Exception Expt)
            { MessageBox.Show(Expt.ToString(), @"Сервер не доступен, или неправильная авторизация", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            iCounterLine = 0;

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

                rowPerson[1] = person.FIO;
                rowPerson[2] = person.NAV;
                rowPerson[3] = person.GroupPerson;
                rowPerson[10] = cellData[2];
                rowPerson[4] = person.ControlInHour;
                rowPerson[5] = person.ControlInMinute;
                rowPerson[6] = controlStart;
                rowPerson[7] = person.ControlOutHour;
                rowPerson[8] = person.ControlOutMinute;
                rowPerson[9] = person.ControlOutDecimal;
                rowPerson[12] = cellData[3];
                rowPerson[13] = cellData[4];
                rowPerson[14] = cellData[5];
                rowPerson[15] = TryParseStringToDecimal(cellData[6]);
                rowPerson[19] = sServer1;
                rowPerson[20] = namePoint;
                rowPerson[21] = nameDirection;
                rowPerson[22] = ConvertStringsTimeToStringHHMM(cellData[7], cellData[8]);
                rowPerson[23] = ConvertStringsTimeToStringHHMM(person.ControlOutHour, person.ControlOutMinute);
                rowPerson[24] = ConvertStringsTimeToStringHHMM(cellData[4], cellData[5]);

                dt.Rows.Add(rowPerson);
            }
            if (iCounterLine > 0)
            { bLoaded = true; }

            listRegistrations.Clear();
            namePoint = null; nameDirection = null;
            hourControlStart = 0; minuteControlStart = 0;
            stringIdCardIntellect = null; personNAVTemp = null; stringSelectedFIO = new string[1]; cellData = new string[1];
        }

        //Get info the selected group from DB and make a few lists with these data
        private void LoadGroupMembersFromDbToDataTable(DataTable dtTarget, string namePointedGroup) //"Select * FROM PersonGroup where GroupPerson like '" + _textBoxReturnText(textBoxGroup) + "';"
        {
            dtTarget?.Dispose();
            dtTarget = dtPeople.Clone();
            DataRow dataRow;
            numberPeopleInLoading = 0;
            using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
            {
                sqlConnection.Open();
                using (var sqlCommand = new SQLiteCommand(
                    "Select * FROM PersonGroup where GroupPerson like '" + namePointedGroup + "';", sqlConnection))
                {
                    using (var sqlReader = sqlCommand.ExecuteReader())
                    {
                        string d1 = "", d2 = "", d3 = "9", d4 = "0", d13 = "18", d14 = "0";

                        foreach (DbDataRecord record in sqlReader)
                        {
                            d1 = ""; d2 = ""; d3 = "9"; d4 = "0"; d13 = "18"; d14 = "0";
                            try { d1 = record["FIO"].ToString().Trim(); } catch { }

                            if (record != null && d1.Length > 1)
                            {
                                dataRow = dtPersonGroup.NewRow();

                                try { d2 = record["NAV"].ToString().Trim(); } catch { }
                                try { d3 = record["HourControlling"].ToString().Trim(); } catch { }
                                try { d4 = record["MinuteControlling"].ToString().Trim(); } catch { }
                                try { d13 = record["HourControllingOut"].ToString().Trim(); } catch { }
                                try { d14 = record["MinuteControllingOut"].ToString().Trim(); } catch { }
                                //HourControllingOut TEXT, MinuteControllingOut TEXT
                                //FIO|NAV|H|M

                                dataRow[1] = d1;
                                dataRow[2] = d2;
                                dataRow[3] = namePointedGroup;
                                dataRow[4] = d3;
                                dataRow[5] = d4;
                                dataRow[6] = ConvertStringsTimeToDecimal(d3, d4);
                                dataRow[7] = d13;
                                dataRow[8] = d14;
                                dataRow[9] = ConvertStringsTimeToDecimal(d13, d14);
                                dataRow[11] = namePointedGroup;
                                dataRow[22] = ConvertStringsTimeToStringHHMM(d3, d4);
                                dataRow[23] = ConvertStringsTimeToStringHHMM(d13, d14);

                                dtPersonGroup.Rows.Add(dataRow);
                                numberPeopleInLoading++;
                            }
                        }
                        d1 = null; d2 = null; d3 = null; d4 = null; d13 = null; d14 = null;
                    }
                }
            }
        }

        private void DeletePersonFromGroupItem_Click(object sender, EventArgs e) //DeletePersonFromGroup()
        { DeletePersonFromGroup(); }

        private void DeletePersonFromGroup()
        {
            _controlVisible(dataGridView1, false);
            numberPeopleInLoading = 0;

            int indexCurrentRow = _dataGridView1CurrentRowIndex();
            if (indexCurrentRow > -1)
            {
                string nameGroup = DefinyGroupNameByIndexRowDatagridview(indexCurrentRow);
                string navCode = FindPersonNAVInDatagridview(indexCurrentRow);

                DeleteDataTableQueryNAV(databasePerson, "PersonGroup", "NAV", navCode, "GroupPerson", nameGroup);
                SeekAndShowMembersOfGroup(nameGroup);
            }

            nameOfLastTableFromDB = "PersonGroup";
            _controlVisible(dataGridView1, true);
        }

        private string FindPersonNAVInDatagridview(int indexRow)
        {
            string foundNavCode = "";
            try
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
                    if (IndexColumn1 > -1 || IndexColumn2 > -1)
                    {
                        foundNavCode = dataGridView1.Rows[indexRow].Cells[IndexColumn1].Value.ToString();
                    }
                }
            } catch { }
            return foundNavCode;
        }

         private void infoItem_Click(object sender, EventArgs e)
        { ShowDataTableQuery(databasePerson, "TechnicalInfo", "SELECT PCName AS 'Версия Windows',POName AS 'Путь к ПО',POVersion AS 'Версия ПО',LastDateStarted AS 'Дата использования' ", "ORDER BY LastDateStarted DESC"); }

        private void EnterEditAnualItem_Click(object sender, EventArgs e) //Select - EnterEditAnual() or ExitEditAnual()
        {
            if (EnterEditAnualItem.Text.Contains(@"Войти в режим редактирования праздников"))
            { EnterEditAnual(); }
            else if (EnterEditAnualItem.Text.Contains(@"Выйти из режим редактирования"))
            { ExitEditAnual(); }
        }

        private void EnterEditAnual()
        {
            dataGridView1.Visible = false;
            panelView.Visible = false;
            FunctionMenuItem.Enabled = false;
            GroupsMenuItem.Enabled = false;
            _MenuItemEnabled(VisualModeItem, false);
            SettingsMenuItem.Enabled = false;
            QuickLoadDataItem.Enabled = false;
            CheckBoxesFiltersAll_Enable(false);
            comboBoxFio.Enabled = false;

            _timer1Enabled(true);
            StatusLabel2.Text = @"Режим работы с праздниками и выходными";

            StatusLabel2.ForeColor = Color.Crimson;
            AddAnualDateItem.Enabled = true;
            DeleteAnualDateItem.Enabled = true;
            EnterEditAnualItem.Text = @"Выйти из режим редактирования";
        }

        private void ExitEditAnual()
        {
            panelView.Visible = true;
            dataGridView1.Visible = true;

            FunctionMenuItem.Enabled = true;
            GroupsMenuItem.Enabled = true;

            if (bLoaded && _dataGridView1RowsCount() > 0)
            {
                _controlEnable(checkBoxReEnter, true);
                _MenuItemEnabled(VisualModeItem, true);
            }

            SettingsMenuItem.Enabled = true;
            QuickLoadDataItem.Enabled = true;

            comboBoxFio.Enabled = true;

            AddAnualDateItem.Enabled = false;
            DeleteAnualDateItem.Enabled = false;
            EnterEditAnualItem.Text = @"Войти в режим редактирования праздников";
            _timer1Enabled(false);
            StatusLabel2.ForeColor = Color.Black;
            StatusLabel2.Text = "Начните работу с кнопки - \"Получить ФИО\"";
        }

        private void AddAnualDateItem_Click(object sender, EventArgs e) //AddAnualDate()
        { AddAnualDate(); }

        private void AddAnualDate()
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

        private void DeleteAnualDateItem_Click(object sender, EventArgs e) //DeleteAnualDay()
        { DeleteAnualDay(); }

        private void DeleteAnualDay()
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

        private int _dataGridView1CurrentCollumnIndex() //add string into  from other threads
        {
            int iDgv = -1;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { try { iDgv = dataGridView1.CurrentCell.ColumnIndex; } catch { iDgv = -1; } }));
            else
                try { iDgv = dataGridView1.CurrentCell.ColumnIndex; } catch { iDgv = -1; }
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
                Invoke(new MethodInvoker(delegate { tBox = txtBox.Text.ToString().Trim(); }));
            else
                tBox = txtBox.Text.ToString().Trim();
            return tBox;
        }

        private void _textBoxSetText(TextBox txtBox, string s) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { txtBox.Text = s.Trim(); }));
            else
                txtBox.Text = s.Trim();
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

        private void _comboBoxAdd(ComboBox comboBx, string s) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { comboBx.Items.Add(s); }));
            else
                comboBx.Items.Add(s);
        }

        private void _comboBoxClr(ComboBox comboBx) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { comboBx.Items.Clear(); }));
            else
                comboBx.Items.Clear();
        }

        private void _comboBoxSelectIndex(ComboBox comboBx, int i) //add string into comboBoxTargedPC from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { comboBx.SelectedIndex = i; }));
            else
                comboBx.SelectedIndex = i;
        }

        private string _comboBoxReturnSelected(ComboBox comboBox) //add string into comboBoxTargedPC from other threads
        {
            string result = "";
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate
                {
                    result = comboBox.SelectedIndex == -1
                     ? comboBox.Text.Trim()
                         : comboBox.SelectedItem.ToString();
                }));
            }
            else
            {
                result = comboBox.SelectedIndex == -1
                               ? comboBox.Text.Trim()
                                   : comboBox.SelectedItem.ToString();
            }
            return result;
        }


        private void _numUpDownSet(NumericUpDown numericUpDown, decimal i) //add string into comboBoxTargedPC from other threads
        {
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate { try { numericUpDown.Value = i; } catch { numUpDownHourStart.Value = 9; } }));
            }
            else
            {
                try { numericUpDown.Value = i; } catch { numericUpDown.Value = 9; }
            }
        }

        private decimal _numUpDownReturn(NumericUpDown numericUpDown)
        {
            decimal iCombo = 0;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { iCombo = numericUpDown.Value; }));
            else
                iCombo = numericUpDown.Value;
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

        private void _MenuItemDefaultColorSet(ToolStripMenuItem tMenuItem) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { tMenuItem.BackColor = SystemColors.Control; }));
            else
                tMenuItem.BackColor = SystemColors.Control;
        }


        private void _CheckboxCheckedSet(CheckBox checkBox, bool checkboxChecked) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { checkBox.Checked = checkboxChecked; }));
            else
                checkBox.Checked = checkboxChecked;
        }

        private bool _CheckboxCheckedStateReturn(CheckBox checkBox) //add string into  from other threads
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


        private void _menuItemTextSet(ToolStripMenuItem menuItem, string newTextControl) //access from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    menuItem.Text = newTextControl;
                }));
            else
            {
                menuItem.Text = newTextControl;
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

        private void _controlDispose(Control control) //access from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    control?.Dispose();
                }));
            else
            {
                control?.Dispose();
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

        //---- End of Block. Access to Controls from other threads ----//



        //---- Start Block with convertors of data ----//

        private double TryParseStringToDouble(string str)  //string -> decimal. if error it will return 0
        {
            double result = 0;
            try { result = double.Parse(str); } catch { }
            return result;
        }

        private decimal TryParseStringToDecimal(string str)  //string -> decimal. if error it will return 0
        {
            decimal result = 0;
            try { result = decimal.Parse(str); } catch { result = 0; }
            return result;
        }

        private int TryParseStringToInt(string str)  //string -> decimal. if error it will return 0
        {
            int result = 0;
            try { result = int.Parse(str); } catch { }
            return result;
        }

        private string[] ConvertDecimalTimeToStringHHMMArray(decimal decimalTime)
        {
            string[] result = new string[3];
            int hour = (int)(decimalTime);
            int minute = Convert.ToInt32(60 * (decimalTime - hour));

            result[0] = String.Format("{0:d2}", hour);
            result[1] = String.Format("{0:d2}", minute);
            result[2] = String.Format("{0:d2}:{1:d2}", hour, minute);
            return result;
        }

        private string ConvertDecimalTimeToStringHHMM(decimal decimalTime)
        {
            string result;
            int hour = (int)(decimalTime);
            int minute = Convert.ToInt32(60 * (decimalTime - hour));
            result = String.Format("{0:d2}:{1:d2}", hour, minute);
            return result;
        }

        private decimal ConvertDecimalSeparatedTimeToDecimal(decimal decimalHour, decimal decimalMinute)
        {
            decimal result = decimalHour + TryParseStringToDecimal(TimeSpan.FromMinutes((double)decimalMinute).TotalHours.ToString());
            return result;
        }

        private string ConvertStringsTimeToStringHHMM(string hour, string minute)
        {
            int h = 9;
            int m = 0;
            try { h = Convert.ToInt32(hour); } catch { }
            try { m = Convert.ToInt32(minute); } catch { }
            string result = String.Format("{0:d2}:{1:d2}", h, m);
            return result;
        }

        private decimal ConvertStringsTimeToDecimal(string hour, string minute)
        {
            decimal result = TryParseStringToDecimal(hour) + TryParseStringToDecimal(TimeSpan.FromMinutes(TryParseStringToDouble(minute)).TotalHours.ToString());
            return result;
        }

        private decimal ConvertStringTimeHHMMToDecimal(string timeInHHMM) //time HH:MM converted to decimal value
        {
            decimal timeConverted = 9;
            if (timeInHHMM.Contains(':'))
            {
                string[] time = timeInHHMM.Split(':');
                timeConverted = TryParseStringToDecimal(time[0]) + TryParseStringToDecimal(TimeSpan.FromMinutes(TryParseStringToDouble(time[1])).TotalHours.ToString());
            }
            else { timeConverted = TryParseStringToDecimal(timeInHHMM); }
            return timeConverted;
        }

        private decimal[] ConvertStringTimeHHMMToDecimalArray(string timeInHHMM) //time HH:MM converted to decimal value
        {
            decimal[] result = new decimal[4];
            string hour = "9";
            string minute = "0";

            if (timeInHHMM.Contains(':'))
            {
                string[] time = timeInHHMM.Split(':');
                hour = time[0];
                minute = time[1];
            }
            else
            {
                hour = timeInHHMM;
            }

            result[0] = TryParseStringToDecimal(hour);                              // hour in decimal          22
            result[1] = TryParseStringToDecimal(minute);                            // Minute in decimal        15
            result[2] = ConvertDecimalSeparatedTimeToDecimal(result[0], result[1]); // hours in decimal         22.25
            result[3] = 60 * result[0] + result[1];                                    // minutes in decimal       1335

            return result;
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

        //---- End Block with convertors of data ----//




        public void CheckBoxesFiltersAll_CheckedState(bool state)
        {
            _CheckboxCheckedSet(checkBoxTimeViolations, state);
            _CheckboxCheckedSet(checkBoxReEnter, state);
            _CheckboxCheckedSet(checkBoxCelebrate, state);
            _CheckboxCheckedSet(checkBoxWeekend, state);
        }

        public void CheckBoxesFiltersAll_Enable(bool state)
        {
            _controlEnable(checkBoxTimeViolations, state);
            _controlEnable(checkBoxReEnter, state);
            _controlEnable(checkBoxCelebrate, state);
            _controlEnable(checkBoxWeekend, state);
        }

        private async void checkBox_CheckStateChanged(object sender, EventArgs e)
        {
            await Task.Run(() => checkBoxCheckStateChanged());
        }

        private void checkBoxCheckStateChanged()
        {
            CheckBoxesFiltersAll_Enable(false);
            _controlVisible(dataGridView1, false);
            _controlVisible(pictureBox1, false);

            string nameGroup = _textBoxReturnText(textBoxGroup);

            DataTable dtTempIntermediate = dtPersonRegisteredFull.Clone();
            dtPersonTempAllCollumns = dtPersonRegisteredFull.Clone();
            Person personCheck = new Person();
            personCheck.FIO = _textBoxReturnText(textBoxFIO);
            personCheck.NAV = _textBoxReturnText(textBoxNav);
            personCheck.GroupPerson = nameGroup;
            personCheck.Department = nameGroup;

            personCheck.ControlInHour = numUpHourStart.ToString();
            personCheck.ControlInHourDecimal = numUpHourStart;
            personCheck.ControlInMinute = numUpMinuteStart.ToString();
            personCheck.ControlInMinuteDecimal = numUpMinuteStart;
            personCheck.ControlInDecimal = ConvertDecimalSeparatedTimeToDecimal(numUpHourStart, numUpMinuteStart);

            personCheck.ControlOutHour = numUpHourEnd.ToString();
            personCheck.ControlOutHourDecimal = numUpHourEnd;
            personCheck.ControlOutMinute = numUpMinuteEnd.ToString();
            personCheck.ControlOutMinuteDecimal = numUpMinuteEnd;
            personCheck.ControlOutDecimal = ConvertDecimalSeparatedTimeToDecimal(numUpHourEnd, numUpMinuteEnd);

            dtPersonTemp?.Clear();

            if ((nameOfLastTableFromDB == "PersonGroupDesciption" || nameOfLastTableFromDB == "PersonGroup") && nameGroup.Length > 0)
            {
                LoadGroupMembersFromDbToDataTable(dtPersonGroup, nameGroup); //result will be in dtPersonGroup  //"Select * FROM PersonGroup where GroupPerson like '" + _textBoxReturnText(textBoxGroup) + "';"


                if (_CheckboxCheckedStateReturn(checkBoxReEnter))
                {
                    foreach (DataRow row in dtPersonGroup.Rows)
                    {
                        if (row[1].ToString().Length > 0 && row[3].ToString() == nameGroup)
                        {
                            personCheck = new Person();

                            personCheck.FIO = row[1].ToString();
                            personCheck.NAV = row[2].ToString();
                            personCheck.GroupPerson = nameGroup;
                            personCheck.Department = nameGroup;

                            personCheck.ControlInHour = row[4].ToString();
                            personCheck.ControlInHourDecimal = TryParseStringToDecimal(row[4].ToString());
                            personCheck.ControlInMinute = row[5].ToString();
                            personCheck.ControlInMinuteDecimal = TryParseStringToDecimal(row[5].ToString());
                            personCheck.ControlInDecimal = TryParseStringToDecimal(row[6].ToString());

                            personCheck.ControlOutHour = row[7].ToString();
                            personCheck.ControlOutHourDecimal = TryParseStringToDecimal(row[7].ToString());
                            personCheck.ControlOutMinute = row[8].ToString();
                            personCheck.ControlOutMinuteDecimal = TryParseStringToDecimal(row[8].ToString());
                            personCheck.ControlOutDecimal = TryParseStringToDecimal(row[9].ToString());

                            FilterDataByNav(personCheck, dtPersonRegisteredFull, dtTempIntermediate);
                        }
                    }
                }
                else
                {
                    dtTempIntermediate = dtPersonRegisteredFull.Select("[Группа] = '" + nameGroup + "'").Distinct().CopyToDataTable();
                }
            }
            else
            {
                if (!_CheckboxCheckedStateReturn(checkBoxReEnter))
                {
                    dtTempIntermediate = dtPersonRegisteredFull.Copy();
                }
                else
                {
                    FilterDataByNav(personCheck, dtPersonRegisteredFull, dtTempIntermediate);
                }
                nameOfLastTableFromDB = "PersonRegistered";
            }

            string[] arrayHiddenCollumns =
            {
                @"№ п/п",//0
                @"Время прихода,часы",//4
                @"Время прихода,минут", //5
                @"Время прихода",//6
                @"Время ухода,часы",//7
                @"Время ухода,минут",//8
                @"Время ухода",//9
                @"№ пропуска", //10
                @"Время регистрации,часы",//13
                @"Время регистрации,минут",//14
                @"Время регистрации", //15
                @"Реальное время ухода,часы",//16
                @"Реальное время ухода,минут",//17
                @"Реальное время ухода", //18
                @"Сервер СКД", //19
                @"Имя точки прохода", //20
                @"Направление прохода", //21
                @"Реальное отработанное время" //26
            };

            //store all collumns
            dtPersonTempAllCollumns = dtTempIntermediate.Copy();
            //show selected data     
            //distinct Records         

            var namesDistinctCollumnsArray = arrayAllCollumnsDataTablePeople.Except(arrayHiddenCollumns).ToArray(); //take distinct data
            dtPersonTemp = GetDistinctRecords(dtTempIntermediate, namesDistinctCollumnsArray);
            ShowDatatableOnDatagridview(dtPersonTemp, arrayHiddenCollumns);

            panelViewResize(numberPeopleInLoading);
            _controlVisible(dataGridView1, true);

            //change enabling of checkboxes
            if (_CheckboxCheckedStateReturn(checkBoxReEnter))// if (checkBoxReEnter.Checked)
            {
                _controlEnable(checkBoxTimeViolations, true);
                _controlEnable(checkBoxWeekend, true);
                _controlEnable(checkBoxCelebrate, true);

                if (_CheckboxCheckedStateReturn(checkBoxTimeViolations))  // if (checkBoxStartWorkInTime.Checked)
                { _ChangeMenuItemBackColor(QuickLoadDataItem, SystemColors.Control); }
            }
            else if (!_CheckboxCheckedStateReturn(checkBoxReEnter))
            {
                _CheckboxCheckedSet(checkBoxTimeViolations, false);
                _CheckboxCheckedSet(checkBoxWeekend, false);
                _CheckboxCheckedSet(checkBoxCelebrate, false);
                _controlEnable(checkBoxTimeViolations, false);
                _controlEnable(checkBoxWeekend, false);
                _controlEnable(checkBoxCelebrate, false);
            }

            dtTempIntermediate = null;
            _controlEnable(checkBoxReEnter, true);
        }

        public static DataTable GetDistinctRecords(DataTable dt, string[] Columns)
        {
            DataTable dtUniqRecords = new DataTable();
            dtUniqRecords = dt.DefaultView.ToTable(true, Columns);
            return dtUniqRecords;
        }

        private void FilterDataByNav(Person personNAV, DataTable dataTableSource, DataTable dataTableForStoring)    //Copy Data from PersonRegistered into PersonTemp by Filter(NAV and anual dates or minimalTime or dayoff)
        {
            DataRow rowDtStoring;
            DataTable dtTemp = dataTableSource.Clone();

            HashSet<string> hsDays = new HashSet<string>();
            DataTable dtAllRegistrationsInSelectedDay = dataTableSource.Clone(); //All registrations in the selected day

            decimal decimalFirstRegistrationInDay;
            string[] stringHourMinuteFirstRegistrationInDay = new string[2];
            decimal decimalLastRegistrationInDay;
            string[] stringHourMinuteLastRegistrationInDay = new string[2];
            decimal workedHours = 0;

            try
            {
                var allWorkedDaysPerson = dataTableSource.Select("[NAV-код] = '" + personNAV.NAV + "'");

                if (_CheckboxCheckedStateReturn(checkBoxReEnter)) //checkBoxReEnter.Checked
                {
                    foreach (DataRow dataRowDate in allWorkedDaysPerson) //make the list of worked days
                    { hsDays.Add(dataRowDate[12].ToString()); }

                    foreach (var workedDay in hsDays.ToArray())
                    {
                        dtAllRegistrationsInSelectedDay = allWorkedDaysPerson.Distinct().CopyToDataTable().Select("[Дата регистрации] = '" + workedDay + "'").CopyToDataTable();

                        //find first registration within the during selected workedDay
                        //var tempDayStoring = dataTableSource.Select("[Время регистрации] = MIN([Время регистрации])"); //select first by pass
                        decimalFirstRegistrationInDay = Convert.ToDecimal(dtAllRegistrationsInSelectedDay.Compute("MIN([Время регистрации])", string.Empty));

                        //find last registration within the during selected workedDay
                        //var tempDayStoring = dataTableSource.Select("[Время регистрации] = MAX([Время регистрации])"); //select last by pass
                        decimalLastRegistrationInDay = Convert.ToDecimal(dtAllRegistrationsInSelectedDay.Compute("MAX([Время регистрации])", string.Empty));

                        //Select only one row with selected NAV for the selected workedDay
                        rowDtStoring = dtAllRegistrationsInSelectedDay.Select("[Дата регистрации] = '" + workedDay + "'").First();
                        //take and convert a real time coming into a string timearray
                        stringHourMinuteFirstRegistrationInDay = ConvertDecimalTimeToStringHHMMArray(decimalFirstRegistrationInDay);
                        //take the first registration time from the timearray and set into the temprow

                        rowDtStoring[13] = stringHourMinuteFirstRegistrationInDay[0];  //("Время регистрации,часы",typeof(string)),//13
                        rowDtStoring[14] = stringHourMinuteFirstRegistrationInDay[1];  //("Время регистрации,минут", typeof(string)),//14
                        rowDtStoring[15] = decimalFirstRegistrationInDay;              //("Время регистрации", typeof(decimal)), //15
                        rowDtStoring[24] = stringHourMinuteFirstRegistrationInDay[2];  //("Реальное время прихода ЧЧ:ММ", typeof(string)),//24

                        //convert a controlling time coming into a string timearray
                        stringHourMinuteFirstRegistrationInDay = ConvertDecimalTimeToStringHHMMArray(TryParseStringToDecimal(rowDtStoring[6].ToString()));
                        //take a controling time from the timearray and set into the temprow
                        // rowDtStoring[22] = stringHourMinuteFirstRegistrationInDay[2];    //("Время прихода ЧЧ:ММ",typeof(string)),//22

                        rowDtStoring[18] = decimalLastRegistrationInDay;                 //("Реальное время ухода", typeof(decimal)), //18
                        stringHourMinuteLastRegistrationInDay = ConvertDecimalTimeToStringHHMMArray(decimalLastRegistrationInDay);
                        rowDtStoring[16] = stringHourMinuteLastRegistrationInDay[0];     //("Реальное время ухода,часы", typeof(string)),//16
                        rowDtStoring[17] = stringHourMinuteLastRegistrationInDay[1];     //("Реальное время ухода,минут", typeof(string)),//17
                        rowDtStoring[25] = stringHourMinuteLastRegistrationInDay[2];     //("Реальное время ухода ЧЧ:ММ", typeof(string)), //25

                        //taking and conversation controling time come out
                        stringHourMinuteLastRegistrationInDay = ConvertDecimalTimeToStringHHMMArray(TryParseStringToDecimal(rowDtStoring[9].ToString()));
                        // rowDtStoring[23] =  stringHourMinuteLastRegistrationInDay[2];    //("Время ухода ЧЧ:ММ", typeof(string)),//23

                        //worked out times
                        workedHours = decimalLastRegistrationInDay - decimalFirstRegistrationInDay;
                        rowDtStoring[26] = workedHours;                                  // ("Реальное отработанное время", typeof(decimal)), //26
                        rowDtStoring[27] = ConvertDecimalTimeToStringHHMMArray(workedHours)[2];  //("Реальное отработанное время ЧЧ:ММ", typeof(string)), //27

                        if (decimalFirstRegistrationInDay > personNAV.ControlInDecimal) // "Опоздание", typeof(bool)),           //28
                        { rowDtStoring[28] = "Да"; }
                        else { rowDtStoring[28] = ""; }

                        if (decimalLastRegistrationInDay < personNAV.ControlOutDecimal)  // "Ранний уход", typeof(bool)),                 //29
                        { rowDtStoring[29] = "Да"; }
                        else { rowDtStoring[29] = ""; }
                        // MessageBox.Show(                            rowDtStoring[15].ToString() + " - " + rowDtStoring[18].ToString() + "\n" +                            rowDtStoring[28].ToString() + "\n" + rowDtStoring[29].ToString());
                        //rowDtStoring[30] = "false";  //("Отпуск (отгул)", typeof(bool)),                 //30
                        //rowDtStoring[31] = "false";  ("Коммандировка", typeof(bool)),                 //31

                        dtTemp.ImportRow(rowDtStoring);
                    }
                }
                else if (!_CheckboxCheckedStateReturn(checkBoxReEnter))
                {
                    foreach (DataRow dr in allWorkedDaysPerson)
                    { dtTemp.ImportRow(dr); }
                }

                if (_CheckboxCheckedStateReturn(checkBoxWeekend))//checkBoxWeekend Checking
                { DeleteAnualDatesFromDataTables(dtTemp, personNAV); }

                if (_CheckboxCheckedStateReturn(checkBoxTimeViolations)) //checkBoxStartWorkInTime Checking
                { QueryDeleteDataFromDataTable(dtTemp, "[Опоздание]='' AND [Ранний уход]=''", personNAV.NAV); }

                foreach (DataRow dr in dtTemp.AsEnumerable())
                { dataTableForStoring.ImportRow(dr); }

                allWorkedDaysPerson = null;
            } catch (Exception expt) { MessageBox.Show(expt.ToString()); }

            stringHourMinuteFirstRegistrationInDay = null; stringHourMinuteLastRegistrationInDay = null; hsDays = null;
            rowDtStoring = null; dtTemp = null; dtAllRegistrationsInSelectedDay = null;
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

        private void DeleteAnualDatesFromDataTables(DataTable dt, Person person) //Exclude Anual Days from the table "PersonTemp" DB
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
                    QueryDeleteDataFromDataTable(dt, "[Дата регистрации]='" + singleDate + "'", person.NAV); // ("Дата регистрации",typeof(string)),//12
                }
            }
            foreach (var myAnualDate in myMonthCalendar.AnnuallyBoldedDates)
            {
                for (var myDate = myMonthCalendar.SelectionStart; myDate <= myMonthCalendar.SelectionEnd; myDate += oneDay)
                {
                    if (myDate == myAnualDate)
                    {
                        singleDate = Regex.Split(myDate.ToString("yyyy-MM-dd"), " ")[0].Trim();
                        QueryDeleteDataFromDataTable(dt, "[Дата регистрации]='" + singleDate + "'", person.NAV); // ("Дата регистрации",typeof(string)),//12
                    }
                }
            }
            dt.AcceptChanges();
        }

        private void QueryDeleteDataFromDataTable(DataTable dt, string queryFull, string NAVcode) //Delete data from the Table of the DB by NAV (both parameters are string)
        {
            DataRow[] rows = new DataRow[1];
            try
            {

                if (queryFull.Length > 0 && NAVcode.Length > 0)
                { rows = dt.Select("(" + queryFull + ") AND [NAV-код]='" + NAVcode + "'"); }
                else if (queryFull.Length > 0)
                { rows = dt.Select(queryFull); }

                foreach (var row in rows)
                { row.Delete(); }
            } catch (Exception expt)
            { MessageBox.Show(expt.ToString()); }
            dt.AcceptChanges();
            rows = null;
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
            StatusLabel2.Text = @"Временные таблицы удалены";
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
            GC.Collect();
            iFIO = 0;
            StatusLabel2.Text = @"База очищена от загруженных данных. Остались только созданные группы";
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
            StatusLabel2.Text = @"Все таблицы очищены";
        }

        private void SelectPersonFromDataGrid(Person personSelected)
        {
            decimal[] timeIn = new decimal[4];
            decimal[] timeOut = new decimal[4];

            try
            {
                int IndexCurrentRow = _dataGridView1CurrentRowIndex();
                if (IndexCurrentRow > -1)
                {
                    int IndexColumn1 = 0;           // индекс 1-й колонки в датагрид
                    int IndexColumn2 = 0;           // индекс 2-й колонки в датагрид
                    int IndexColumn5 = 0;           // индекс 5-й колонки в датагрид
                    int IndexColumn6 = 0;           // индекс 6-й колонки в датагрид
                    int IndexColumn7 = 0;           // индекс 7-й колонки в датагрид

                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        try
                        {
                            if (dataGridView1.Columns[i].HeaderText.ToString() == "Фамилия Имя Отчество")
                            { IndexColumn1 = i; }
                        } catch { }
                        try
                        {
                            if (dataGridView1.Columns[i].HeaderText.ToString() == "NAV-код")
                            { IndexColumn2 = i; }
                        } catch { }
                        try
                        {
                            if (dataGridView1.Columns[i].HeaderText.ToString() == "Группа")
                            { IndexColumn5 = i; }
                        } catch { }
                        try
                        {
                            if (dataGridView1.Columns[i].HeaderText.ToString() == "Время прихода ЧЧ:ММ")
                            { IndexColumn6 = i; }
                        } catch { }
                        try
                        {
                            if (dataGridView1.Columns[i].HeaderText.ToString() == "Время ухода ЧЧ:ММ")
                            { IndexColumn7 = i; }
                        } catch { }
                    }

                    if (nameOfLastTableFromDB == "PersonGroup")
                    {
                        personSelected.FIO = _dataGridView1CellValue(IndexCurrentRow, IndexColumn1);
                        personSelected.NAV = _dataGridView1CellValue(IndexCurrentRow, IndexColumn2); //Take the name of selected group
                        personSelected.GroupPerson = _dataGridView1CellValue(IndexCurrentRow, IndexColumn5); //Take the name of selected group

                        personSelected.ControlInHHMM = _dataGridView1CellValue(IndexCurrentRow, IndexColumn6); //Take the name of selected group
                        timeIn = ConvertStringTimeHHMMToDecimalArray(personSelected.ControlInHHMM);
                        personSelected.ControlInHour = timeIn[0].ToString();
                        personSelected.ControlInHourDecimal = timeIn[0];
                        personSelected.ControlInMinute = timeIn[1].ToString();
                        personSelected.ControlInMinuteDecimal = timeIn[1];
                        personSelected.ControlInDecimal = timeIn[2];

                        personSelected.ControlOutHHMM = _dataGridView1CellValue(IndexCurrentRow, IndexColumn7); //Take the name of selected group
                        timeOut = ConvertStringTimeHHMMToDecimalArray(personSelected.ControlOutHHMM);
                        personSelected.ControlOutHour = timeOut[0].ToString();
                        personSelected.ControlOutHourDecimal = timeOut[0];
                        personSelected.ControlOutMinute = timeOut[1].ToString();
                        personSelected.ControlOutMinuteDecimal = timeOut[1];
                        personSelected.ControlOutDecimal = timeOut[2];

                        _numUpDownSet(numUpDownHourStart, personSelected.ControlInHourDecimal);
                        _numUpDownSet(numUpDownMinuteStart, personSelected.ControlInMinuteDecimal);

                        StatusLabel2.Text = @"Выбрана группа: " + personSelected.GroupPerson + @" | Курсор на: " + personSelected.FIO;

                        groupBoxPeriod.BackColor = Color.PaleGreen;
                    }
                    else if (nameOfLastTableFromDB == "PersonRegistered")
                    {
                        personSelected.FIO = _dataGridView1CellValue(IndexCurrentRow, IndexColumn1);
                        personSelected.NAV = _dataGridView1CellValue(IndexCurrentRow, IndexColumn2); //Take the name of selected group
                        personSelected.GroupPerson = _textBoxReturnText(textBoxGroup);

                        personSelected.ControlInHHMM = _dataGridView1CellValue(IndexCurrentRow, IndexColumn6); //Take the name of selected group
                        timeIn = ConvertStringTimeHHMMToDecimalArray(personSelected.ControlInHHMM);
                        personSelected.ControlInHour = timeIn[0].ToString();
                        personSelected.ControlInHourDecimal = timeIn[0];
                        personSelected.ControlInMinute = timeIn[1].ToString();
                        personSelected.ControlInMinuteDecimal = timeIn[1];
                        personSelected.ControlInDecimal = timeIn[2];

                        personSelected.ControlOutHHMM = _dataGridView1CellValue(IndexCurrentRow, IndexColumn7); //Take the name of selected group
                        timeOut = ConvertStringTimeHHMMToDecimalArray(personSelected.ControlOutHHMM);
                        personSelected.ControlOutHour = timeOut[0].ToString();
                        personSelected.ControlOutHourDecimal = timeOut[0];
                        personSelected.ControlOutMinute = timeOut[1].ToString();
                        personSelected.ControlOutMinuteDecimal = timeOut[1];
                        personSelected.ControlOutDecimal = timeOut[2];

                        _numUpDownSet(numUpDownHourStart, personSelected.ControlInHourDecimal);
                        _numUpDownSet(numUpDownMinuteStart, personSelected.ControlInMinuteDecimal);

                        StatusLabel2.Text = @"Выбран: " + personSelected.FIO + @" |  Всего ФИО: " + iFIO;
                    }
                }
            } catch (Exception expt) { MessageBox.Show(expt.ToString()); }

            if (personSelected.FIO.Length == 0)
            {
                SelectPersonFromControls(personSelected);
            }
        }

        //gathering a person's features from textboxes and other controls
        private void SelectPersonFromControls(Person personSelected)
        {
            try
            {
                personSelected.FIO = _textBoxReturnText(textBoxFIO);
                personSelected.NAV = _textBoxReturnText(textBoxNav);
                personSelected.GroupPerson = _textBoxReturnText(textBoxGroup);

                personSelected.ControlInHourDecimal = _numUpDownReturn(numUpDownHourStart);
                personSelected.ControlInHour = personSelected.ControlInHourDecimal.ToString();
                personSelected.ControlInMinuteDecimal = _numUpDownReturn(numUpDownMinuteStart);
                personSelected.ControlInMinute = personSelected.ControlInMinuteDecimal.ToString();
                personSelected.ControlInHHMM = ConvertStringsTimeToStringHHMM(personSelected.ControlInHour, personSelected.ControlInMinute);

                personSelected.ControlOutHourDecimal = _numUpDownReturn(numUpDownHourEnd);
                personSelected.ControlOutHour = personSelected.ControlOutHourDecimal.ToString();
                personSelected.ControlOutMinuteDecimal = _numUpDownReturn(numUpDownMinuteEnd);
                personSelected.ControlOutMinute = personSelected.ControlOutMinuteDecimal.ToString();
                personSelected.ControlOutHHMM = ConvertStringsTimeToStringHHMM(personSelected.ControlOutHour, personSelected.ControlOutMinute);
            } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
        }

        private void FindWorkDatesInSelected() //
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

        private void VisualItem_Click(object sender, EventArgs e)
        {
            Person personVisual = new Person();
            if (bLoaded)
            {
                SelectPersonFromDataGrid(personVisual);
                dataGridView1.Visible = false;
                FindWorkDatesInSelected();
                CheckBoxesFiltersAll_Enable(false);

                if (_CheckboxCheckedStateReturn(checkBoxReEnter))
                { DrawFullWorkedPeriodRegistration(personVisual); }
                else
                { DrawRegistration(personVisual); }
                _MenuItemVisible(TableModeItem, true);
                _MenuItemVisible(VisualModeItem, false);
                _MenuItemVisible(TableExportToExcelItem, false);
            }
            else
            { MessageBox.Show("Таблица с данными пустая.\nНе загружены данные для визуализации!"); }
        }

        private int numberPeopleInLoading = 1;
        private void DrawRegistration(Person personDraw)  // Visualisation of registration
        {
            try
            {
                //  int iPanelBorder = 2;
                int iMinutesInHour = 60;
                int iShiftStart = 300;
                int iStringHeight = 19;
                int iShiftHeightAll = 36;

                int iLenghtRect = 0; //количество  входов-выходов в рабочие дни для всех отобранных людей для  анализа регистраций входа-выхода

                //constant for a person
                string fio = personDraw.FIO;
                string nav = personDraw.NAV;
                string group = personDraw.GroupPerson;
                string dayRegistration = ""; string directionPass = ""; //string pointName = "";
                int minutesIn = 0;     // время входа в минутах планируемое
                int minutesInFact = 0;     // время выхода в минутах фактическое
                int minutesOut = 0;    // время входа в минутах планируемое
                int minutesOutFact = 0;    // время выхода в минутах фактическое

                //variable for a person
                string dayPrevious = "";      //дата в предыдущей выборке
                string directionPrevious = "";
                int timePrevious = 0;

                //select and distinct dataRow
                var rowsPersonRegistrationsForDraw = dtPersonTempAllCollumns.AsEnumerable();
                if (group.Length > 0)
                { rowsPersonRegistrationsForDraw = dtPersonTempAllCollumns.Select("[Группа] = '" + group + "'").CopyToDataTable().AsEnumerable(); }
                else if (nav.Length == 6)
                { rowsPersonRegistrationsForDraw = dtPersonTempAllCollumns.Select("[NAV-код] = '" + nav + "'").CopyToDataTable().AsEnumerable(); }
                else if (nav.Length != 6 && fio.Length > 1)
                { rowsPersonRegistrationsForDraw = dtPersonTempAllCollumns.Select("[Фамилия Имя Отчество] = '" + fio + "'").CopyToDataTable().AsEnumerable(); }

                //select and count unique NAV-codes - the number of selected people
                HashSet<string> hsNAV = new HashSet<string>(); //unique NAV-codes
                foreach (DataRow row in rowsPersonRegistrationsForDraw)
                { hsNAV.Add(row["NAV-код"].ToString()); }
                string[] arrayNAVs = hsNAV.ToArray(); //unique NAV-codes
                int countNAVs = arrayNAVs.Count(); //the number of selected people
                numberPeopleInLoading = countNAVs;

                //count max number of events in-out all of selected people (the group or a single person)
                //It needs to prevent the error "index of scope"
                foreach (DataRow row in rowsPersonRegistrationsForDraw)
                {
                    for (int k = 0; k < workSelectedDays.Length; k++)
                    {
                        if (workSelectedDays[k].Length == 10 && row["Дата регистрации"].ToString().Contains(workSelectedDays[k]))  //если приход пришел в рабочий день - считаем
                        { iLenghtRect++; }
                    }
                }

                panelView.Height = iShiftHeightAll + iStringHeight * workSelectedDays.Count() * countNAVs;
                // panelView.AutoScroll = false;
                // panelView.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                // panelView.Anchor = AnchorStyles.Bottom;
                // panelView.Anchor = AnchorStyles.Left;
                // panelView.Dock = DockStyle.None;
                panelView.ResumeLayout();

                //   bmp?.Dispose();
                pictureBox1?.Dispose();
                if (panelView.Controls.Count > 1)
                { panelView.Controls.RemoveAt(1); }

                pictureBox1 = new PictureBox
                {
                    //    Location = new Point(0, 0),
                    SizeMode = PictureBoxSizeMode.AutoSize,
                    Size = new Size(
                        iShiftStart + 24 * iMinutesInHour,
                        iShiftHeightAll + iStringHeight * workSelectedDays.Count() * countNAVs
                    // (iShiftStart + 24 * iMinutesInHour + 2) / 2 // 1740 на 870 - 24 часа и 43 строчки
                    // 2 * (iShiftStart + 24 * iMinutesInHour + 2) / 5  //1740 на 696 - 24 часа и 34 строчки
                    ),
                    BorderStyle = BorderStyle.FixedSingle
                };
                //comment next row if it needs to set the PictureBox at the Center of panelview
                bmp = new Bitmap(pictureBox1.Width, pictureBox1.Height);


                // Start of the Block with Drawing //
                //---------------------------------------------------------------//
                var font = new Font("Courier", 10, FontStyle.Regular);
                using (Graphics gr = Graphics.FromImage(bmp))
                {
                    var myBrushWorkHour = new SolidBrush(Color.Gray);
                    var myBrushRealWorkHour = new SolidBrush(clrRealRegistration);
                    var myBrushAxis = new SolidBrush(Color.Black);
                    var pointForN_A = new PointF(0, iStringHeight);
                    var pointForN_B = new PointF(200, iStringHeight);

                    var axis = new Pen(Color.Black);

                    Rectangle[] rectsReal = new Rectangle[iLenghtRect]; //количество пересечений
                    Rectangle[] rectsRealMark = new Rectangle[iLenghtRect];
                    Rectangle[] rects = new Rectangle[workSelectedDays.Length * countNAVs];

                    int irectsTempReal = 0;

                    int numberRectangle_rectsRealMark = 0;
                    int numberRectangle_rects = 0;

                    int pointDrawYfor_rectsReal = 39;
                    int pointDrawYfor_rects = 42;

                    foreach (string singleNav in arrayNAVs)
                    {
                        foreach (string workDay in workSelectedDays)
                        {
                            foreach (DataRow row in rowsPersonRegistrationsForDraw)
                            {
                                nav = row["NAV-код"].ToString();
                                dayRegistration = row["Дата регистрации"].ToString();

                                if (singleNav.Contains(nav) && dayRegistration.Contains(workDay))
                                {
                                    fio = row["Фамилия Имя Отчество"].ToString();
                                    minutesIn = (int)ConvertStringTimeHHMMToDecimalArray(row["Время прихода ЧЧ:ММ"].ToString())[3];
                                    minutesInFact = (int)ConvertStringTimeHHMMToDecimalArray(row["Реальное время прихода ЧЧ:ММ"].ToString())[3];
                                    minutesOut = (int)ConvertStringTimeHHMMToDecimalArray(row["Время ухода ЧЧ:ММ"].ToString())[3];
                                    minutesOutFact = (int)ConvertStringTimeHHMMToDecimalArray(row["Реальное время ухода ЧЧ:ММ"].ToString())[3];
                                    directionPass = row["Направление прохода"].ToString().ToLower();

                                    //pass by a point
                                    rectsRealMark[numberRectangle_rectsRealMark] = new Rectangle(
                                    iShiftStart + minutesInFact,             /* X */
                                    pointDrawYfor_rectsReal,                 /* Y */
                                    2,                                       /* width */
                                    14                                       /* height */
                                    );

                                    //being of the current person in the workplace
                                    if (directionPass.Contains("вход"))
                                    {
                                        timePrevious = minutesInFact;
                                        dayPrevious = dayRegistration;
                                        directionPrevious = directionPass;
                                    }
                                    else if (directionPass.Contains("выход") && directionPrevious.Contains("вход") && dayPrevious.Contains(dayRegistration))
                                    {
                                        if (minutesInFact > timePrevious)
                                        {
                                            rectsReal[irectsTempReal] = new Rectangle(
                                                    iShiftStart + timePrevious,
                                                    pointDrawYfor_rectsReal,
                                                    minutesInFact - timePrevious,
                                                    14);                //height

                                            irectsTempReal++;
                                            timePrevious = minutesInFact;
                                            dayPrevious = dayRegistration;
                                            directionPrevious = directionPass;
                                        }
                                    }

                                    numberRectangle_rectsRealMark++;
                                }
                            }

                            //work shift
                            rects[numberRectangle_rects] = new Rectangle(
                               iShiftStart + minutesIn,                     /* X */
                               pointDrawYfor_rects,                         /* Y */
                               minutesOut - minutesIn,                      /* width */
                               6                                            /* height */
                               );

                            pointDrawYfor_rectsReal += 19;
                            pointDrawYfor_rects += 19;
                            numberRectangle_rects++;
                        }

                        //place the current FIO and days in visualisation
                        foreach (string workDay in workSelectedDays)
                        {
                            pointForN_A.Y += iStringHeight;
                            gr.DrawString(
                                workDay + " (" + ShortFIO(fio) + ")",
                                font,
                                myBrushAxis,
                                pointForN_A); //Paint workdays and people' FIO
                        }
                    }

                    //Fill with rectangles RealWork
                    gr.FillRectangles(myBrushRealWorkHour, rectsReal);

                    // Fill rectangles WorkTime shit
                    gr.FillRectangles(myBrushWorkHour, rects);

                    //Fill All Mark at Passthrow Points
                    gr.FillRectangles(myBrushRealWorkHour, rectsRealMark); //draw the real first come of the person

                    //Draw axes for days 
                    for (int k = 0; k < workSelectedDays.Length * countNAVs; k++)
                    {
                        pointForN_B.Y += iStringHeight;
                        gr.DrawLine(
                            axis,
                            new Point(0, iShiftHeightAll + k * iStringHeight),
                            new Point(pictureBox1.Width, iShiftHeightAll + k * iStringHeight));
                    }

                    //Draw other axes
                    gr.DrawString(
                        "Время, часы:",
                        font,
                        SystemBrushes.WindowText,
                        new Point(iShiftStart - 110, iStringHeight / 4));
                    gr.DrawString("Дата (ФИО)",
                        font,
                        SystemBrushes.WindowText,
                        new Point(10, iStringHeight));
                    gr.DrawLine(
                        axis, new Point(0, 0),
                        new Point(iShiftStart, iShiftHeightAll));
                    gr.DrawLine(
                        axis,
                        new Point(iShiftStart, 0),
                        new Point(iShiftStart, iShiftHeightAll));

                    for (int k = 0; k <= 23; k++)
                    {
                        gr.DrawLine(
                            axis,
                            new Point(iShiftStart + k * iMinutesInHour, iShiftHeightAll),
                            new Point(iShiftStart + k * iMinutesInHour, Convert.ToInt32(pictureBox1.Height)));
                        gr.DrawString(
                            Convert.ToString(k),
                            font,
                            SystemBrushes.WindowText,
                            new Point(320 + k * iMinutesInHour, iStringHeight));
                    }
                    gr.DrawLine(
                        axis,
                        new Point(iShiftStart + 24 * iMinutesInHour, iShiftHeightAll),
                        new Point(iShiftStart + 24 * iMinutesInHour, Convert.ToInt32(pictureBox1.Height)));

                    axis.Dispose();
                    myBrushAxis = null;
                    rectsReal = null;
                    rects = null;
                    myBrushRealWorkHour = null;
                    myBrushWorkHour = null;
                }
                //---------------------------------------------------------------//
                // End of the Block with Drawing //


                pictureBox1.Image = bmp;
                pictureBox1.Refresh();
                panelView.Controls.Add(pictureBox1);
                _RefreshPictureBox(pictureBox1, bmp);
                panelViewResize(numberPeopleInLoading);

                fio = null; nav = null; dayRegistration = null; directionPass = null;
                font.Dispose(); hsNAV = null;

                //---------------------------------------------------------------//
                //указываем выбранный ФИО и устанавливаем на него фокус
                textBoxFIO.Text = fio;
                textBoxNav.Text = nav;
                StatusLabel2.Text = @"Выбран: " + ShortFIO(fio) + @" |  NAV: " + nav;
                if (comboBoxFio.FindString(fio) != -1) comboBoxFio.SelectedIndex = comboBoxFio.FindString(fio); //ищем в комбобокс выбранный ФИО и устанавливаем на него фокус

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
            } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
            sLastSelectedElement = "DrawRegistration";
        }

        private void DrawFullWorkedPeriodRegistration(Person personDraw)  // Draw the whole period registration
        {
            //  int iPanelBorder = 2;
            int iMinutesInHour = 60;
            int iShiftStart = 300;
            int iStringHeight = 19;
            int iShiftHeightAll = 36;

            int iLenghtRect = 0; //количество  входов-выходов в рабочие дни для всех отобранных людей для  анализа регистраций входа-выхода

            //constant for a person
            string fio = personDraw.FIO;
            string nav = personDraw.NAV;
            string group = personDraw.GroupPerson;
            string dayRegistration = ""; string directionPass = ""; //string pointName = "";
            int minutesIn = 0;     // время входа в минутах планируемое
            int minutesInFact = 0;     // время выхода в минутах фактическое
            int minutesOut = 0;    // время входа в минутах планируемое
            int minutesOutFact = 0;    // время выхода в минутах фактическое

            //select and distinct dataRow
            var rowsPersonRegistrationsForDraw = dtPersonTempAllCollumns.AsEnumerable();
            if (group.Length > 0)
            { rowsPersonRegistrationsForDraw = dtPersonTempAllCollumns.Select("[Группа] = '" + group + "'").CopyToDataTable().AsEnumerable(); }
            else if (nav.Length == 6)
            { rowsPersonRegistrationsForDraw = dtPersonTempAllCollumns.Select("[NAV-код] = '" + nav + "'").CopyToDataTable().AsEnumerable(); }
            else if (nav.Length != 6 && fio.Length > 1)
            { rowsPersonRegistrationsForDraw = dtPersonTempAllCollumns.Select("[Фамилия Имя Отчество] = '" + fio + "'").CopyToDataTable().AsEnumerable(); }

            //select and count unique NAV-codes - the number of selected people
            HashSet<string> hsNAV = new HashSet<string>(); //unique NAV-codes
            foreach (DataRow row in rowsPersonRegistrationsForDraw)
            { hsNAV.Add(row["NAV-код"].ToString()); }
            string[] arrayNAVs = hsNAV.ToArray(); //unique NAV-codes
            int countNAVs = arrayNAVs.Count(); //the number of selected people
            numberPeopleInLoading = countNAVs;

            //count max number of events in-out all of selected people (the group or a single person)
            //It needs to prevent the error "index of scope"
            foreach (DataRow row in rowsPersonRegistrationsForDraw)
            {
                for (int k = 0; k < workSelectedDays.Length; k++)
                {
                    if (workSelectedDays[k].Length == 10 && row["Дата регистрации"].ToString().Contains(workSelectedDays[k]))  //если приход пришел в рабочий день - считаем
                    { iLenghtRect++; }
                }
            }

            panelView.Height = iShiftHeightAll + iStringHeight * workSelectedDays.Count() * countNAVs;
            // panelView.AutoScroll = false;
            // panelView.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            // panelView.Anchor = AnchorStyles.Bottom;
            // panelView.Anchor = AnchorStyles.Left;
            // panelView.Dock = DockStyle.None;
            panelView.ResumeLayout();

            //   bmp?.Dispose();
            pictureBox1?.Dispose();
            if (panelView.Controls.Count > 1)
            { panelView.Controls.RemoveAt(1); }

            pictureBox1 = new PictureBox
            {
                //Location = new Point(0, 0),
                SizeMode = PictureBoxSizeMode.AutoSize,
                Size = new Size(
                    iShiftStart + 24 * iMinutesInHour,
                    iShiftHeightAll + iStringHeight * workSelectedDays.Count() * countNAVs
                // (iShiftStart + 24 * iMinutesInHour + 2) / 2 // 1740 на 870 - 24 часа и 43 строчки
                // 2 * (iShiftStart + 24 * iMinutesInHour + 2) / 5  //1740 на 696 - 24 часа и 34 строчки
                ),
                BorderStyle = BorderStyle.FixedSingle
            };
            //comment next row if it needs to set the PictureBox at the Center of panelview
            bmp = new Bitmap(pictureBox1.Width, pictureBox1.Height);


            // Start of the Block with Drawing //
            //---------------------------------------------------------------//
            var font = new Font("Courier", 10, FontStyle.Regular);
            using (Graphics gr = Graphics.FromImage(bmp))
            {
                var myBrushWorkHour = new SolidBrush(Color.Gray);
                var myBrushRealWorkHour = new SolidBrush(clrRealRegistration);
                var myBrushAxis = new SolidBrush(Color.Black);
                var pointForN_A = new PointF(0, iStringHeight);
                var pointForN_B = new PointF(200, iStringHeight);

                var axis = new Pen(Color.Black);

                Rectangle[] rectsRealMark = new Rectangle[iLenghtRect];
                Rectangle[] rects = new Rectangle[workSelectedDays.Length * countNAVs];

                int numberRectangle_rectsRealMark = 0;
                int numberRectangle_rects = 0;

                int pointDrawYfor_rectsReal = 39;
                int pointDrawYfor_rects = 42;

                foreach (string singleNav in arrayNAVs)
                {
                    foreach (string workDay in workSelectedDays)
                    {
                        foreach (DataRow row in rowsPersonRegistrationsForDraw)
                        {
                            nav = row["NAV-код"].ToString();
                            dayRegistration = row["Дата регистрации"].ToString();

                            if (singleNav.Contains(nav) && dayRegistration.Contains(workDay))
                            {
                                fio = row["Фамилия Имя Отчество"].ToString();
                                minutesIn = (int)ConvertStringTimeHHMMToDecimalArray(row["Время прихода ЧЧ:ММ"].ToString())[3];
                                minutesInFact = (int)ConvertStringTimeHHMMToDecimalArray(row["Реальное время прихода ЧЧ:ММ"].ToString())[3];
                                minutesOut = (int)ConvertStringTimeHHMMToDecimalArray(row["Время ухода ЧЧ:ММ"].ToString())[3];
                                minutesOutFact = (int)ConvertStringTimeHHMMToDecimalArray(row["Реальное время ухода ЧЧ:ММ"].ToString())[3];
                                directionPass = row["Направление прохода"].ToString().ToLower();

                                //pass by a point
                                rectsRealMark[numberRectangle_rectsRealMark] = new Rectangle(
                                iShiftStart + minutesInFact,             /* X */
                                pointDrawYfor_rectsReal,                 /* Y */
                                minutesOutFact - minutesInFact,           /* width */
                                14                                       /* height */
                                );

                                numberRectangle_rectsRealMark++;
                            }
                        }

                        //work shift
                        rects[numberRectangle_rects] = new Rectangle(
                           iShiftStart + minutesIn,                     /* X */
                           pointDrawYfor_rects,                         /* Y */
                           minutesOut - minutesIn,                      /* width */
                           6                                            /* height */
                           );

                        pointDrawYfor_rectsReal += 19;
                        pointDrawYfor_rects += 19;
                        numberRectangle_rects++;
                    }

                    //place the current FIO and days in visualisation
                    foreach (string workDay in workSelectedDays)
                    {
                        pointForN_A.Y += iStringHeight;
                        gr.DrawString(
                            workDay + " (" + ShortFIO(fio) + ")",
                            font,
                            myBrushAxis,
                            pointForN_A); //Paint workdays and people' FIO
                    }
                }

                //Fill All Mark at Passthrow Points
                gr.FillRectangles(myBrushRealWorkHour, rectsRealMark); //draw the real first come of the person

                // Fill rectangles WorkTime shit
                gr.FillRectangles(myBrushWorkHour, rects);

                //Draw axes for days 
                for (int k = 0; k < workSelectedDays.Length * countNAVs; k++)
                {
                    pointForN_B.Y += iStringHeight;
                    gr.DrawLine(
                        axis,
                        new Point(0, iShiftHeightAll + k * iStringHeight),
                        new Point(pictureBox1.Width, iShiftHeightAll + k * iStringHeight));
                }

                //Draw other axes
                gr.DrawString(
                    "Время, часы:",
                    font,
                    SystemBrushes.WindowText,
                    new Point(iShiftStart - 110, iStringHeight / 4));
                gr.DrawString("Дата (ФИО)",
                    font,
                    SystemBrushes.WindowText,
                    new Point(10, iStringHeight));
                gr.DrawLine(
                    axis, new Point(0, 0),
                    new Point(iShiftStart, iShiftHeightAll));
                gr.DrawLine(
                    axis,
                    new Point(iShiftStart, 0),
                    new Point(iShiftStart, iShiftHeightAll));

                for (int k = 0; k <= 23; k++)
                {
                    gr.DrawLine(
                        axis,
                        new Point(iShiftStart + k * iMinutesInHour, iShiftHeightAll),
                        new Point(iShiftStart + k * iMinutesInHour, Convert.ToInt32(pictureBox1.Height)));
                    gr.DrawString(
                        Convert.ToString(k),
                        font,
                        SystemBrushes.WindowText,
                        new Point(320 + k * iMinutesInHour, iStringHeight));
                }
                gr.DrawLine(
                    axis,
                    new Point(iShiftStart + 24 * iMinutesInHour, iShiftHeightAll),
                    new Point(iShiftStart + 24 * iMinutesInHour, Convert.ToInt32(pictureBox1.Height)));

                axis.Dispose();
                myBrushAxis = null;
                rects = null;
                myBrushRealWorkHour = null;
                myBrushWorkHour = null;
            }
            //---------------------------------------------------------------//
            // End of the Block with Drawing //


            pictureBox1.Image = bmp;
            pictureBox1.Refresh();
            panelView.Controls.Add(pictureBox1);
            _RefreshPictureBox(pictureBox1, bmp);
            panelViewResize(numberPeopleInLoading);

            fio = null; nav = null; dayRegistration = null; directionPass = null;
            font.Dispose(); hsNAV = null;

            //---------------------------------------------------------------//
            //указываем выбранный ФИО и устанавливаем на него фокус
            textBoxFIO.Text = fio;
            textBoxNav.Text = nav;
            StatusLabel2.Text = @"Выбран: " + ShortFIO(fio) + @" |  NAV: " + nav;
            if (comboBoxFio.FindString(fio) != -1) comboBoxFio.SelectedIndex = comboBoxFio.FindString(fio); //ищем в комбобокс выбранный ФИО и устанавливаем на него фокус

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
            sLastSelectedElement = "DrawFullWorkedPeriodRegistration";
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
            try
            {
                if (panelView != null && panelView.Controls.Count > 1) panelView.Controls.RemoveAt(1);
                bmp?.Dispose();
                pictureBox1?.Dispose();
            } catch { }

            dataGridView1.Visible = true;
            sLastSelectedElement = "dataGridView";
            panelViewResize(numberPeopleInLoading);

            CheckBoxesFiltersAll_Enable(true);
            _MenuItemVisible(TableExportToExcelItem, true);
            _MenuItemVisible(TableModeItem, false);
            _MenuItemVisible(VisualModeItem, true);
            _MenuItemVisible(VisualSelectColorMenuItem, false);
        }

        private void panelView_SizeChanged(object sender, EventArgs e)
        { panelViewResize(numberPeopleInLoading); }

        private void panelViewResize(int numberPeople) //Change PanelView
        {
            int iStringHeight = 19;
            int iShiftHeightAll = 36;
            switch (sLastSelectedElement)
            {
                case "DrawFullWorkedPeriodRegistration":
                    _panelSetHeight(panelView, iShiftHeightAll + iStringHeight * workSelectedDays.Length * numberPeople); //Fixed size of Picture. If need autosize - disable this row
                    break;
                case "DrawRegistration":
                    _panelSetHeight(panelView, iShiftHeightAll + iStringHeight * workSelectedDays.Length * numberPeople); //Fixed size of Picture. If need autosize - disable this row
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

        private void BlueItem_Click(object sender, EventArgs e)
        { ColorizeDraw(Color.MediumAquamarine); }

        private void RedItem_Click(object sender, EventArgs e)
        { ColorizeDraw(Color.DarkOrange); }

        private void GreenItem_Click(object sender, EventArgs e)
        { ColorizeDraw(Color.LimeGreen); }

        private void YellowItem_Click(object sender, EventArgs e)
        { ColorizeDraw(Color.Orange); }

        public void ColorizeDraw(Color color)
        {
            Person personSelected = new Person();
            SelectPersonFromControls(personSelected);

            clrRealRegistration = color;
            switch (sLastSelectedElement)
            {
                case "DrawFullWorkedPeriodRegistration":
                    DrawFullWorkedPeriodRegistration(personSelected);
                    break;
                case "DrawRegistration":
                    DrawRegistration(personSelected);
                    break;
                default:
                    break;
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

        private bool isPerson = true; //Check
        private void PersonOrGroupItem_Click(object sender, EventArgs e) //PersonOrGroup()
        { PersonOrGroup(isPerson); }

        private void PersonOrGroup(bool isPerson)
        {
            if (isPerson)
            {
                _menuItemTextSet(PersonOrGroupItem, "Работа с группой");
                nameOfLastTableFromDB = "PersonGroup";
                isPerson = false;
            }
            else
            {
                _menuItemTextSet(PersonOrGroupItem, "Работа с одной персоной");
                nameOfLastTableFromDB = "PersonRegistered";
                isPerson = true;
            }
        }

        private void SetupItem_Click(object sender, EventArgs e) //GetInfoSetup()
        { GetInfoSetup(); }

        private void GetInfoSetup()
        {
            DialogResult result = MessageBox.Show(
                @"Перед получением информации необходимо в Настройках:" + "\n\n" +
                 "1. Добавить имя СКД-сервера Интеллект (SERVER.DOMAIN.SUBDOMAIN),\nа также - имя и пароль пользователя для доступа к SQL-серверу СКД\n" +
                 "2. Сохранить введенные параметры\n" +
                 "3.1. После этого можно получить список сотрудников, хранимый на СКД-сервере\n" +
                 "3.2. Использовать ранее сохраненные группы пользователей локально\n" +
                 "3.3. Добавить праздничные дни, отгулы, отпуски\n" +
                 "4. Загрузить данные по группам или отдельным сотрудников, регистрировавшимся на входных дверях\n" +
                 "5. Далее можно проводить анализ данных в табличном виде или визуально, экспортировать данные с таблиц в Excel файл.\n\nДата и время локального ПК: " +
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



        //---- Start. Features of programm ---//
        private void MailingItem_Click(object sender, EventArgs e) //MailingItem()
        {
            nameOfLastTableFromDB = "Mailing";
            _MenuItemEnabled(SettingsMenuItem, false);
            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            _MenuItemEnabled(AnualDatesMenuItem, false);
            CheckBoxesFiltersAll_Enable(false);
            _controlVisible(panelView, false);

            btnPropertiesSave.Text = "Сохранить рассылку";

            MailingItem();
        }

        private void MailingItem()
        {
            List<string> listComboParameters = new List<string>();
            listComboParameters.Add("Все");

            List<string> periodComboParameters = new List<string>();
            periodComboParameters.Add("Текущая неделя");
            periodComboParameters.Add("Текущий месяц");
            periodComboParameters.Add("Предыдущая неделя");
            periodComboParameters.Add("Предыдущий месяц");
            
            //get list groups from DB and add to listComboParameters
            using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
            {
                sqlConnection.Open();
                using (var sqlCommand = new SQLiteCommand(
                    "SELECT GroupPerson, GroupPersonDescription FROM PersonGroupDesciption;", sqlConnection))
                {
                    using (var sqlReader = sqlCommand.ExecuteReader())
                    {
                        foreach (DbDataRecord record in sqlReader)
                        {
                            if (record != null && record["GroupPerson"].ToString().Length > 0)
                            {
                                listComboParameters.Add(record["GroupPerson"].ToString().Trim());
                            }
                        }
                    }
                }
                sqlConnection.Close();
            }
                        
            SettingsView(
                    "Отправитель", mailServerUserName, @"Отправитель рассылки в виде: Name@Domain.Subdomain",
                    "Получатель рассылки", "", @"Получатель рассылки в виде: Name@Domain.Subdomain",
                    "", "", "Неиспользуемое поле",
                    "Наименование", "", @"Краткое наименование рассылки",
                    "Описание", "", "Краткое описание рассылки",
                    "", "", "Неиспользуемое поле",
                    "Отчет", listComboParameters, "Варинт генерируемого отчета",
                    "Период", periodComboParameters, "Выбрать, за какой период делать отчет"
                    );
        }

        private void MailingSave(string recipientEmail, string senderEmail, string typeReport, string description, string schedule, string period)
        {
            bool recipientValid = false;
            bool senderValid = false;

            DateTime localDate = DateTime.Now;

            if (recipientEmail.Length > 0 && recipientEmail.Contains('.') && recipientEmail.Contains('@') && recipientEmail.Split('.').Count() > 1)
            { recipientValid = true; }

            if (senderEmail.Length > 0 && senderEmail.Contains('.') && senderEmail.Contains('@') && senderEmail.Split('.').Count() > 1)
            { senderValid = true; }

            _controlVisible(groupBoxProperties, false);

            if (databasePerson.Exists && typeReport.Length > 0 && senderValid && recipientValid)
            {
                using (SQLiteConnection sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();

                    SQLiteCommand sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();

                    using (SQLiteCommand sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'Mailing' (SenderEmail, RecipientEmail, Schedule, TypeReport, Description, DateCreated, Reserv1)" +
                               " VALUES (@SenderEmail, @RecipientEmail, @Schedule, @TypeReport, @Description, @DateCreated, @Reserv1)", sqlConnection))
                    {
                        sqlCommand.Parameters.Add("@SenderEmail", DbType.String).Value = senderEmail;
                        sqlCommand.Parameters.Add("@RecipientEmail", DbType.String).Value = recipientEmail;
                        sqlCommand.Parameters.Add("@Schedule", DbType.String).Value = schedule;
                        sqlCommand.Parameters.Add("@TypeReport", DbType.String).Value = typeReport;
                        sqlCommand.Parameters.Add("@Description", DbType.String).Value = description;
                        sqlCommand.Parameters.Add("@Reserv1", DbType.String).Value = period;
                        sqlCommand.Parameters.Add("@DateCreated", DbType.String).Value = localDate.ToString();

                        try { sqlCommand.ExecuteNonQuery(); } catch { }
                    }

                    sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                }
            }

            if (databasePerson.Exists && typeReport.Length > 0 && senderValid && recipientValid)
            {
                _toolStripStatusLabelSetText(StatusLabel2, "Добавлена рассылка: " + typeReport+"| Всего рассылок: "+_dataGridView1RowsCount());
            }
        }

        private void SettingsProgrammItem_Click(object sender, EventArgs e)
        {
            EnableMainMenuItems(false);
            _controlVisible(panelView, false);

            btnPropertiesSave.Text = "Сохранить настройки";
            SettingsView(
                "Server", sServer1, "Имя сервера \"Server\" с базой Intellect в виде - NameOfServer.Domain.Subdomain",
                "Имя администратора", sServer1UserName, "Имя администратора \"sa\" SQL-сервера",
                "Password", sServer1UserPassword, "Пароль администратора \"sa\" SQL-сервера \"Server\"",
                "MailServer", mailServer, "Имя почтового сервера \"Mail Server\" в виде - NameOfServer.Domain.Subdomain",
                "Sender's e-mail", mailServerUserName, "E-mail отправителя рассылок виде - User.Name@MailServer.Domain.Subdomain",
                "Sender's password", mailServerUserPassword, "Пароль E-mail отправителя почты",
                "", new List<string>(),"",
                "", new List<string>(), ""
                );
        }

        private void SettingsView(
            string label1, string txtbox1, string tooltip1,
            string label2, string txtbox2, string tooltip2,
            string label3, string txtboxPassword3, string tooltip3,
            string label4, string txtbox4, string tooltip4,
            string label5, string txtbox5, string tooltip5,
            string label6, string txtboxPassword6, string tooltip6,
            string nameLabel7, List<string> listStrings7, string tooltip7,
            string periodLabel8, List<string> periodStrings8, string tooltip8
            )
        {
            panelViewResize(numberPeopleInLoading);

            if (label1.Length > 0)
            {
                labelServer1 = new Label
                {
                    Text = label1,
                    BackColor = Color.PaleGreen,
                    Location = new Point(20, 60),
                    Size = new Size(590, 22),
                    BorderStyle = BorderStyle.None,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Parent = groupBoxProperties
                };
                textBoxServer1 = new TextBox
                {
                    Text = txtbox1,
                    Location = new Point(90, 61),
                    Size = new Size(90, 20),
                    BorderStyle = BorderStyle.FixedSingle,
                    Parent = groupBoxProperties
                };
                toolTip1.SetToolTip(textBoxServer1, tooltip1);
            }

            if (label2.Length > 0)
            {
                labelServer1UserName = new Label
                {
                    Text = label2,
                    BackColor = Color.PaleGreen,
                    Location = new Point(220, 61),
                    Size = new Size(70, 20),
                    BorderStyle = BorderStyle.None,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Parent = groupBoxProperties
                };
                textBoxServer1UserName = new TextBox
                {
                    Text = txtbox2,
                    //PasswordChar = '*',
                    Location = new Point(300, 61),
                    Size = new Size(90, 20),
                    BorderStyle = BorderStyle.FixedSingle,
                    Parent = groupBoxProperties
                };
                toolTip1.SetToolTip(textBoxServer1UserName, tooltip2);
            }

            if (label3.Length > 0)
            {

                labelServer1UserPassword = new Label
                {
                    Text = label3,
                    BackColor = Color.PaleGreen,
                    Location = new Point(420, 61),
                    Size = new Size(70, 20),
                    BorderStyle = BorderStyle.None,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Parent = groupBoxProperties
                };
                textBoxServer1UserPassword = new TextBox
                {
                    Text = txtboxPassword3,
                    PasswordChar = '*',
                    Location = new Point(500, 61),
                    Size = new Size(90, 20),
                    BorderStyle = BorderStyle.FixedSingle,
                    Parent = groupBoxProperties
                };
                toolTip1.SetToolTip(textBoxServer1UserPassword, tooltip3);
            }

            //new - mail server
            if (label4.Length > 0)
            {
                labelMailServerName = new Label
                {
                    Text = label4,
                    BackColor = Color.PaleGreen,
                    Location = new Point(20, 90),
                    Size = new Size(590, 22),
                    BorderStyle = BorderStyle.None,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Parent = groupBoxProperties
                };
                textBoxMailServerName = new TextBox
                {
                    Text = txtbox4,
                    Location = new Point(90, 91),
                    Size = new Size(90, 20),
                    BorderStyle = BorderStyle.FixedSingle,
                    Parent = groupBoxProperties
                };
                toolTip1.SetToolTip(textBoxMailServerName, tooltip4);
            }

            if (label5.Length > 0)
            {
                labelMailServerUserName = new Label
                {
                    Text = label5,
                    BackColor = Color.PaleGreen,
                    Location = new Point(220, 91),
                    Size = new Size(90, 22),
                    BorderStyle = BorderStyle.None,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Parent = groupBoxProperties
                };
                textBoxMailServerUserName = new TextBox
                {
                    Text = txtbox5,
                    Location = new Point(300, 91),
                    Size = new Size(90, 20),
                    BorderStyle = BorderStyle.FixedSingle,
                    Parent = groupBoxProperties
                };
                toolTip1.SetToolTip(textBoxMailServerUserName, label5);
            }

            if (label6.Length > 0)
            {
                labelMailServerUserPassword = new Label
                {
                    Text = label6,
                    BackColor = Color.PaleGreen,
                    Location = new Point(420, 91),
                    Size = new Size(70, 20),
                    BorderStyle = BorderStyle.None,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Parent = groupBoxProperties
                };
                textBoxMailServerUserPassword = new TextBox
                {
                    Text = txtboxPassword6,
                    PasswordChar = '*',
                    Location = new Point(500, 91),
                    Size = new Size(90, 20),
                    BorderStyle = BorderStyle.FixedSingle,
                    Parent = groupBoxProperties
                };
                toolTip1.SetToolTip(textBoxMailServerUserPassword, tooltip6);
            }

            if (listStrings7.Count > 1 && nameLabel7.Length > 0)
            {
                listComboLabel = new Label
                {
                    Text = nameLabel7,
                    BackColor = Color.PaleGreen,
                    Location = new Point(20, 120),
                    Size = new Size(590, 22),
                    BorderStyle = BorderStyle.None,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Parent = groupBoxProperties
                };

                listCombo = new ComboBox
                {
                    Location = new Point(90, 121),
                    Size = new Size(120, 20),
                    Parent = groupBoxProperties
                };
                toolTip1.SetToolTip(listCombo, tooltip7);

                listCombo.DataSource = listStrings7;
                listCombo.KeyPress += new KeyPressEventHandler(listCombo_KeyPress);
            }


            if (periodStrings8.Count > 1 && periodLabel8.Length > 0)
            {
                periodComboLabel = new Label
                {
                    Text = periodLabel8,
                    BackColor = Color.PaleGreen,
                    Location = new Point(150, 120),
                    Size = new Size(90, 22),
                    BorderStyle = BorderStyle.None,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Parent = groupBoxProperties
                };

                periodCombo = new ComboBox
                {
                    Location = new Point(190, 121),
                    Size = new Size(120, 20),
                    Parent = groupBoxProperties
                };
                toolTip1.SetToolTip(periodCombo, tooltip8);

                periodCombo.DataSource = periodStrings8;
                periodCombo.KeyPress += new KeyPressEventHandler(periodCombo_KeyPress);
            }

            labelServer1?.BringToFront();
            labelServer1UserName?.BringToFront();
            labelServer1UserPassword?.BringToFront();
            labelMailServerName?.BringToFront();
            labelMailServerUserName?.BringToFront();
            labelMailServerUserPassword?.BringToFront();

            textBoxServer1?.BringToFront();
            textBoxServer1UserName?.BringToFront();
            textBoxServer1UserPassword?.BringToFront();
            textBoxMailServerName?.BringToFront();
            textBoxMailServerUserName?.BringToFront();
            textBoxMailServerUserPassword?.BringToFront();
            listComboLabel?.BringToFront();
            listCombo?.BringToFront();

            groupBoxProperties.Visible = true;
        }

       private void periodCombo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)//если нажата Enter
            {
                periodCombo.Items.Add(periodCombo.Text);
            }
        }

        private void listCombo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)//если нажата Enter
            {
                listCombo.Items.Add(listCombo.Text);
            }
        }

        private void DisposeTemporaryControls()
        {
            _controlVisible(groupBoxProperties,false);
            _controlDispose(labelServer1);
            _controlDispose(labelServer1UserName);
            _controlDispose(labelServer1UserPassword);
            _controlDispose(labelMailServerName);
            _controlDispose(labelMailServerUserName);
            _controlDispose(labelMailServerUserPassword);

            _controlDispose(textBoxServer1);
            _controlDispose(textBoxServer1UserName);
            _controlDispose(textBoxServer1UserPassword);
            _controlDispose(textBoxMailServerName);
            _controlDispose(textBoxMailServerUserName);
            _controlDispose(textBoxMailServerUserPassword);

            _controlDispose(listComboLabel);
            _controlDispose(listCombo);

            _controlDispose(periodComboLabel);
            _controlDispose(periodCombo);

        }

        private void EnableMainMenuItems(bool enabled)
        {
            _MenuItemEnabled(SettingsMenuItem, enabled);
            _MenuItemEnabled(FunctionMenuItem, enabled);
            _MenuItemEnabled(GroupsMenuItem, enabled);
            _MenuItemEnabled(AnualDatesMenuItem, enabled);

            CheckBoxesFiltersAll_Enable(enabled);
        }

        private void buttonPropertiesCancel_Click(object sender, EventArgs e)
        {
            string btnName = btnPropertiesSave.Text;

            DisposeTemporaryControls();
            EnableMainMenuItems(true);

            if (btnName == @"Сохранить рассылку")
            {
                ShowDataTableQuery(databasePerson, "Mailing", "SELECT SenderEmail AS 'Отправитель', RecipientEmail AS 'Получатель', Schedule AS 'Отчет', " + 
                    "TypeReport AS 'Наименование', Description AS 'Описание', Period AS 'Период', DateCreated AS 'Дата создания/модификации' ", " ORDER BY RecipientEmail asc; ");
            }
        }
        
        private void buttonPropertiesSave_Click(object sender, EventArgs e) //PropertiesSave()
        {
            string btnName = btnPropertiesSave.Text.ToString();

            if (btnName == @"Сохранить настройки")
            { PropertiesSave(); }
            else if (btnName == @"Сохранить рассылку")
            {
                string recipientEmail = _textBoxReturnText(textBoxServer1UserName);
                string senderEmail = mailServerUserName;
                if (mailServerUserName.Length == 0)
                { senderEmail = _textBoxReturnText(textBoxServer1); }
                string typeReport = _textBoxReturnText(textBoxMailServerName);
                string description = _textBoxReturnText(textBoxMailServerUserName);
                string schedule = _comboBoxReturnSelected(listCombo);
                string period = _comboBoxReturnSelected(periodCombo);

                MailingSave(recipientEmail, senderEmail, typeReport, description, schedule, period);
                ShowDataTableQuery(databasePerson, "Mailing", "SELECT SenderEmail AS 'Отправитель', RecipientEmail AS 'Получатель', Schedule AS 'Отчет', " +
                    "TypeReport AS 'Наименование', Description AS 'Описание', Period AS 'Период', DateCreated AS 'Дата создания/модификации' ", " ORDER BY RecipientEmail asc; ");
            }

            DisposeTemporaryControls();
            EnableMainMenuItems(true);
            _controlVisible(panelView, true);
        }

        private void PropertiesSave() //Save Parameters into Registry and variables
        {
            string server = _textBoxReturnText(textBoxServer1);
            string user = _textBoxReturnText(textBoxServer1UserName);
            string password = _textBoxReturnText(textBoxServer1UserPassword);

            string stringMailServer = _textBoxReturnText(textBoxMailServerName);
            string stringMailUser = _textBoxReturnText(textBoxMailServerUserName);
            string stringMailUserpassword = _textBoxReturnText(textBoxMailServerUserPassword);

            CheckAliveServer(server, user, password);

            if (bServer1Exist)
            {
                _controlVisible(groupBoxProperties, false);

                sServer1 = server;
                sServer1UserName = user;
                sServer1UserPassword = password;

                mailServer = stringMailServer;
                mailServerUserName = stringMailUser;
                mailServerUserPassword = stringMailUserpassword;

                try
                {
                    using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(myRegKey))
                    {
                        EvUserKey.SetValue("SKDServer", server, Microsoft.Win32.RegistryValueKind.String);
                        try { EvUserKey.SetValue("SKDUser", EncryptStringToBase64Text(user, btsMess1, btsMess2), Microsoft.Win32.RegistryValueKind.String); } catch { }
                        try { EvUserKey.SetValue("SKDUserPassword", EncryptStringToBase64Text(password, btsMess1, btsMess2), Microsoft.Win32.RegistryValueKind.String); } catch { }

                        EvUserKey.SetValue("MailServer", mailServer, Microsoft.Win32.RegistryValueKind.String);
                        EvUserKey.SetValue("MailUser", mailServerUserName, Microsoft.Win32.RegistryValueKind.String);
                        try { EvUserKey.SetValue("MailUserPassword", EncryptStringToBase64Text(mailServerUserPassword, btsMess1, btsMess2), Microsoft.Win32.RegistryValueKind.String); } catch { }
                    }
                } catch { MessageBox.Show("Ошибки с доступом на запись в реестр. Данные сохранены не корректно."); }

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
                            sqlCommand.Parameters.Add("@EquipmentParameterName", DbType.String).Value = "SKDUser";
                            sqlCommand.Parameters.Add("@EquipmentParameterValue", DbType.String).Value = "SKDServer";
                            sqlCommand.Parameters.Add("@EquipmentParameterServer", DbType.String).Value = sServer1;
                            try { sqlCommand.Parameters.Add("@Reserv1", DbType.String).Value = EncryptStringToBase64Text(sServer1UserName, btsMess1, btsMess2); } catch { }
                            try { sqlCommand.Parameters.Add("@Reserv2", DbType.String).Value = EncryptStringToBase64Text(sServer1UserPassword, btsMess1, btsMess2); } catch { }
                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }

                        using (SQLiteCommand sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'EquipmentSettings' (EquipmentParameterName, EquipmentParameterValue, EquipmentParameterServer, Reserv1, Reserv2)" +
                                " VALUES (@EquipmentParameterName, @EquipmentParameterValue, @EquipmentParameterServer, @Reserv1, @Reserv2)", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@EquipmentParameterName", DbType.String).Value = "MailUser";
                            sqlCommand.Parameters.Add("@EquipmentParameterValue", DbType.String).Value = "MailServer";
                            sqlCommand.Parameters.Add("@EquipmentParameterServer", DbType.String).Value = mailServer;
                            sqlCommand.Parameters.Add("@Reserv1", DbType.String).Value = mailServerUserName;
                            try { sqlCommand.Parameters.Add("@Reserv2", DbType.String).Value = EncryptStringToBase64Text(mailServerUserPassword, btsMess1, btsMess2); } catch { }
                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }

                        sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                        sqlCommand1.ExecuteNonQuery();
                        sqlCommand1.Dispose();
                    }
                }
            }
            else
            {
                GetInfoSetup();
            }
        }

        private void ClearRegistryItem_Click(object sender, EventArgs e) //ClearRegistryData()
        { ClearRegistryData(); }

        private void ClearRegistryData()
        {
            try
            {
                using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(myRegKey))
                {
                    EvUserKey?.DeleteSubKey("SKDServer");
                    EvUserKey?.DeleteSubKey("SKDUser");
                    EvUserKey?.DeleteSubKey("SKDUserPassword");
                    EvUserKey?.DeleteSubKey("MailServer");
                    EvUserKey?.DeleteSubKey("MailUser");
                    EvUserKey?.DeleteSubKey("MailUserPassword");
                }
            } catch { MessageBox.Show("Ошибки с доступом у реестру на запись. Данные не удалены."); }
        }

        //--- End. Features of programm ---//



        /// //////////////// Start  DatagridView functions

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e) //dataGridView1CellClick()
        { dataGridView1CellClick(); }

        private void dataGridView1CellClick()
        {
            if (dataGridView1.Rows.Count > 0 && dataGridView1.CurrentRow.Index < dataGridView1.Rows.Count)
            {
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
                                if (dataGridView1.Columns[i].HeaderText.ToString() == "Время прихода ЧЧ:ММ" ||
                                    dataGridView1.Columns[i].HeaderText.ToString() == "Контрольное время")
                                    IndexColumn3 = i;
                                if (dataGridView1.Columns[i].HeaderText.ToString() == "Время ухода ЧЧ:ММ" ||
                                        dataGridView1.Columns[i].HeaderText.ToString() == "Уход с работы")
                                    IndexColumn4 = i;
                            }

                            textBoxFIO.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString();
                            textBoxNav.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn2].Value.ToString(); //Take the name of selected group
                            StatusLabel2.Text = @"Выбрана группа: " + textBoxGroup.Text + @" | Курсор на: " + textBoxFIO.Text;
                            groupBoxPeriod.BackColor = Color.PaleGreen;
                            decimal[] timeIn = ConvertStringTimeHHMMToDecimalArray(dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn3].Value.ToString());

                            numUpDownHourStart.Value = timeIn[0];
                            numUpDownMinuteStart.Value = timeIn[1];
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

                            textBoxFIO.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString();
                            textBoxNav.Text = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn2].Value.ToString(); //Take the name of selected group
                            StatusLabel2.Text = @"Выбран: " + textBoxFIO.Text + @" |  Всего ФИО: " + iFIO;

                            groupBoxPeriod.BackColor = Color.PaleGreen;
                            groupBoxTimeStart.BackColor = Color.PaleGreen;
                            groupBoxTimeEnd.BackColor = Color.PaleGreen;

                            groupBoxRemoveDays.BackColor = SystemColors.Control;

                            numUpDownHourStart.Value = 9;
                            numUpDownMinuteStart.Value = 0;
                            if (textBoxFIO.TextLength > 3)
                            {
                                comboBoxFio.SelectedIndex = comboBoxFio.FindString(textBoxFIO.Text);
                            }
                        }
                    }
                } catch (Exception expt)
                {
                    MessageBox.Show(expt.ToString());
                }
            }
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e) //SearchMembersSelectedGroup()
        { SearchMembersSelectedGroup(); }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e) //dataGridView1CellEndEdit()
        { dataGridView1CellEndEdit(); }

        private void dataGridView1CellEndEdit()
        {
            string fio = "";
            string nav = "";
            string[] timeIn = { "09", "00", "09:00" };
            string[] timeOut = { "18", "00", "18:00" };
            decimal[] timeInDecimal = { 9, 0, 09 };
            decimal[] timeOutDecimal = { 18, 0, 18 };
            string group = "";

            try
            {
                if (nameOfLastTableFromDB == "PersonGroup")
                {
                    int IndexCurrentRow = _dataGridView1CurrentRowIndex();
                    int IndexCurrentCollumn = _dataGridView1CurrentCollumnIndex();

                    int IndexColumn1 = -1;
                    int IndexColumn2 = -1;
                    int IndexColumn5 = -1;
                    int IndexColumn6 = -1;
                    int IndexColumn7 = -1;

                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        if (dataGridView1.Columns[i].HeaderText.ToString() == "Фамилия Имя Отчество")
                            IndexColumn1 = i;
                        else if (dataGridView1.Columns[i].HeaderText.ToString() == "NAV-код")
                            IndexColumn2 = i;
                        else if (dataGridView1.Columns[i].HeaderText.ToString() == "Группа")
                            IndexColumn5 = i;
                        else if (dataGridView1.Columns[i].HeaderText.ToString() == "Время прихода ЧЧ:ММ" ||
                                    dataGridView1.Columns[i].HeaderText.ToString() == "Контрольное время")
                            IndexColumn6 = i;
                        else if (dataGridView1.Columns[i].HeaderText.ToString() == "Время ухода ЧЧ:ММ" ||
                                        dataGridView1.Columns[i].HeaderText.ToString() == "Уход с работы")
                            IndexColumn7 = i;
                    }

                    fio = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString();
                    textBoxFIO.Text = fio;

                    nav = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn2].Value.ToString();
                    textBoxNav.Text = nav; //Take the name of selected group

                    group = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn5].Value.ToString();
                    textBoxGroup.Text = group; //Take the name of selected group

                    timeIn = ConvertDecimalTimeToStringHHMMArray(ConvertStringTimeHHMMToDecimal(dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn6].Value.ToString()));
                    timeInDecimal = ConvertStringTimeHHMMToDecimalArray(ConvertStringsTimeToStringHHMM(timeIn[0], timeIn[1]));

                    timeOut = ConvertDecimalTimeToStringHHMMArray(ConvertStringTimeHHMMToDecimal(dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn7].Value.ToString()));
                    timeOutDecimal = ConvertStringTimeHHMMToDecimalArray(ConvertStringsTimeToStringHHMM(timeOut[0], timeOut[1]));

                    using (var connection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                    {
                        connection.Open();

                        using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PersonGroup' (FIO, NAV, GroupPerson, HourControlling, MinuteControlling, Controlling, HourControllingOut, MinuteControllingOut, ControllingOut, ControllingHHMM, ControllingOUTHHMM) " +
                                                    "VALUES (@FIO, @NAV, @GroupPerson, @HourControlling, @MinuteControlling, @Controlling, @HourControllingOut, @MinuteControllingOut, @ControllingOut, @ControllingHHMM, @ControllingOUTHHMM)", connection))
                        {
                            command.Parameters.Add("@FIO", DbType.String).Value = fio;
                            command.Parameters.Add("@NAV", DbType.String).Value = nav;
                            command.Parameters.Add("@GroupPerson", DbType.String).Value = group;
                            command.Parameters.Add("@HourControlling", DbType.String).Value = timeIn[0];
                            command.Parameters.Add("@MinuteControlling", DbType.String).Value = timeIn[1];
                            command.Parameters.Add("@Controlling", DbType.Decimal).Value = timeInDecimal[2];

                            command.Parameters.Add("@HourControllingOut", DbType.String).Value = timeOut[0];
                            command.Parameters.Add("@MinuteControllingOut", DbType.String).Value = timeOut[1];
                            command.Parameters.Add("@ControllingOut", DbType.Decimal).Value = timeOutDecimal[2];

                            command.Parameters.Add("@ControllingHHMM", DbType.String).Value = timeIn[2];
                            command.Parameters.Add("@ControllingOUTHHMM", DbType.String).Value = timeOut[2];
                            try { command.ExecuteNonQuery(); } catch { }
                        }
                    }
                    SeekAndShowMembersOfGroup(group);

                    StatusLabel2.Text = @"Обновлено время прихода " + ShortFIO(textBoxFIO.Text) + " в группе: " + textBoxGroup.Text;
                }
                else if (nameOfLastTableFromDB == "Mailing")
                {
                    FindSelectedMailingOnDatagridviewAndSave();
                    ShowDataTableQuery(databasePerson, "Mailing", "SELECT SenderEmail AS 'Отправитель', RecipientEmail AS 'Получатель', Schedule AS 'Отчет', " +
                        "TypeReport AS 'Наименование', Description AS 'Описание', Period AS 'Период', DateCreated AS 'Дата создания/модификации' ", " ORDER BY RecipientEmail asc; ");
                }
            } catch
            {
                try
                {
                    if (nameOfLastTableFromDB == "PersonGroup")
                    {
                        int IndexCurrentRow = _dataGridView1CurrentRowIndex();
                        int IndexCurrentCollumn = _dataGridView1CurrentCollumnIndex();

                        int IndexColumn1 = -1;
                        int IndexColumn2 = -1;
                        int IndexColumn5 = -1;
                        int IndexColumn6 = -1;
                        int IndexColumn7 = -1;

                        for (int i = 0; i < dataGridView1.ColumnCount; i++)
                        {
                            if (dataGridView1.Columns[i].HeaderText.ToString() == "Фамилия Имя Отчество")
                                IndexColumn1 = i;
                            else if (dataGridView1.Columns[i].HeaderText.ToString() == "NAV-код")
                                IndexColumn2 = i;
                            else if (dataGridView1.Columns[i].HeaderText.ToString() == "Группа")
                                IndexColumn5 = i;
                            else if (dataGridView1.Columns[i].HeaderText.ToString() == "Время прихода ЧЧ:ММ" ||
                                    dataGridView1.Columns[i].HeaderText.ToString() == "Контрольное время")
                                IndexColumn6 = i;
                            else if (dataGridView1.Columns[i].HeaderText.ToString() == "Время ухода ЧЧ:ММ" ||
                                        dataGridView1.Columns[i].HeaderText.ToString() == "Уход с работы")
                                IndexColumn7 = i;
                        }
                        fio = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString();
                        textBoxFIO.Text = fio;

                        nav = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn2].Value.ToString();
                        textBoxNav.Text = nav; //Take the name of selected group

                        group = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn5].Value.ToString();
                        textBoxGroup.Text = group; //Take the name of selected group

                        timeIn = ConvertDecimalTimeToStringHHMMArray(ConvertStringTimeHHMMToDecimal(dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn6].Value.ToString()));
                        timeInDecimal = ConvertStringTimeHHMMToDecimalArray(ConvertStringsTimeToStringHHMM(timeIn[0], timeIn[1]));

                        timeOut = ConvertDecimalTimeToStringHHMMArray(ConvertStringTimeHHMMToDecimal(dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn7].Value.ToString()));
                        timeOutDecimal = ConvertStringTimeHHMMToDecimalArray(ConvertStringsTimeToStringHHMM(timeOut[0], timeOut[1]));

                        using (var connection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                        {
                            connection.Open();

                            using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PersonGroup' (FIO, NAV, GroupPerson, HourControlling, MinuteControlling, Controlling, HourControllingOut, MinuteControllingOut, ControllingOut, ControllingHHMM, ControllingOUTHHMM) " +
                                                        "VALUES (@FIO, @NAV, @GroupPerson, @HourControlling, @MinuteControlling, @Controlling, @HourControllingOut, @MinuteControllingOut, @ControllingOut, @ControllingHHMM, @ControllingOUTHHMM)", connection))
                            {
                                command.Parameters.Add("@FIO", DbType.String).Value = fio;
                                command.Parameters.Add("@NAV", DbType.String).Value = nav;
                                command.Parameters.Add("@GroupPerson", DbType.String).Value = group;
                                command.Parameters.Add("@HourControlling", DbType.String).Value = timeIn[0];
                                command.Parameters.Add("@MinuteControlling", DbType.String).Value = timeIn[1];
                                command.Parameters.Add("@Controlling", DbType.Decimal).Value = timeInDecimal[2];

                                command.Parameters.Add("@HourControllingOut", DbType.String).Value = timeOut[0];
                                command.Parameters.Add("@MinuteControllingOut", DbType.String).Value = timeOut[1];
                                command.Parameters.Add("@ControllingOut", DbType.Decimal).Value = timeOutDecimal[2];

                                command.Parameters.Add("@ControllingHHMM", DbType.String).Value = timeIn[2];
                                command.Parameters.Add("@ControllingOUTHHMM", DbType.String).Value = timeOut[2];
                                try { command.ExecuteNonQuery(); } catch { }
                            }
                        }

                        SeekAndShowMembersOfGroup(group);

                        StatusLabel2.Text = @"Обновлено время прихода " + ShortFIO(textBoxFIO.Text) + " в группе: " + textBoxGroup.Text;
                    }
                } catch { }
            }
        }

        private void FindSelectedMailingOnDatagridviewAndSave()
        {
            string recipientEmail = "";
            string senderEmail = "";
            string typeReport = "";
            string description = "";
            string schedule = "";
            string period = "";
            
            int IndexCurrentRow = _dataGridView1CurrentRowIndex();
            if (IndexCurrentRow > -1)
            {
                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    if (dataGridView1.Columns[i].HeaderText == @"Получатель" || dataGridView1.Columns[i].HeaderText == @"RecipientEmail")
                    {
                        recipientEmail = dataGridView1.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (dataGridView1.Columns[i].HeaderText == @"Отправитель" || dataGridView1.Columns[i].HeaderText == @"SenderEmail")
                    {
                        senderEmail = dataGridView1.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (dataGridView1.Columns[i].HeaderText == @"Наименование" || dataGridView1.Columns[i].HeaderText == @"TypeReport")
                    {
                        typeReport = dataGridView1.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (dataGridView1.Columns[i].HeaderText == @"Описание" || dataGridView1.Columns[i].HeaderText == @"Description")
                    {
                        description = dataGridView1.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (dataGridView1.Columns[i].HeaderText == @"Отчет" || dataGridView1.Columns[i].HeaderText == @"Schedule")
                    {
                        schedule = dataGridView1.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (dataGridView1.Columns[i].HeaderText == @"Период" || dataGridView1.Columns[i].HeaderText == @"Period")
                    {
                        period = dataGridView1.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                }
            }
            MailingSave(recipientEmail, senderEmail, typeReport, description, schedule, period);
        }

        //Show help to Edit on some collumns DataGridView
        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            DataGridViewCell cell;
            if (nameOfLastTableFromDB == "PersonGroup")
            {
                try
                {
                    if ((e.ColumnIndex == this.dataGridView1.Columns["Фамилия Имя Отчество"].Index) && e.Value != null)
                    {
                        cell = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                        cell.ToolTipText = cell.Value.ToString();
                    }
                    else if ((e.ColumnIndex == this.dataGridView1.Columns["NAV-код"].Index) && e.Value != null)
                    {
                        cell = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                        cell.ToolTipText = cell.Value.ToString();
                    }
                    else if ((e.ColumnIndex == this.dataGridView1.Columns["Контрольное время"].Index || e.ColumnIndex == this.dataGridView1.Columns["Время прихода ЧЧ:ММ"].Index) && e.Value != null)
                    {
                        cell = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                        cell.ToolTipText = "Для установки нового значения нажмите F2,\nвнесите новое значение,\nа затем нажмите Enter";
                    }
                    else if ((e.ColumnIndex == this.dataGridView1.Columns["Уход с работы"].Index || e.ColumnIndex == this.dataGridView1.Columns["Время ухода ЧЧ:ММ"].Index) && e.Value != null)
                    {
                        cell = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                        cell.ToolTipText = "Для установки нового значения нажмите F2,\nвнесите новое значение,\nа затем нажмите Enter";
                    }
                    else
                    {
                        try
                        {
                            cell = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                            cell.ToolTipText = cell.Value.ToString();
                        } catch { }
                    }
                } catch { }
            }
            if (nameOfLastTableFromDB == "Mailing" && e.RowIndex > -1 && e.ColumnIndex > -1 && e.Value != null)
            {
                cell = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                cell.ToolTipText = "Для установки нового значения нажмите F2,\nвнесите новое значение,\nа затем нажмите Enter";
            }
            cell = null;
        }

        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                int currentMouseOverRow = dataGridView1.HitTest(e.X, e.Y).RowIndex;
                if (( nameOfLastTableFromDB == "PersonGroupDesciption") && currentMouseOverRow > -1)
                {
                    ContextMenu mRightClick = new ContextMenu();
                    mRightClick.MenuItems.Add(new MenuItem("Удалить выделенную группу", DeleteCurrentRow));
                    mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                }
                else if ((nameOfLastTableFromDB == "Mailing" ) && currentMouseOverRow > -1)
                {
                    ContextMenu mRightClick = new ContextMenu();
                    mRightClick.MenuItems.Add(new MenuItem("Удалить выделенную рассылку", DeleteCurrentRow));
                    mRightClick.MenuItems.Add(new MenuItem("Выполнить выделенную рассылку", DoMainAction));

                    mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                }

            }
        }

        private void DoMainAction(object sender, EventArgs e) //DoMainAction()
        { DoMainAction(); }

        private void DoMainAction()
        {
            int IndexColumn1 = -1;           // индекс 1-й колонки в датагрид
            int IndexColumn2 = -1;           // индекс 2-й колонки в датагрид

            int IndexCurrentRow = _dataGridView1CurrentRowIndex();

            if (IndexCurrentRow > -1)
            {
                switch (nameOfLastTableFromDB)
                {
                    case "PersonGroupDesciption":
                        {
                            break;
                        }
                    case "PersonGroup" when textBoxGroup.Text.Trim().Length > 0:
                        {
                            break;
                        }
                    case "Mailing":
                        {
                            string typeReport = "";
                            for (int i = 0; i < dataGridView1.ColumnCount; i++)
                            {
                                if (dataGridView1.Columns[i].HeaderText == "Получатель" || dataGridView1.Columns[i].HeaderText == "RecipientEmail")
                                { IndexColumn1 = i; }
                                else if (dataGridView1.Columns[i].HeaderText == "Наименование" || dataGridView1.Columns[i].HeaderText == "TypeReport")
                                {
                                    IndexColumn2 = i;
                                    typeReport = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn2].Value.ToString();
                                }
                            }

                            /

                            if (IndexColumn1 > -1 && IndexColumn2 > -1)
                            {
                                DeleteDataTableQueryParameters(databasePerson, "Mailing",
                                    "RecipientEmail", dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString(),
                                    "TypeReport", typeReport);
                            }
                            _toolStripStatusLabelSetText(StatusLabel2, "Удалена рассылка отчета " + typeReport + "| Всего рассылок: " + _dataGridView1RowsCount());
                            ShowDataTableQuery(databasePerson, "Mailing", "SELECT SenderEmail AS 'Отправитель', RecipientEmail AS 'Получатель', Schedule AS 'Отчет', " +
                                "TypeReport AS 'Наименование', Description AS 'Описание', Period AS 'Период', DateCreated AS 'Дата создания/модификации' ", " ORDER BY RecipientEmail asc; ");
                            break;
                        }
                    default:
                        break;
                }
            }
        }

        private void DeleteCurrentRow(object sender, EventArgs e) //DeleteCurrentRow()
        { DeleteCurrentRow(); }

        private void DeleteCurrentRow()
        {
            int IndexColumn1 = -1;           // индекс 1-й колонки в датагрид
            int IndexColumn2 = -1;           // индекс 2-й колонки в датагрид

            int IndexCurrentRow = _dataGridView1CurrentRowIndex();

            if (IndexCurrentRow > -1)
            {
                switch (nameOfLastTableFromDB)
                {
                    case "PersonGroupDesciption":
                        {
                            string selectedGroup = "";
                            for (int i = 0; i < dataGridView1.ColumnCount; i++)
                            {
                                if (dataGridView1.Columns[i].HeaderText == "GroupPerson"|| dataGridView1.Columns[i].HeaderText == "Группа")
                                { IndexColumn1 = i; }
                            }

                            if (IndexColumn1 > -1)
                            {
                                selectedGroup = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString().Trim();
                                DeleteDataTableQueryParameters(databasePerson, "PersonGroup", "GroupPerson", selectedGroup);
                                DeleteDataTableQueryParameters(databasePerson, "PersonGroupDesciption", "GroupPerson", selectedGroup);
                            }
                            PersonOrGroupItem.Text = "Работа с одной персоной";

                            ListGroups();
                            MembersGroupItem.Enabled = true;
                            _toolStripStatusLabelSetText(StatusLabel2, "Удалена группа: " + selectedGroup + "| Всего групп: " + _dataGridView1RowsCount());
                            break;
                        }
                    case "PersonGroup" when textBoxGroup.Text.Trim().Length > 0:
                        {
                            DeleteDataTableQueryParameters(databasePerson, "PersonGroup", "GroupPerson", textBoxGroup.Text.Trim());
                            DeleteDataTableQueryParameters(databasePerson, "PersonGroupDesciption", "GroupPerson", textBoxGroup.Text.Trim());
                            textBoxGroup.BackColor = Color.White;
                            PersonOrGroupItem.Text = "Работа с одной персоной";

                            MembersGroupItem.Enabled = true;
                            ListGroups();
                            break;
                        }
                    case "Mailing":
                        {
                            string typeReport = "";
                            for (int i = 0; i < dataGridView1.ColumnCount; i++)
                            {
                                if (dataGridView1.Columns[i].HeaderText == "Получатель" || dataGridView1.Columns[i].HeaderText == "RecipientEmail")
                                { IndexColumn1 = i; }
                                else if (dataGridView1.Columns[i].HeaderText == "Наименование" || dataGridView1.Columns[i].HeaderText == "TypeReport")
                                {
                                    IndexColumn2 = i;
                                    typeReport = dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn2].Value.ToString();
                                }
                            }

                            if (IndexColumn1 > -1 && IndexColumn2 > -1)
                            {
                                DeleteDataTableQueryParameters(databasePerson, "Mailing",
                                    "RecipientEmail", dataGridView1.Rows[IndexCurrentRow].Cells[IndexColumn1].Value.ToString(),
                                    "TypeReport", typeReport);
                            }
                            _toolStripStatusLabelSetText(StatusLabel2, "Удалена рассылка отчета " + typeReport+"| Всего рассылок: "+_dataGridView1RowsCount());
                            ShowDataTableQuery(databasePerson, "Mailing", "SELECT SenderEmail AS 'Отправитель', RecipientEmail AS 'Получатель', Schedule AS 'Отчет', " +
                                "TypeReport AS 'Наименование', Description AS 'Описание', Period AS 'Период', DateCreated AS 'Дата создания/модификации' ", " ORDER BY RecipientEmail asc; ");
                            break;
                        }
                    default:
                        break;
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
            { StatusLabel2.Text = @"Создать группу: " + textBoxGroup.Text.Trim().ToString() + "(" + textBoxGroupDescription.Text.Trim() + ")"; }
            else
            { StatusLabel2.Text = @"Создать группу: " + textBoxGroup.Text.Trim().ToString(); }
        }

        private void NumUpDown_ValueChanged(object sender, EventArgs e) //numUpDownValueChanged()
        { numUpDownValueChanged(); }

        private void numUpDownValueChanged()
        {
            numUpHourStart = _numUpDownReturn(numUpDownHourStart);
            numUpMinuteStart = _numUpDownReturn(numUpDownMinuteStart);
            numUpHourEnd = _numUpDownReturn(numUpDownHourEnd);
            numUpMinuteEnd = _numUpDownReturn(numUpDownMinuteEnd);
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

        private void comboBoxFio_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)//если нажата Enter
            {
                comboBoxFio.Items.Add(comboBoxFio.Text);
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
        public string PositionInDepartment = "";

        public string GroupPerson = "Office";
        public string ControlInHour = "9";
        public decimal ControlInHourDecimal = 9;
        public string ControlInMinute = "0";
        public decimal ControlInMinuteDecimal = 0;
        public decimal ControlInDecimal = 9;
        public string ControlOutHour = "18";
        public decimal ControlOutHourDecimal = 18;
        public string ControlOutMinute = "0";
        public decimal ControlOutMinuteDecimal = 0;
        public decimal ControlOutDecimal = 18;
        public string ControlInHHMM = "09:00";
        public string ControlOutHHMM = "18:00";

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

        public bool Late = false;
        public bool Early = false;

        public string serverSKD = "";
        public string namePassPoint = "";
        public string directionPass = "";

        private decimal ConvertTwoStringsTimeToDecimal(string hour, string minute)
        {
            decimal result = TryParseStringToDecimal(hour) + TryParseStringToDecimal(TimeSpan.FromMinutes(TryParseStringToDouble(minute)).TotalHours.ToString());
            return result;
        }

        private decimal TryParseStringToDecimal(string str)  //string -> decimal. if error it will return 0
        {
            decimal result = 0;
            try { result = decimal.Parse(str); } catch { result = 0; }
            return result;
        }

        private double TryParseStringToDouble(string str)  //string -> decimal. if error it will return 0
        {
            double result = 0;
            try { result = double.Parse(str); } catch { }
            return result;
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
