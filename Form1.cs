using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
//using System.Data.SqlClient;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
//using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
//using NLog;
//in nuget console - 
//install-package nlog
//install-package nlog.config

// for Crypography
using System.Security.Cryptography;

namespace PersonViewerSCA2
{
    public partial class FormPersonViewerSCA :Form
    {
        private NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        private System.Diagnostics.FileVersionInfo myFileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath);
        private string productName = "";
        private string myRegKey = @"SOFTWARE\RYIK\PersonViewerSCA2";
        private string currentAction = "";
        private bool currentModeAppManual;

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
        private bool bServer1Exist = false;
        private Label labelServer1;
        private TextBox textBoxServer1;
        private Label labelServer1UserName;
        private TextBox textBoxServer1UserName;
        private Label labelServer1UserPassword;
        private TextBox textBoxServer1UserPassword;
        private string sServer1 = "";
        private string sServer1Registry = "";
        private string sServer1DB = "";
        private string sServer1UserName = "";
        private string sServer1UserNameRegistry = "";
        private string sServer1UserNameDB = "";
        private string sServer1UserPassword = "";
        private string sServer1UserPasswordRegistry = "";
        private string sServer1UserPasswordDB = "";
        
        private Label labelMailServerName;
        private TextBox textBoxMailServerName;
        private Label labelMailServerUserName;
        private TextBox textBoxMailServerUserName;
        private Label labelMailServerUserPassword;
        private TextBox textBoxMailServerUserPassword;
        private string mailServer = "";
        private string mailServerRegistry = "";
        private string mailServerDB = "";
        private string mailServerUserName = "";
        private string mailServerUserNameRegistry = "";
        private string mailServerUserNameDB = "";
        private string mailServerUserPassword = "";
        private string mailServerUserPasswordRegistry = "";
        private string mailServerUserPasswordDB = "";

        private Label labelmysqlServer;
        private TextBox textBoxmysqlServer;
        private Label labelmysqlServerUserName;
        private TextBox textBoxmysqlServerUserName;
        private Label labelmysqlServerUserPassword;
        private TextBox textBoxmysqlServerUserPassword;
        private string mysqlServer = "";
        private string mysqlServerRegistry = "";
        private string mysqlServerDB = "";
        private string mysqlServerUserName = "";
        private string mysqlServerUserNameRegistry = "";
        private string mysqlServerUserNameDB = "";
        private string mysqlServerUserPassword = "";
        private string mysqlServerUserPasswordRegistry = "";
        private string mysqlServerUserPasswordDB = "";
        
        private Color clrRealRegistration = Color.PaleGreen;
        private string sLastSelectedElement = "MainForm";

        //Settings of Programm

        private Label listComboLabel;
        private ComboBox listCombo = new ComboBox();

        private Label periodComboLabel;
        private ComboBox periodCombo = new ComboBox();

        private Label labelSettings9;
        private ComboBox comboSettings9 = new ComboBox();


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
                                  new DataColumn(@"Отсутствовал на работе",typeof(string)),     //34
                                  new DataColumn(@"Код",typeof(string)),                        //35
                                  new DataColumn(@"Вышестоящая группа",typeof(string)),         //36
                                  new DataColumn(@"Описание группы",typeof(string)),            //37
                                  new DataColumn(@"Комментарии",typeof(string))                 //38

                };
        public readonly string[] arrayAllColumnsDataTablePeople =
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
                                  @"Отсутствовал на работе",      //34
                                  @"Код",                           //35
                                  @"Вышестоящая группа",            //36
                                  @"Описание группы",                //37
                                  @"Комментарии"                    //38
        };
        public readonly string[] orderColumnsFinacialReport =
            {
                                  @"Фамилия Имя Отчество",//1
                                  @"NAV-код",//2
                                  @"Отдел",//11
                                  @"Дата регистрации",//12
                                  @"День недели",                    //32
                                  @"Время прихода ЧЧ:ММ",//22
                                  @"Время ухода ЧЧ:ММ",//23
                                  @"Реальное время прихода ЧЧ:ММ",//24
                                  @"Реальное время ухода ЧЧ:ММ", //25
                                  @"Реальное отработанное время ЧЧ:ММ", //27
                                  @"Опоздание",                    //28
                                  @"Ранний уход",                 //29
                                  @"Отпуск (отгул)",                 //30
                                  @"Коммандировка",                 //31
                                  @"Больничный",                    //33
                                  @"Отсутствовал на работе",      //34
                                  @"Комментарии"                    //38

        };
        public readonly string[] arrayHiddenColumnsFIO =
            {
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
                            @"Отсутствовал на работе",      //34
                            @"Комментарии",                    //38
                            @"Код",                           //35
                            @"Вышестоящая группа",            //36
                            @"Описание группы"                //37
         };

        private DataTable dtPersonTemp = new DataTable("PersonTemp");
        private DataTable dtPersonTempAllColumns = new DataTable("PersonTempAllColumns");
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
        {
            logger.Trace("Test trace message");
            logger.Debug("Test debug message");

            logger.Warn("Test warn message");
            logger.Error("Test error message");
            logger.Fatal("Test fatal message");
            bool existed;
            // получаем GIUD приложения
            string guid = System.Runtime.InteropServices.Marshal.GetTypeLibGuidForAssembly(System.Reflection.Assembly.GetExecutingAssembly()).ToString();
            System.Threading.Mutex mutexObj = new System.Threading.Mutex(true, guid, out existed);

            if (existed)
            {
                Form1Load();
            }
            else
            {
                ApplicationExit(); //Bolck other this app to run
                System.Threading.Thread.Sleep(1000);
                return;
            }
        }

        private void Form1Load()
        {
            currentModeAppManual = true;
            _MenuItemTextSet(modeItem, "Сменить на автоматический режим");
            _menuItemTooltipSet(modeItem, "Включен интеррактивный режим. Все рассылки остановлены.");

            myFileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath);
            strVersion = myFileVersionInfo.Comments + " ver." + myFileVersionInfo.FileVersion + " " + myFileVersionInfo.LegalCopyright;
            StatusLabel1.Text = myFileVersionInfo.ProductName + " ver." + myFileVersionInfo.FileVersion + " " + myFileVersionInfo.LegalCopyright;
            StatusLabel1.Alignment = ToolStripItemAlignment.Right;
            productName = myFileVersionInfo.ProductName;
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

            mysqlServer = mysqlServerRegistry.Length > 0 ? mysqlServerRegistry : mysqlServerDB;
            mysqlServerUserName = mysqlServerUserNameRegistry.Length > 0 ? mysqlServerUserNameRegistry : mysqlServerUserNameDB;
            mysqlServerUserPassword = mysqlServerUserPasswordRegistry.Length > 0 ? mysqlServerUserPasswordRegistry : mysqlServerUserPasswordDB;


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


            //Clone default column name and structure from 'dtPeople' to other DataTables
            dtPersonTemp = dtPeople.Clone();  //Copy only structure(Name of columns)
            dtPersonRegisteredFull = dtPeople.Clone();  //Copy only structure(Name of columns)
            dtPersonRegistered = dtPeople.Clone();  //Copy only structure(Name of columns)
            dtPersonGroup = dtPeople.Clone();  //Copy only structure(Name of columns)
            dtPersonsLastList = dtPeople.Clone();  //Copy only structure(Name of columns)
            dtPersonsLastComboList = dtPeople.Clone();  //Copy only structure(Name of columns)

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
                    "RecipientEmail TEXT, GroupsReport TEXT, NameReport TEXT, Description TEXT, DateCreated TEXT, Period TEXT, Status TEXT, Reserv1 TEXT, Reserv2 TEXT);", databasePerson);
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
            TryUpdateStructureSqlDB("Mailing", "SenderEmail TEXT, RecipientEmail TEXT, GroupsReport TEXT, NameReport TEXT, Description TEXT, DateCreated TEXT, Period TEXT, Status TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
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
                                    if (record != null && record["PersonsList"].ToString().Length > 0)
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
                                    if (record != null && record["ComboList"].ToString().Length > 0)
                                    {
                                        _comboBoxAdd(comboBoxFio, record["ComboList"].ToString().Trim());
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
                                    if (record != null && record["EquipmentParameterValue"].ToString().Length > 0)
                                    {
                                        if (record["EquipmentParameterValue"].ToString().Trim() == "SKDServer" && record["EquipmentParameterName"].ToString().Trim() == "SKDUser")
                                        {
                                            sServer1DB = record["EquipmentParameterServer"].ToString();
                                            try { sServer1UserNameDB = DecryptBase64ToString(record["Reserv1"].ToString(), btsMess1, btsMess2); } catch { }
                                            try { sServer1UserPasswordDB = DecryptBase64ToString(record["Reserv2"].ToString(), btsMess1, btsMess2); } catch { }
                                        }
                                       else if (record["EquipmentParameterValue"].ToString().Trim() == "MailServer" && record["EquipmentParameterName"].ToString().Trim() == "MailUser")
                                        {
                                            mailServerDB = record["EquipmentParameterServer"].ToString();
                                            mailServerUserNameDB = record["Reserv1"].ToString();
                                            try { mailServerUserPasswordDB = DecryptBase64ToString(record["Reserv2"].ToString(), btsMess1, btsMess2); } catch { }
                                        }
                                        else if (record["EquipmentParameterValue"].ToString().Trim() == "MySQLServer" && record["EquipmentParameterName"].ToString().Trim() == "MySQLUser")
                                        {
                                            mysqlServerDB = record["EquipmentParameterServer"].ToString();
                                            mysqlServerUserNameDB = record["Reserv1"].ToString();
                                            try { mysqlServerUserPasswordDB = DecryptBase64ToString(record["Reserv2"].ToString(), btsMess1, btsMess2); } catch { }
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

                try { mysqlServerRegistry = EvUserKey.GetValue("MySQLServer").ToString().Trim(); } catch { }
                try { mysqlServerUserNameRegistry = EvUserKey.GetValue("MySQLUser").ToString().Trim(); } catch { }
                try { mysqlServerUserPasswordRegistry = DecryptBase64ToString(EvUserKey.GetValue("MySQLUserPassword").ToString(), btsMess1, btsMess2).ToString().Trim(); } catch { }
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

        private void TryUpdateStructureSqlDB(string tableName, string listColumnsWithType, System.IO.FileInfo FileDB) //Update Table in DB and execute of SQL Query
        {
            using (var connection = new SQLiteConnection($"Data Source={databasePerson.FullName};Version=3;"))
            {
                connection.Open();
                foreach (string column in listColumnsWithType.Split(','))
                {
                    using (var command = new SQLiteCommand("ALTER TABLE " + tableName + " ADD COLUMN " + column, connection))
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

        private void ShowDatatableOnDatagridview(DataTable dt, string[] nameHidenColumnsArray) //Query data from the Table of the DB
        {
            DataTable dataTable = dt.Copy();
            for (int i = 0; i < nameHidenColumnsArray.Length; i++)
            {
                if (nameHidenColumnsArray[i] != null && nameHidenColumnsArray[i].Length > 0)
                    try { dataTable.Columns[nameHidenColumnsArray[i]].ColumnMapping = MappingType.Hidden; } catch { }
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

        private void DeleteDataTableQueryParameters(System.IO.FileInfo databasePerson, string myTable, string mySqlParameter1, string mySqlData1, string mySqlParameter2, string mySqlData2, string mySqlParameter3, string mySqlData3) //Delete data from the Table of the DB by NAV (both parameters are string)
        {
            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();

                    if (mySqlParameter1.Length > 0 && mySqlParameter2.Length > 0 && mySqlParameter3.Length > 0)
                    {
                        using (var sqlCommand = new SQLiteCommand("DELETE FROM '" + myTable + "' Where " + mySqlParameter1 + "= @" + mySqlParameter1 + " AND " + mySqlParameter2 + "= @" + mySqlParameter2 + " AND " + mySqlParameter3 + "= @" + mySqlParameter3 + ";", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@" + mySqlParameter1, DbType.String).Value = mySqlData1;
                            sqlCommand.Parameters.Add("@" + mySqlParameter2, DbType.String).Value = mySqlData2;
                            sqlCommand.Parameters.Add("@" + mySqlParameter3, DbType.String).Value = mySqlData3;

                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }
                    }
                    else if (mySqlParameter1.Length > 0 && mySqlParameter2.Length > 0)
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
                using (var sqlConnection = new System.Data.SqlClient.SqlConnection(stringConnection))
                {
                    sqlConnection.Open();
                    string query = "SELECT id FROM OBJ_PERSON ";
                    using (var sqlCommand = new System.Data.SqlClient.SqlCommand(query, sqlConnection))
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
            {
                _toolStripStatusLabelSetText(StatusLabel2, "Есть доступ к серверу " + serverName);
                _toolStripStatusLabelBackColor(StatusLabel2, SystemColors.Control);
                _MenuItemEnabled(GetFioItem, true);
            }
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


        private void GetFio_Click(object sender, EventArgs e)  //GetFIO()
        { GetFIO(); }

        private async void GetFIO()  // CheckAliveServer()   GetFioFromServers()  ImportTablePeopleToTableGroupsInLocalDB()
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
                _ProgressBar1Start();
                dataGridView1.Visible = true;
                pictureBox1.Visible = false;

                await Task.Run(() => GetFioFromServers(dtTempIntermediate));

                await Task.Run(() => ImportTablePeopleToTableGroupsInLocalDB(databasePerson.ToString(), dtTempIntermediate));

                //show selected data     
                //distinct Records                
                var namesDistinctColumnsArray = arrayAllColumnsDataTablePeople.Except(arrayHiddenColumnsFIO).ToArray(); //take distinct data
                dtPeople = GetDistinctRecords(dtTempIntermediate, namesDistinctColumnsArray);

                await Task.Run(() => ShowDatatableOnDatagridview(dtPeople, arrayHiddenColumnsFIO));

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
            DataRow row;
            string stringConnection;
            string query;

            List<string> ListFIOTemp = new List<string>();
            listFIO = new List<string>();
            try
            {
                _comboBoxClr(comboBoxFio);
                _toolStripStatusLabelSetText(StatusLabel2, "Запрашиваю списки персонала с " + sServer1 + ". Ждите окончания процесса...");
                stimerPrev = "Запрашиваю списки персонала с " + sServer1 + ". Ждите окончания процесса...";
                stringConnection = "Data Source=" + sServer1 + "\\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID=" + sServer1UserName + ";Password=" + sServer1UserPassword + "; Connect Timeout=60";
                using (var sqlConnection = new System.Data.SqlClient.SqlConnection(stringConnection))
                {
                    sqlConnection.Open();

                    //получить список пользователей с сервера
                    query = "SELECT id, name, surname, patronymic, post, tabnum, parent_id FROM OBJ_PERSON ";
                    using (var sqlCommand = new System.Data.SqlClient.SqlCommand(query, sqlConnection))
                    {
                        using (var sqlReader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in sqlReader)
                            {
                                _ProgressWork1Step(1);
                                try
                                {
                                    string id = record?["id"].ToString();
                                    if (record?["name"].ToString().Trim().Length > 0)
                                    {
                                        iFIO++;
                                        row = dataTablePeopple.NewRow();

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
                                _ProgressWork1Step(1);
                            }
                        }
                    }

                    //получить список департаментов с сервера
                    query = "SELECT id,level_id,name,owner_id,parent_id,region_id,schedule_id  FROM OBJ_DEPARTMENT";
                    using (var sqlCommand = new System.Data.SqlClient.SqlCommand(query, sqlConnection))
                    {
                        using (var sqlReader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in sqlReader)
                            {
                                _ProgressWork1Step(1);
                                try
                                {
                                    if (record?["name"].ToString().Trim().Length > 0)
                                    {
                                        row = dtGroup.NewRow();
                                        row[2] = record["id"].ToString().Trim();
                                        row[4] = record["name"].ToString().Trim();
                                        row[3] = record["parent_id"].ToString().Trim();

                                        dtGroup.Rows.Add(row);
                                    }
                                } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                                _ProgressWork1Step(1);
                            }
                        }
                    }
                }
                //todo import users and group from www.ais

                stringConnection = @"server=" + mysqlServer + @";User=" + mysqlServerUserName + @";Password=" + mysqlServerUserPassword + @";database=wwwais;pooling = false; convert zero datetime=True;Connect Timeout=60";
                logger.Info(stringConnection);
                logger.Info(query);
                using (var sqlConnection = new MySql.Data.MySqlClient.MySqlConnection(stringConnection))
                {
                    sqlConnection.Open();

                    query = "Select id,name,hourly,visibled_name FROM out_reasons";
                    using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(query, sqlConnection))
                    {
                        using (MySql.Data.MySqlClient.MySqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                if (reader.GetString(@"name") != null && reader.GetString(@"name").Length > 0)
                                {
                                    resons.Add(new OutReasons()
                                    {
                                        _id = Convert.ToInt32(reader.GetString(@"id")),
                                        _hourly = Convert.ToInt32(reader.GetString(@"hourly")),
                                        _name = reader.GetString(@"name"),
                                        _visibleName = reader.GetString(@"visibled_name")
                                    });
                                }
                            }
                        }
                    }
                    int idReason = 0;
                    string date = "";
                    string name = "";
                    query = "Select * FROM out_users where user_code='" + person.NAV + "' AND reason_date >= '" + startDate.Split(' ')[0] + "' AND reason_date <= '" + endDate.Split(' ')[0] + "' ";
                    using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(query, sqlConnection))
                    {
                        using (MySql.Data.MySqlClient.MySqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                if (reader.GetString(@"reason_id") != null && reader.GetString(@"reason_id").Length > 0)
                                {
                                    logger.Info(reader.GetString(@"reason_date"));
                                    try { idReason = Convert.ToInt32(reader.GetString(@"reason_id")); } catch { idReason = 0; }
                                    try { name = resons.Find((x) => x._id == idReason)._name; } catch { name = ""; }
                                    try { date = DateTime.Parse(reader.GetString(@"reason_date")).ToString("yyyy-MM-dd"); } catch { date = ""; }

                                    outPerson.Add(new OutPerson()
                                    {
                                        _reason_id = idReason,
                                        _reason_Name = name,
                                        _nav = reader.GetString(@"user_code"),
                                        _date = date,
                                        _from = ConvertStringsTimeToSeconds(reader.GetString(@"from_hour"), reader.GetString(@"from_min")),
                                        _to = ConvertStringsTimeToSeconds(reader.GetString(@"to_hour"), reader.GetString(@"to_min")),
                                        _hourly = 0
                                    });
                                }
                            }
                        }
                    }
                    sqlConnection.Close();
                }
                logger.Info(" всего записей: " + outPerson.Count);








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
                    DeleteDataTableQueryParameters(databasePerson, "PersonGroup", "GroupPerson", dr[2].ToString(), "", "", "", "");
                    DeleteDataTableQueryParameters(databasePerson, "PersonGroupDesciption", "GroupPerson", dr[2].ToString(), "", "", "", "");
                }
                foreach (var dr in dtGroup.AsEnumerable())
                {
                    CreateGroupInDB(databasePerson, dr[2].ToString(), dr[4].ToString());
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
                { _comboBoxAdd(comboBoxFio, str); }
                try
                { _comboBoxSelectIndex(comboBoxFio, 0); } catch { };

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
            _ProgressBar1Stop();
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
            await Task.Run(() => ShowDatatableOnDatagridview(dtPeople, arrayHiddenColumnsFIO));
        }




        /// <summary>
        /// ///////////////////////////////////////////////////////////////
        /// </summary>
        //check dubled function!!!!!!!!
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






        /*
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

            int iDGColumns = _dataGridView1ColumnCount();
            int iDGRows = _dataGridView1RowsCount();
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);           //Книга
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);    //Таблица.
            ExcelApp.Columns.ColumnWidth = iDGColumns;

            for (int i = 0; i < iDGColumns; i++)
            {
                ExcelApp.Cells[1, 1 + i] = _dataGridView1ColumnHeaderText(i);
                ExcelApp.Columns[1 + i].NumberFormat = "@";
                ExcelApp.Columns[1 + i].AutoFit();
            }

            for (int i = 0; i < iDGRows; i++)
            {
                for (int j = 0; j < iDGColumns; j++)
                { ExcelApp.Cells[i + 2, j + 1] = _dataGridView1CellValue(i, j); }
            }

            ExcelApp.Visible = true;      //Вызываем нашу созданную эксельку.
            ExcelApp.UserControl = true;

            _MenuItemBackColorChange(TableExportToExcelItem, SystemColors.Control);
            stimerPrev = "";
            _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);
            sLastSelectedElement = "ExportExcel";
            iDGColumns = 0; iDGRows = 0;
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
        */

        private string filePathApplication = Application.ExecutablePath;
        private string filePathExcelReport;

        private void ExportDatatableSelectedColumnsToExcel(DataTable dtExport, string nameReport, string filePath)  //Export DataTable to Excel 
        {
            try
            {
                int numberColumns = dtExport.Columns.Count;
                int[] indexColumns = new int[numberColumns];
                string[] nameColumns = new string[numberColumns];

                for (int i = 0; i < numberColumns; i++)
                {
                    nameColumns[i] = dtExport.Columns[i].ColumnName;
                    indexColumns[i] = dtExport.Columns.IndexOf(nameColumns[i]);
                }

                int rows = 1;
                int rowsInTable = dtExport.Rows.Count;
                int columnsInTable = indexColumns.Length - 1; // p.Length;
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application
                {
                    Visible = false, //делаем объект не видимым
                    DisplayAlerts = false, //не отображать диалоги о перезаписи
                    SheetsInNewWorkbook = 1//количество листов в книге
                };
                Microsoft.Office.Interop.Excel.Workbooks workbooks = excel.Workbooks;
                excel.Workbooks.Add(); //добавляем книгу
                Microsoft.Office.Interop.Excel.Workbook workbook = workbooks[1];
                Microsoft.Office.Interop.Excel.Sheets sheets = workbook.Worksheets;
                Microsoft.Office.Interop.Excel.Worksheet sheet = sheets.get_Item(1);
                sheet.Name = nameReport;
                //sheet.Names.Add("next", "=" + Path.GetFileNameWithoutExtension(filePathExcelReport) + "!$A$1", true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //  MessageBox.Show(columnsInTable.ToString());

                //colourize of columns
                sheet.Columns[columnsInTable].Interior.Color = Color.Silver;  // последняя колонка

                try
                {
                    sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(@"Фамилия Имя Отчество")) + 1)]
                    .Interior.Color = Color.DarkSeaGreen;
                } catch { } //"Фамилия Имя Отчество"

                try
                {
                    sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(@"Опоздание")) + 1)]
                    .Interior.Color = System.Drawing.Color.SandyBrown;
                } catch { } //"Опоздание"

                try
                {
                    sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(@"Ранний уход")) + 1)]
                    .Interior.Color = System.Drawing.Color.SandyBrown;
                } catch { } //"Ранний уход"

                for (int column = 0; column < columnsInTable; column++)
                {
                    //  MessageBox.Show( column.ToString());

                    sheet.Cells[column + 1].WrapText = true;
                    sheet.Cells[1, column + 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    sheet.Cells[1, column + 1].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
                    sheet.Cells[1, column + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    sheet.Cells[1, column + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    sheet.Cells[1, column + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    sheet.Cells[1, column + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    sheet.Cells[1, column + 1].Value = nameColumns[column];
                    sheet.Cells[1, column + 1].Interior.Color = System.Drawing.Color.Silver; //colourize the first row
                    sheet.Columns[column + 1].Font.Size = 8;
                    sheet.Columns[column + 1].Font.Name = "Tahoma";
                }

                //input data and set type of cells - numbers /text
                foreach (DataRow row in dtExport.Rows)
                {
                    rows++;
                    for (int column = 0; column < columnsInTable; column++)
                    {
                        // MessageBox.Show(rows+" "+ column);
                        sheet.Columns[column + 1].NumberFormat = "@";
                        sheet.Cells[rows, column + 1].Value = row[indexColumns[column]];
                        sheet.Cells[rows, column + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        sheet.Cells[rows, column + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                        sheet.Cells[rows, column + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        sheet.Cells[rows, column + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                        sheet.Columns[column + 1].AutoFit();
                    }
                }

                //Область сортировки          
                Microsoft.Office.Interop.Excel.Range range = sheet.Range["A2", GetExcelColumnName(columnsInTable) + (rows - 1)];

                //По какому столбцу сортировать
                string nameColumnSorted = GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(@"Фамилия Имя Отчество")) + 1);
                Microsoft.Office.Interop.Excel.Range rangeKey = sheet.Range[nameColumnSorted + (rowsInTable - 1)];

                //Добавляем параметры сортировки
                sheet.Sort.SortFields.Add(rangeKey);
                sheet.Sort.SetRange(range);
                sheet.Sort.Orientation = Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns;
                sheet.Sort.SortMethod = Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin;
                sheet.Sort.Apply();
                //Очищаем фильтр
                sheet.Sort.SortFields.Clear();

                //Autofilter
                range = sheet.UsedRange; //sheet.Cells.Range["A1", GetExcelColumnName(columnsInTable) + rowsInTable];
                range.Select();
                range.AutoFilter(1, Type.Missing, Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);

                workbook.SaveAs(filePath,
                    Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                    true, false, //save without asking
                    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
                    Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value
                    );
                workbook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                workbooks.Close();

                releaseObject(range);
                releaseObject(rangeKey);
                releaseObject(sheet);
                releaseObject(sheets);
                releaseObject(workbook);
                releaseObject(workbooks);
                excel.Quit();
                releaseObject(excel);
                indexColumns = null;
                nameColumns = null;
                //
                stimerPrev = "";
                _toolStripStatusLabelSetText(StatusLabel2, "Путь к отчету: " + filePath);
                _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);
            } catch (Exception expt)
            {
                logger.Error("ExportDatatableSelectedColumnsToExcel - " + expt.ToString());
            }
            logger.Info("Отчет сгенерирован " + nameReport + " и сохранен в " + filePath);

            _ProgressBar1Stop();
            sLastSelectedElement = "ExportExcel";
        }

        private void releaseObject(object obj) //for function - ExportToExcel()
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

        static string GetExcelColumnName(int number)
        {
            string result;
            if (number > 0)
            {
                int alphabets = (number - 1) / 26;
                int remainder = (number - 1) % 26;
                result = ((char)('A' + remainder)).ToString();
                if (alphabets > 0)
                    result = GetExcelColumnName(alphabets) + result;
            }
            else
                result = null;
            return result;
        }

        private async void Export_Click(object sender, EventArgs e)
        {
            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(SettingsMenuItem, false);
            _MenuItemEnabled(AnualDatesMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            _controlEnable(dataGridView1, false);

            _ProgressBar1Start();

            _toolStripStatusLabelSetText(StatusLabel2, "Генерирую Excel-файл");
            stimerPrev = "Наполняю файл данными из текущей таблицы";


            dtPersonTemp.SetColumnsOrder(orderColumnsFinacialReport);
            filePathExcelReport = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(filePathApplication), "InOutPeople" + @".xlsx");

            await Task.Run(() => ExportDatatableSelectedColumnsToExcel(dtPersonTemp, "InOutPeople", filePathExcelReport));
            //            await Task.Run(() => ExportDatagridToExcel());

            _MenuItemEnabled(FunctionMenuItem, true);
            _MenuItemEnabled(SettingsMenuItem, true);
            _MenuItemEnabled(AnualDatesMenuItem, true);
            _MenuItemEnabled(GroupsMenuItem, true);
            _controlEnable(dataGridView1, true);
            _ProgressBar1Stop();
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
                groupBoxFilterReport.BackColor = SystemColors.Control;
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

        private void CreateGroupInDB(System.IO.FileInfo fileInfo, string nameGroup, string descriptionGroup)
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
                _toolStripStatusLabelSetText(StatusLabel2, "Группа - \"" + nameGroup + "\" создана");
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
                groupBoxFilterReport.BackColor = SystemColors.Control;
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

            string[] nameHidenColumnsArray =
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
                                  @"Отсутствовал на работе",      //34
                                  @"Код",                           //35
                                  @"Вышестоящая группа",            //36
                                  @"Описание группы"                //37
                        };

            var namesDistinctCollumnsArray = arrayAllColumnsDataTablePeople.Except(nameHidenColumnsArray).ToArray(); //take distinct data
            dtPersonTemp = GetDistinctRecords(dtTemp, namesDistinctCollumnsArray);

            ShowDatatableOnDatagridview(dtPersonTemp, nameHidenColumnsArray);
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
                using (var sqlConnection = new System.Data.SqlClient.SqlConnection(stringConnection))
                {
                    sqlConnection.Open();
                    sqlQuery = "Select id, name FROM OBJ_ABC_ARC_READER;";
                    using (var sqlCommand = new System.Data.SqlClient.SqlCommand(sqlQuery, sqlConnection))
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
            _MenuItemBackColorChange(QuickLoadDataItem, SystemColors.Control);

            _MenuItemEnabled(QuickLoadDataItem, false);
            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(SettingsMenuItem, false);
            _MenuItemEnabled(AnualDatesMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            CheckBoxesFiltersAll_Enable(false);

            _controlVisible(dataGridView1, false);
            _controlVisible(pictureBox1, false);

            _ProgressBar1Start();

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


                dtPersonRegisteredFull.Clear();

                GetRegistrations(_textBoxReturnText(textBoxGroup), _dateTimePickerStart(), _dateTimePickerEnd(), "");

                dtPersonTemp = dtPersonRegisteredFull.Copy();
                dtPersonTempAllColumns = dtPersonRegisteredFull.Copy(); //store all columns

                string[] nameHidenColumnsArray =
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
                @"День недели",  //32
                @"Больничный",  //33
                @"Отсутствовал на работе",      //34
                @"Код",         //35
                @"Вышестоящая группа",            //36
                @"Описание группы"                //37
            };

                //show selected data     
                //distinct Records                
                var namesDistinctColumnsArray = arrayAllColumnsDataTablePeople.Except(nameHidenColumnsArray).ToArray(); //take distinct data
                dtPersonTemp = GetDistinctRecords(dtPersonTempAllColumns, namesDistinctColumnsArray);
                ShowDatatableOnDatagridview(dtPersonTemp, nameHidenColumnsArray);

                stimerPrev = "";

                _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);
                _MenuItemBackColorChange(QuickLoadDataItem, SystemColors.Control);
                _MenuItemBackColorChange(TableExportToExcelItem, Color.PaleGreen);

                _MenuItemEnabled(QuickLoadDataItem, true);
                _MenuItemEnabled(FunctionMenuItem, true);
                _MenuItemEnabled(SettingsMenuItem, true);
                _MenuItemEnabled(AnualDatesMenuItem, true);
                _MenuItemEnabled(GroupsMenuItem, true);
                _MenuItemEnabled(VisualModeItem, true);
                _MenuItemEnabled(VisualSelectColorMenuItem, true);
                _MenuItemEnabled(TableModeItem, true);
                _MenuItemEnabled(TableExportToExcelItem, true);

                _MenuItemVisible(VisualModeItem, true);
                _MenuItemVisible(VisualSelectColorMenuItem, true);
                _MenuItemVisible(TableModeItem, false);
                _MenuItemVisible(TableExportToExcelItem, true);
                _controlVisible(dataGridView1, true);

                _controlEnable(checkBoxReEnter, true);
                _controlEnable(checkBoxTimeViolations, false);
                _controlEnable(checkBoxWeekend, false);
                _controlEnable(checkBoxCelebrate, false);
                CheckBoxesFiltersAll_CheckedState(false);

                panelViewResize(numberPeopleInLoading);

                nameHidenColumnsArray = null;
                namesDistinctColumnsArray = null;
            }
            else
            {
                GetInfoSetup();
                _MenuItemEnabled(SettingsMenuItem, true);
            }
            _changeControlBackColor(groupBoxFilterReport, Color.PaleGreen);
            _ProgressBar1Stop();
        }

        private void GetRegistrations(string selectedGroup, string startDate, string endDate, string doPostAction)
        {
            Person person = new Person();

            decimal dControlHourIn = _numUpDownReturn(numUpDownHourStart);
            decimal dControlMinuteIn = _numUpDownReturn(numUpDownMinuteStart);
            decimal dControlHourOut = _numUpDownReturn(numUpDownHourEnd);
            decimal dControlMinuteOut = _numUpDownReturn(numUpDownMinuteEnd);

            if ((nameOfLastTableFromDB == "PersonGroupDesciption" || nameOfLastTableFromDB == "PersonGroup" || doPostAction == "sendEmail") && selectedGroup.Length > 0)
            {
                LoadGroupMembersFromDbToDataTable(dtPersonGroup, selectedGroup); //result will be in dtPersonGroup

                foreach (DataRow row in dtPersonGroup.Rows)
                {
                    if (row[1].ToString().Length > 0 && row[3].ToString() == selectedGroup)
                    {
                        person = new Person();
                        if (!(doPostAction == "sendEmail"))
                        {
                            _textBoxSetText(textBoxFIO, row[1].ToString());   //иммитируем выбор данных
                            _textBoxSetText(textBoxNav, row[2].ToString());   //Select person                  
                        }
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

                        GetPersonRegistrationFromServer(dtPersonRegisteredFull, person, startDate, endDate);     //Search Registration at checkpoints of the selected person
                    }
                }

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

                GetPersonRegistrationFromServer(dtPersonRegisteredFull, person, startDate, endDate);

                nameOfLastTableFromDB = "PersonRegistered";
                _toolStripStatusLabelSetText(StatusLabel2, "Данные с СКД по \"" + ShortFIO(_textBoxReturnText(textBoxFIO)) + "\" получены!");
            }

            person = null;
        }

        private void GetPersonRegistrationFromServer(DataTable dt, Person person, string startDate, string endDate)
        {
            DataRow rowPerson;
            string stringConnection = "";
            string query = "";

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
                    stringConnection = @"Data Source=" + sServer1 + @"\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID=" + sServer1UserName + @";Password=" + sServer1UserPassword + @";Connect Timeout=240";
                    using (var sqlConnection = new System.Data.SqlClient.SqlConnection(stringConnection))
                    {
                        sqlConnection.Open();
                        using (var cmd = new System.Data.SqlClient.SqlCommand("Select id, tabnum FROM OBJ_PERSON where tabnum like '%" + person.NAV + "%';", sqlConnection))
                        {
                            using (var reader = cmd.ExecuteReader())
                            {
                                foreach (DbDataRecord record in reader)
                                {
                                    _ProgressWork1Step(1);
                                    try
                                    {
                                        if (record?["tabnum"].ToString().Trim() == person.NAV)
                                        {
                                            stringIdCardIntellect = record["id"].ToString().Trim();
                                            person.idCard = Convert.ToInt32(record["id"].ToString().Trim());
                                            break;
                                        }
                                    }
                                    catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                                }
                            }
                        }
                    }
                }
                else if (person.NAV.Length != 6)
                {
                    foreach (var strRowWithNav in listFIO.ToArray())
                    {
                        _ProgressWork1Step(1);
                        if (strRowWithNav.Contains(person.NAV) && person.NAV.Length > 0 && strRowWithNav.Contains(sServer1))
                            try
                            {
                                stringIdCardIntellect = Regex.Split(strRowWithNav, "[|]")[3].Trim();
                                person.idCard = Convert.ToInt32(stringIdCardIntellect);
                            }
                            catch (Exception expt) { MessageBox.Show(expt.ToString()); }
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
                                }
                                catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                                try
                                {
                                    personNAVTemp = Regex.Split(strRowWithNav, "[|]")[4].Trim();
                                }
                                catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                                if (person.NAV.Length < 1 && personNAVTemp.Length > 0)
                                {
                                    person.NAV = personNAVTemp;
                                    _ProgressWork1Step(1); break;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception expt) { MessageBox.Show(expt.ToString()); }

            try
            {
                stringConnection = @"Data Source=" + sServer1 + @"\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID=" + sServer1UserName + ";Password=" + sServer1UserPassword + "; Connect Timeout=120";
                query = "SELECT param0, param1, objid, CONVERT(varchar, date, 120) AS date, CONVERT(varchar, PROTOCOL.time, 114) AS time FROM protocol " +
                   " where action like 'ACCESS_IN' AND param1 like '" + stringIdCardIntellect + "' AND date >= '" + startDate + "' AND date <= '" + endDate + "' " +
                   " ORDER BY date ASC";
                string stringDataNew = "";
                int idCardIntellect = 0;

                decimal hourManaging = 0;
                decimal minuteManaging = 0;
                decimal managingHours = 0;

                try { idCardIntellect = Convert.ToInt32(stringIdCardIntellect); } catch { idCardIntellect = 0; }
                if (idCardIntellect > 0)
                {
                    using (var sqlConnection = new System.Data.SqlClient.SqlConnection(stringConnection))
                    {
                        sqlConnection.Open();
                        using (var cmd = new System.Data.SqlClient.SqlCommand(query, sqlConnection))
                        {
                            using (var reader = cmd.ExecuteReader())
                            {
                                foreach (DbDataRecord record in reader)
                                {
                                    try
                                    {
                                        if (record != null && record["param0"].ToString().Trim().Length > 0)
                                        {
                                            stringDataNew = Regex.Split(record["date"].ToString().Trim(), "[ ]")[0];
                                            hourManaging = Convert.ToDecimal(Regex.Split(record["time"].ToString().Trim(), "[:]")[0]);
                                            minuteManaging = Convert.ToDecimal(Regex.Split(record["time"].ToString().Trim(), "[:]")[1]);
                                            managingHours = ConvertDecimalSeparatedTimeToDecimal(hourManaging, minuteManaging);
                                            person.idCard = Convert.ToInt32(record["param1"].ToString().Trim());
                                            listRegistrations.Add(
                                                person.FIO + "|" + person.NAV + "|" + record["param1"].ToString().Trim() + "|" + stringDataNew + "|" +
                                                hourManaging + "|" + minuteManaging + "|" + managingHours.ToString("#.###") + "|" +
                                                hourControlStart + "|" + minuteControlStart + "|" + controlStart.ToString("#.###") + "|" +
                                                sServer1 + "|" + record["objid"].ToString().Trim()
                                                );

                                            _ProgressWork1Step(1);
                                        }
                                    }
                                    catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                                }
                            }
                        }
                    }
                }

                stringDataNew = null; query = null; stringConnection = null;

                _ProgressWork1Step(1);
            }
            catch (Exception Expt)
            { MessageBox.Show(Expt.ToString(), @"Сервер не доступен, или неправильная авторизация", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            string[] cellData;
            string namePoint = "";
            string nameDirection = "";
            foreach (var rowData in listRegistrations.ToArray())
            {
                cellData = Regex.Split(rowData, "[|]");
                namePoint = "";
                nameDirection = "";

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
                        }
                        catch { }
                }

                rowPerson = dt.NewRow();
                rowPerson[@"Фамилия Имя Отчество"] = person.FIO;
                rowPerson[@"NAV-код"] = person.NAV;
                rowPerson[@"Группа"] = person.GroupPerson;
                rowPerson[@"№ пропуска"] = cellData[2];
                rowPerson[@"Время прихода,часы"] = person.ControlInHour;
                rowPerson[@"Время прихода,минут"] = person.ControlInMinute;
                rowPerson[@"Время прихода"] = controlStart;
                rowPerson[@"Время ухода,часы"] = person.ControlOutHour;
                rowPerson[@"Время ухода,минут"] = person.ControlOutMinute;
                rowPerson[@"Время ухода"] = person.ControlOutDecimal;
                //day of registration
                rowPerson[@"Дата регистрации"] = cellData[3];
                rowPerson[@"День недели"] = DayOfWeekRussian((DateTime.Parse(cellData[3])).DayOfWeek.ToString());
                //real data
                rowPerson[@"Время регистрации,часы"] = cellData[4];
                rowPerson[@"Время регистрации,минут"] = cellData[5];
                rowPerson[@"Время регистрации"] = TryParseStringToDecimal(cellData[6]);
                rowPerson[@"Сервер СКД"] = sServer1;
                rowPerson[@"Имя точки прохода"] = namePoint;
                rowPerson[@"Направление прохода"] = nameDirection;
                rowPerson[@"Время прихода ЧЧ:ММ"] = ConvertStringsTimeToStringHHMM(person.ControlInHour, person.ControlInMinute);
                rowPerson[@"Время ухода ЧЧ:ММ"] = ConvertStringsTimeToStringHHMM(person.ControlOutHour, person.ControlOutMinute);
                rowPerson[@"Реальное время прихода ЧЧ:ММ"] = ConvertStringsTimeToStringHHMM(cellData[4], cellData[5]);

                dt.Rows.Add(rowPerson);
            }
            if (listPoints.Count > 0)
            { bLoaded = true; }


            DataRow[] dtr;
            FindWorkDaysInSelected();

            // @"Отсутствовал на работе"      //34
            // рабочие дни в которые отсутствовал данная персона
            foreach (string day in workSelectedDays)
            {
                dtr = dt.Select(@"[Дата регистрации] = '" + day + "'");
                if (dtr.Count() == 0)
                {
                    rowPerson = dt.NewRow();
                    rowPerson[@"Фамилия Имя Отчество"] = person.FIO;
                    rowPerson[@"NAV-код"] = person.NAV;
                    rowPerson[@"Группа"] = person.GroupPerson;
                    rowPerson[@"№ пропуска"] = person.idCard;
                    rowPerson[@"Время прихода,часы"] = person.ControlInHour;
                    rowPerson[@"Время прихода,минут"] = person.ControlInMinute;
                    rowPerson[@"Время прихода"] = controlStart;
                    rowPerson[@"Время ухода,часы"] = person.ControlOutHour;
                    rowPerson[@"Время ухода,минут"] = person.ControlOutMinute;
                    rowPerson[@"Время ухода"] = person.ControlOutDecimal;
                    rowPerson[@"Время регистрации,часы"] = "0";
                    rowPerson[@"Время регистрации,минут"] = "0";
                    rowPerson[@"Время регистрации"] = "0";
                    rowPerson[@"Дата регистрации"] = day;
                    rowPerson[@"День недели"] = DayOfWeekRussian((DateTime.Parse(day)).DayOfWeek.ToString());
                    rowPerson[@"Время прихода ЧЧ:ММ"] = ConvertStringsTimeToStringHHMM(person.ControlInHour, person.ControlInMinute);
                    rowPerson[@"Время ухода ЧЧ:ММ"] = ConvertStringsTimeToStringHHMM(person.ControlOutHour, person.ControlOutMinute);
                    rowPerson[@"Отсутствовал на работе"] = "Да";

                    dt.Rows.Add(rowPerson);
                }
            }

            List<OutReasons> resons = new List<OutReasons>();
            resons.Add(new OutReasons()
            {
                _id = 0,
                _hourly = 1,
                _name = @"Unknow",
                _visibleName = @"Неизвестная"
            });

            List<OutPerson> outPerson = new List<OutPerson>();

            stringConnection = @"server=" + mysqlServer + @";User=" + mysqlServerUserName + @";Password=" + mysqlServerUserPassword + @";database=wwwais;pooling = false; convert zero datetime=True;Connect Timeout=60";
            logger.Info(stringConnection);
            using (var sqlConnection = new MySql.Data.MySqlClient.MySqlConnection(stringConnection))
            {
                sqlConnection.Open();
                query = "Select id,name,hourly,visibled_name FROM out_reasons";
                logger.Info(query);

                using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(query, sqlConnection))
                {
                    using (MySql.Data.MySqlClient.MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (reader.GetString(@"name") != null && reader.GetString(@"name").Length > 0)
                            {
                                resons.Add(new OutReasons()
                                {
                                    _id = Convert.ToInt32(reader.GetString(@"id")),
                                    _hourly = Convert.ToInt32(reader.GetString(@"hourly")),
                                    _name = reader.GetString(@"name"),
                                    _visibleName = reader.GetString(@"visibled_name")
                                });
                            }
                        }
                    }
                }
                int idReason = 0;
                string date = "";
                string name = "";
                query = "Select * FROM out_users where user_code='" + person.NAV + "' AND reason_date >= '" + startDate.Split(' ')[0] + "' AND reason_date <= '" + endDate.Split(' ')[0] + "' ";
                logger.Info(query);
                using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(query, sqlConnection))
                {
                    using (MySql.Data.MySqlClient.MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (reader.GetString(@"reason_id") != null && reader.GetString(@"reason_id").Length > 0)
                            {
                                logger.Info(reader.GetString(@"reason_date"));
                                try { idReason = Convert.ToInt32(reader.GetString(@"reason_id")); } catch { idReason = 0; }
                                try { name = resons.Find((x) => x._id == idReason)._name; } catch { name = ""; }
                                try { date = DateTime.Parse(reader.GetString(@"reason_date")).ToString("yyyy-MM-dd"); } catch { date = ""; }

                                outPerson.Add(new OutPerson()
                                {
                                    _reason_id = idReason,
                                    _reason_Name = name,
                                    _nav = reader.GetString(@"user_code"),
                                    _date = date,
                                    _from = ConvertStringsTimeToSeconds(reader.GetString(@"from_hour"), reader.GetString(@"from_min")),
                                    _to = ConvertStringsTimeToSeconds(reader.GetString(@"to_hour"), reader.GetString(@"to_min")),
                                    _hourly = 0
                                });
                            }
                        }
                    }
                }
                sqlConnection.Close();
            }
            logger.Info(person.NAV+ " - на сайте всего записей с отсутствиями: " + outPerson.Count);


            foreach (DataRow dr in dt.Rows) // search whole table
            {
                foreach (string day in workSelectedDays)
                {
                    if (dr[@"Дата регистрации"].ToString() == day) // if id==2
                    {
                        try { dr[@"Комментарии"] = outPerson.Find((x) => x._date == day)._reason_Name; } catch { }  // dr[@"Комментарии"] = "";
                    }
                }
            }


            listRegistrations.Clear(); rowPerson = null;
            namePoint = null; nameDirection = null;
            hourControlStart = 0; minuteControlStart = 0;
            stringIdCardIntellect = null; personNAVTemp = null; stringSelectedFIO = new string[1]; cellData = new string[1];
        }

        private string DayOfWeekRussian(string dayEnglish) //return short day of week in Russian
        {
            string result = "";
            switch (dayEnglish.ToLower())
            {
                case "monday":
                    result = "Пн";
                    break;
                case "tuesday":
                    result = "Вт";
                    break;
                case "wednesday":
                    result = "Ср";
                    break;
                case "thursday":
                    result = "Чт";
                    break;
                case "friday":
                    result = "Пт";
                    break;
                case "saturday":
                    result = "Сб";
                    break;
                case "sunday":
                    result = "Вс";
                    break;
                default:
                    result = dayEnglish;
                    break;
            }
            return result;
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
            dataRow = null;
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

            _ProgressBar1Start();
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
            _ProgressBar1Stop();
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

        private int _dataGridView1CurrentColumnIndex() //add string into  from other threads
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

        private void _comboBoxSelectIndex(ComboBox comboBx, int i) //from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { comboBx.SelectedIndex = i; }));
            else
                comboBx.SelectedIndex = i;
        }

        private string _comboBoxReturnSelected(ComboBox comboBox) //from other threads
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
                Invoke(new MethodInvoker(delegate
                {
                    StatusLabel2.Text = s;
                }));
            else
            {
                StatusLabel2.Text = s;
            }
            logger.Info(s);
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


        private void _MenuItemBackColorChange(ToolStripMenuItem tMenuItem, Color colorMenu) //add string into  from other threads
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


        private void _MenuItemTextSet(ToolStripMenuItem menuItem, string newTextControl) //access from other threads
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

        private void _menuItemTooltipSet(ToolStripMenuItem menuItem, string newTooltip) //access from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    menuItem.ToolTipText = newTooltip;
                }));
            else
            {
                menuItem.ToolTipText = newTooltip;
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

        private void _ProgressWork1Step(int step) //add into progressBar Value 2 from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    if (ProgressBar1.Value > 99)
                    { ProgressBar1.Value = 0; }
                    ProgressBar1.Maximum = 100;
                    ProgressBar1.Value += step;
                }));
            else
            {
                if (ProgressBar1.Value > 99)
                { ProgressBar1.Value = 0; }
                ProgressBar1.Maximum = 100;
                ProgressBar1.Value += step;
            }
        }

        private void _ProgressBar1Start() //Set progressBar Value into 0 from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    timer1.Enabled = true;
                    ProgressBar1.Value = 0;
                    timer1.Enabled = true;
                }));
            else
            {
                timer1.Enabled = true;
                ProgressBar1.Value = 0;
            }
        }

        private void _ProgressBar1Stop() //Set progressBar Value into 100 from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    timer1.Stop();
                    ProgressBar1.Value = 100;
                    StatusLabel1.ForeColor = Color.Black;
                }));
            else
            {
                timer1.Stop();
                ProgressBar1.Value = 100;
                StatusLabel1.ForeColor = Color.Black;
            }
        }

        //---- End. Access to Controls from other threads ----//



        //---- Start. Convertors of data types ----//

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

        private int ConvertStringsTimeToSeconds(string hour, string minute)
        {
            int h = 0;
            int m = 0;
            try { h = Convert.ToInt32(hour); } catch { }
            try { m = Convert.ToInt32(minute); } catch { }
            int result = h * 60 * 60 + m * 60;
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

        //---- End. Convertors of data types ----//




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
        { await Task.Run(() => checkBoxCheckStateChanged()); }

        private void checkBoxCheckStateChanged()
        {
            CheckBoxesFiltersAll_Enable(false);
            _controlVisible(dataGridView1, false);
            _controlVisible(pictureBox1, false);

            string nameGroup = _textBoxReturnText(textBoxGroup);

            DataTable dtTempIntermediate = dtPersonRegisteredFull.Clone();
            dtPersonTempAllColumns = dtPersonRegisteredFull.Clone();
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

            string[] arrayHiddenColumns =
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
                @"Реальное отработанное время", //26
                @"Код",         //35
                @"Вышестоящая группа",            //36
                @"Описание группы"                //37            
            };

            //store all columns
            dtPersonTempAllColumns = dtTempIntermediate.Copy();
            //show selected data     
            //distinct Records         

            var namesDistinctColumnsArray = arrayAllColumnsDataTablePeople.Except(arrayHiddenColumns).ToArray(); //take distinct data
            dtPersonTemp = GetDistinctRecords(dtTempIntermediate, namesDistinctColumnsArray);
            ShowDatatableOnDatagridview(dtPersonTemp, arrayHiddenColumns);

            panelViewResize(numberPeopleInLoading);
            _controlVisible(dataGridView1, true);

            //change enabling of checkboxes
            if (_CheckboxCheckedStateReturn(checkBoxReEnter))// if (checkBoxReEnter.Checked)
            {
                _controlEnable(checkBoxTimeViolations, true);
                _controlEnable(checkBoxWeekend, true);
                _controlEnable(checkBoxCelebrate, true);

                if (_CheckboxCheckedStateReturn(checkBoxTimeViolations))  // if (checkBoxStartWorkInTime.Checked)
                { _MenuItemBackColorChange(QuickLoadDataItem, SystemColors.Control); }
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

                if (_CheckboxCheckedStateReturn(checkBoxReEnter) || currentAction == "sendEmail") //checkBoxReEnter.Checked
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

                if (_CheckboxCheckedStateReturn(checkBoxWeekend) || currentAction == "sendEmail")//checkBoxWeekend Checking
                {
                    DeleteAnualDatesFromDataTables(dtTemp, personNAV,
                        dateTimePickerStart.Value.Year, dateTimePickerStart.Value.Month, dateTimePickerStart.Value.Day,
                        dateTimePickerEnd.Value.Year, dateTimePickerEnd.Value.Month, dateTimePickerEnd.Value.Day);
                }

                if (_CheckboxCheckedStateReturn(checkBoxTimeViolations)) //checkBoxStartWorkInTime Checking
                { QueryDeleteDataFromDataTable(dtTemp, "[Опоздание]='' AND [Ранний уход]=''", personNAV.NAV); }

                foreach (DataRow dr in dtTemp.AsEnumerable())
                { dataTableForStoring.ImportRow(dr); }

                allWorkedDaysPerson = null;
            } catch (Exception expt) { MessageBox.Show(expt.ToString()); }

            stringHourMinuteFirstRegistrationInDay = null; stringHourMinuteLastRegistrationInDay = null; hsDays = null;
            rowDtStoring = null; dtTemp = null; dtAllRegistrationsInSelectedDay = null;
        }

        private void DeleteAnualDatesFromDataTables(DataTable dt, Person person, int startYear, int startMonth, int startDay, int endYear, int endMonth, int endDay) //Exclude Anual Days from the table "PersonTemp" DB
        {
            var oneDay = TimeSpan.FromDays(1);

            var mySelectedStartDay = new DateTime(startYear, startMonth, startDay);
            var mySelectedEndDay = new DateTime(endYear, endMonth, endDay);
            //  int myYearNow = DateTime.Now.Year;
            var myMonthCalendar = new MonthCalendar();

            myMonthCalendar.MaxSelectionCount = 60;
            myMonthCalendar.SelectionRange = new SelectionRange(mySelectedStartDay, mySelectedEndDay);
            myMonthCalendar.FirstDayOfWeek = Day.Monday;

            for (int year = 0; year < 4; year++)
            {
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear - year, 1, 1));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear - year, 1, 2));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear - year, 3, 8));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear - year, 5, 1));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear - year, 5, 2));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear - year, 5, 9));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear - year, 6, 28));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear - year, 8, 24));    // (plavayuschaya data)
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear - year, 10, 16));   // (plavayuschaya data)
            }

            // Алгоритм для вычисления католической Пасхи http://snippets.dzone.com/posts/show/765
            int Y = startYear;
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
            myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear, monthEaster, dayEaster) + oneDay);
            //Independence day
            DateTime dayBolded = new DateTime(startYear, 8, 24);
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Sunday:
                    myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear, 8, 24) + oneDay);    // (plavayuschaya data)
                    break;
                default:
                    break;
            }
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Saturday:
                    myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear, 8, 24) + oneDay + oneDay);    // (plavayuschaya data)
                    break;
                default:
                    break;
            }
            //day of Ukraine Force
            dayBolded = new DateTime(startYear, 10, 16);
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Sunday:
                    myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear, 10, 16) + oneDay);    // (plavayuschaya data)
                    break;
                default:
                    break;
            }
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Saturday:
                    myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear, 10, 16) + oneDay + oneDay);    // (plavayuschaya data)
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
            myMonthCalendar.Dispose();
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

        private void FindWorkDaysInSelected() //
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
            boldeddDates.Add(dayBolded.ToString("yyyy-MM-dd"));
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Saturday:
                    myTempDate = new DateTime(myYearNow, 8, 24);
                    boldeddDates.Add(myTempDate.AddDays(2).ToString("yyyy-MM-dd"));
                    break;
                case (int)Day.Sunday:
                    myTempDate = new DateTime(myYearNow, 8, 24);
                    boldeddDates.Add(myTempDate.AddDays(1).ToString("yyyy-MM-dd"));
                    break;
                default:
                    break;
            }


            //day of Ukraine Force
            dayBolded = new DateTime(myYearNow, 10, 16);
            boldeddDates.Add(dayBolded.ToString("yyyy-MM-dd"));
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Saturday:
                    myTempDate = new DateTime(myYearNow, 10, 16);
                    boldeddDates.Add(myTempDate.AddDays(2).ToString("yyyy-MM-dd"));
                    break;
                case (int)Day.Sunday:
                    myTempDate = new DateTime(myYearNow, 10, 16);
                    boldeddDates.Add(myTempDate.AddDays(1).ToString("yyyy-MM-dd"));
                    break;
                default:
                    break;
            }

            //New Year day
            dayBolded = new DateTime(myYearNow, 1, 1);
            boldeddDates.Add(dayBolded.ToString("yyyy-MM-dd"));
            boldeddDates.Add(dayBolded.AddDays(1).ToString("yyyy-MM-dd"));
           switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Saturday:
                    myTempDate = new DateTime(myYearNow, 10, 16);
                    boldeddDates.Add(myTempDate.AddDays(2).ToString("yyyy-MM-dd"));
                    break;
                case (int)Day.Sunday:
                    myTempDate = new DateTime(myYearNow, 10, 16);
                    boldeddDates.Add(myTempDate.AddDays(1).ToString("yyyy-MM-dd"));
                    break;
                default:
                    break;
            }
             
            //Cristmas day
            dayBolded = new DateTime(myYearNow, 7, 1);
            boldeddDates.Add(dayBolded.ToString("yyyy-MM-dd"));
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Saturday:
                    myTempDate = new DateTime(myYearNow, 7, 1);
                    boldeddDates.Add(myTempDate.AddDays(2).ToString("yyyy-MM-dd"));
                    break;
                case (int)Day.Sunday:
                    myTempDate = new DateTime(myYearNow, 7, 1);
                    boldeddDates.Add(myTempDate.AddDays(1).ToString("yyyy-MM-dd"));
                    break;
                default:
                    break;
            }

            //Troitsa
            dayBolded = new DateTime(myYearNow, monthEaster, dayEaster) + fiftyDays;
            boldeddDates.Add(dayBolded.ToString("yyyy-MM-dd"));
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Saturday:
                    myTempDate = new DateTime(myYearNow, monthEaster, dayEaster);
                    boldeddDates.Add(myTempDate.AddDays(52).ToString("yyyy-MM-dd"));
                    break;
                case (int)Day.Sunday:
                    myTempDate = new DateTime(myYearNow, monthEaster, dayEaster);
                    boldeddDates.Add(myTempDate.AddDays(51).ToString("yyyy-MM-dd"));
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



        //---- Start. Drawing ---//

        private void VisualItem_Click(object sender, EventArgs e) //FindWorkDaysInSelected() , DrawFullWorkedPeriodRegistration()
        {
            Person personVisual = new Person();
            if (bLoaded)
            {
                SelectPersonFromDataGrid(personVisual);
                dataGridView1.Visible = false;
                FindWorkDaysInSelected();
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
                var rowsPersonRegistrationsForDraw = dtPersonTempAllColumns.AsEnumerable();
                if (group.Length > 0)
                { rowsPersonRegistrationsForDraw = dtPersonTempAllColumns.Select("[Группа] = '" + group + "'").CopyToDataTable().AsEnumerable(); }
                else if (nav.Length == 6)
                { rowsPersonRegistrationsForDraw = dtPersonTempAllColumns.Select("[NAV-код] = '" + nav + "'").CopyToDataTable().AsEnumerable(); }
                else if (nav.Length != 6 && fio.Length > 1)
                { rowsPersonRegistrationsForDraw = dtPersonTempAllColumns.Select("[Фамилия Имя Отчество] = '" + fio + "'").CopyToDataTable().AsEnumerable(); }

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
            var rowsPersonRegistrationsForDraw = dtPersonTempAllColumns.AsEnumerable();
            if (group.Length > 0)
            { rowsPersonRegistrationsForDraw = dtPersonTempAllColumns.Select("[Группа] = '" + group + "'").CopyToDataTable().AsEnumerable(); }
            else if (nav.Length == 6)
            { rowsPersonRegistrationsForDraw = dtPersonTempAllColumns.Select("[NAV-код] = '" + nav + "'").CopyToDataTable().AsEnumerable(); }
            else if (nav.Length != 6 && fio.Length > 1)
            { rowsPersonRegistrationsForDraw = dtPersonTempAllColumns.Select("[Фамилия Имя Отчество] = '" + fio + "'").CopyToDataTable().AsEnumerable(); }

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

        //---- End. Drawing ---//




        //---- Start. Parameters of programm ---//

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

            List<string> listComboParameters9 = new List<string>();
            listComboParameters9.Add("Активная");
            listComboParameters9.Add("Неактивная");


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
                    "Отчет", listComboParameters, "Выполнить отчет по группам",
                    "Период", periodComboParameters, "Выбрать, за какой период делать отчет",
                    "Статус", listComboParameters9, "Статус рассылки",
                    "", "", "",
                    "", "", "",
                    "", "", ""
                    );
        }

        private void MailingSave(string recipientEmail, string senderEmail, string groupsReport, string nameReport, string description, string period, string status)
        {
            bool recipientValid = false;
            bool senderValid = false;

            DateTime localDate = DateTime.Now;

            if (recipientEmail.Length > 0 && recipientEmail.Contains('.') && recipientEmail.Contains('@') && recipientEmail.Split('.').Count() > 1)
            { recipientValid = true; }

            if (senderEmail.Length > 0 && senderEmail.Contains('.') && senderEmail.Contains('@') && senderEmail.Split('.').Count() > 1)
            { senderValid = true; }

            _controlVisible(groupBoxProperties, false);

            if (databasePerson.Exists && nameReport.Length > 0 && senderValid && recipientValid)
            {
                using (SQLiteConnection sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();

                    SQLiteCommand sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();

                    using (SQLiteCommand sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'Mailing' (SenderEmail, RecipientEmail, GroupsReport, NameReport, Description, Period, Status, DateCreated)" +
                               " VALUES (@SenderEmail, @RecipientEmail, @GroupsReport, @NameReport, @Description, @Period, @Status, @DateCreated)", sqlConnection))
                    {
                        sqlCommand.Parameters.Add("@SenderEmail", DbType.String).Value = senderEmail;
                        sqlCommand.Parameters.Add("@RecipientEmail", DbType.String).Value = recipientEmail;
                        sqlCommand.Parameters.Add("@GroupsReport", DbType.String).Value = groupsReport;
                        sqlCommand.Parameters.Add("@NameReport", DbType.String).Value = nameReport;
                        sqlCommand.Parameters.Add("@Description", DbType.String).Value = description;
                        sqlCommand.Parameters.Add("@Period", DbType.String).Value = period;
                        sqlCommand.Parameters.Add("@Status", DbType.String).Value = status;
                        sqlCommand.Parameters.Add("@DateCreated", DbType.String).Value = localDate.ToString();

                        try { sqlCommand.ExecuteNonQuery(); } catch { }
                    }

                    sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                }
            }

            if (databasePerson.Exists && nameReport.Length > 0 && senderValid && recipientValid)
            {
                _toolStripStatusLabelSetText(StatusLabel2, "Добавлена рассылка: " + nameReport + "| Всего рассылок: " + _dataGridView1RowsCount());
            }
        }

        private void SettingsProgrammItem_Click(object sender, EventArgs e)
        {
            EnableMainMenuItems(false);
            _controlVisible(panelView, false);

            btnPropertiesSave.Text = "Сохранить настройки";
            SettingsView(
                "Сервер СКД", sServer1, "Имя сервера \"Server\" с базой Intellect в виде - NameOfServer.Domain.Subdomain",
                "Имя пользователя", sServer1UserName, "Имя администратора \"sa\" SQL-сервера",
                "Пароль", sServer1UserPassword, "Пароль администратора \"sa\" SQL-сервера \"Server\"",
                "Почтовый сервер", mailServer, "Имя почтового сервера \"Mail Server\" в виде - NameOfServer.Domain.Subdomain",
                "e-mail пользователя", mailServerUserName, "E-mail отправителя рассылок виде - User.Name@MailServer.Domain.Subdomain",
                "Пароль", mailServerUserPassword, "Пароль E-mail отправителя почты",
                "", new List<string>(), "",
                "", new List<string>(), "",
                "", new List<string>(), "",

                "MySQLServer", mysqlServer, "Имя сервера \"MySQLServer\" с базой регистраций отпусков и проч. на вэбсайте компании в виде - NameOfServer.Domain.Subdomain",
                "Имя пользователя", mysqlServerUserName, "Имя пользователя MySQL-сервера",
                "Пароль", mysqlServerUserPassword, "Пароль пользователя MySQL-сервера \"MySQLServer\""
                );
        }

        private void SettingsView(
            string label1, string txtbox1, string tooltip1, string label2, string txtbox2, string tooltip2, string label3, string txtboxPassword3, string tooltip3,
            string label4, string txtbox4, string tooltip4, string label5, string txtbox5, string tooltip5, string label6, string txtboxPassword6, string tooltip6,
            string nameLabel7, List<string> listStrings7, string tooltip7, string periodLabel8, List<string> periodStrings8, string tooltip8, string label9, List<string> listStrings9, string tooltip9,
            string label10, string txtbox10, string tooltip10, string label11, string txtbox11, string tooltip11, string label12, string txtboxPassword12, string tooltip12
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
                    Location = new Point(220, 60),
                    Size = new Size(70, 22),
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
                    Location = new Point(420, 60),
                    Size = new Size(70, 22),
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
                    Location = new Point(220, 90),
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
                    Location = new Point(420, 90),
                    Size = new Size(70, 22),
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
                    Location = new Point(220, 120),
                    Size = new Size(90, 22),
                    BorderStyle = BorderStyle.None,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Parent = groupBoxProperties
                };

                periodCombo = new ComboBox
                {
                    Location = new Point(300, 121),
                    Size = new Size(120, 20),
                    Parent = groupBoxProperties
                };
                toolTip1.SetToolTip(periodCombo, tooltip8);

                periodCombo.DataSource = periodStrings8;
                periodCombo.KeyPress += new KeyPressEventHandler(periodCombo_KeyPress);
            }

            if (listStrings9.Count > 1 && label9.Length > 0)
            {
                labelSettings9 = new Label
                {
                    Text = label9,
                    BackColor = Color.PaleGreen,
                    Location = new Point(20, 150),
                    Size = new Size(590, 22),
                    BorderStyle = BorderStyle.None,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Parent = groupBoxProperties
                };
                comboSettings9 = new ComboBox
                {
                    Location = new Point(300, 151),
                    Size = new Size(120, 20),
                    Parent = groupBoxProperties
                };
                comboSettings9.DataSource = listStrings9;
                toolTip1.SetToolTip(comboSettings9, tooltip9);
            }

            if (label10.Length > 0)
            {
                labelmysqlServer = new Label
                {
                    Text = label10,
                    BackColor = Color.PaleGreen,
                    Location = new Point(20, 120),
                    Size = new Size(590, 22),
                    BorderStyle = BorderStyle.None,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Parent = groupBoxProperties
                };
                textBoxmysqlServer = new TextBox
                {
                    Text = txtbox10,
                    Location = new Point(90, 121),
                    Size = new Size(90, 20),
                    BorderStyle = BorderStyle.FixedSingle,
                    Parent = groupBoxProperties
                };
                toolTip1.SetToolTip(textBoxmysqlServer, tooltip10);
            }

            if (label11.Length > 0)
            {
               labelmysqlServerUserName = new Label
                {
                    Text = label11,
                    BackColor = Color.PaleGreen,
                    Location = new Point(220, 120),
                    Size = new Size(70, 22),
                    BorderStyle = BorderStyle.None,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Parent = groupBoxProperties
                };
               textBoxmysqlServerUserName = new TextBox
                {
                    Text = txtbox11,
                    //PasswordChar = '*',
                    Location = new Point(300, 121),
                    Size = new Size(90, 20),
                    BorderStyle = BorderStyle.FixedSingle,
                    Parent = groupBoxProperties
                };
                toolTip1.SetToolTip(textBoxmysqlServerUserName, tooltip11);
            }

            if (label12.Length > 0)
            {

               labelmysqlServerUserPassword = new Label
                {
                    Text = label12,
                    BackColor = Color.PaleGreen,
                    Location = new Point(420, 120),
                    Size = new Size(70, 22),
                    BorderStyle = BorderStyle.None,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Parent = groupBoxProperties
                };
               textBoxmysqlServerUserPassword = new TextBox
                {
                    Text = txtboxPassword12,
                    PasswordChar = '*',
                    Location = new Point(500, 121),
                    Size = new Size(90, 20),
                    BorderStyle = BorderStyle.FixedSingle,
                    Parent = groupBoxProperties
                };
                toolTip1.SetToolTip(textBoxmysqlServerUserPassword, tooltip12);
            }




            labelServer1?.BringToFront();
            labelServer1UserName?.BringToFront();
            labelServer1UserPassword?.BringToFront();
            labelMailServerName?.BringToFront();
            labelMailServerUserName?.BringToFront();
            labelMailServerUserPassword?.BringToFront();
            listComboLabel?.BringToFront();
            labelmysqlServer?.BringToFront();
            labelmysqlServerUserName?.BringToFront();
            labelmysqlServerUserPassword?.BringToFront();


            textBoxServer1?.BringToFront();
            textBoxServer1UserName?.BringToFront();
            textBoxServer1UserPassword?.BringToFront();
            textBoxMailServerName?.BringToFront();
            textBoxMailServerUserName?.BringToFront();
            textBoxMailServerUserPassword?.BringToFront();
            listCombo?.BringToFront();
            textBoxmysqlServer?.BringToFront();
            textBoxmysqlServerUserName?.BringToFront();
            textBoxmysqlServerUserPassword?.BringToFront();

            periodComboLabel?.BringToFront();
            periodCombo?.BringToFront();

            labelSettings9?.BringToFront();
            comboSettings9?.BringToFront();

            groupBoxProperties.Visible = true;
        }
        
        private void buttonPropertiesCancel_Click(object sender, EventArgs e)
        {
            string btnName = btnPropertiesSave.Text;

            DisposeTemporaryControls();

            if (btnName == @"Сохранить рассылку")
            {
                ShowDataTableQuery(databasePerson, "Mailing", "SELECT SenderEmail AS 'Отправитель', RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', " +
                    "NameReport AS 'Наименование', Description AS 'Описание', Period AS 'Период', DateCreated AS 'Дата создания/модификации', Status AS 'Статус' ", " ORDER BY RecipientEmail asc, DateCreated desc; ");
            }
            EnableMainMenuItems(true);
            _controlVisible(panelView, true);
        }

        private void DisposeTemporaryControls()
        {
            _controlVisible(groupBoxProperties, false);
            _controlDispose(labelServer1);
            _controlDispose(labelServer1UserName);
            _controlDispose(labelServer1UserPassword);
            _controlDispose(labelMailServerName);
            _controlDispose(labelMailServerUserName);
            _controlDispose(labelMailServerUserPassword);
            _controlDispose(labelmysqlServer);
            _controlDispose(labelmysqlServerUserName);
            _controlDispose(labelmysqlServerUserPassword);

            _controlDispose(textBoxServer1);
            _controlDispose(textBoxServer1UserName);
            _controlDispose(textBoxServer1UserPassword);
            _controlDispose(textBoxMailServerName);
            _controlDispose(textBoxMailServerUserName);
            _controlDispose(textBoxMailServerUserPassword);
            _controlDispose(textBoxmysqlServer);
            _controlDispose(textBoxmysqlServerUserName);
            _controlDispose(textBoxmysqlServerUserPassword);
            
            _controlDispose(listComboLabel);
            _controlDispose(periodComboLabel);
            _controlDispose(labelSettings9);

            _controlDispose(listCombo);
            _controlDispose(periodCombo);
            _controlDispose(comboSettings9);
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
                string nameReport = _textBoxReturnText(textBoxMailServerName);
                string description = _textBoxReturnText(textBoxMailServerUserName);
                string report = _comboBoxReturnSelected(listCombo);
                string period = _comboBoxReturnSelected(periodCombo);
                string status = _comboBoxReturnSelected(comboSettings9);

                MailingSave(recipientEmail, senderEmail, report, nameReport, description, period, status);
                ShowDataTableQuery(databasePerson, "Mailing", "SELECT SenderEmail AS 'Отправитель', RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', " +
                    "NameReport AS 'Наименование', Description AS 'Описание', Period AS 'Период', DateCreated AS 'Дата создания/модификации', Status AS 'Статус' ", " ORDER BY RecipientEmail asc, DateCreated desc; ");
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

            string sMailServer = _textBoxReturnText(textBoxMailServerName);
            string sMailUser = _textBoxReturnText(textBoxMailServerUserName);
            string sMailUserPassword = _textBoxReturnText(textBoxMailServerUserPassword);

            string sMySqlServer = _textBoxReturnText(textBoxmysqlServer);
            string sMySqlServerUser = _textBoxReturnText(textBoxmysqlServerUserName);
            string sMySqlServerUserPassword = _textBoxReturnText(textBoxmysqlServerUserPassword);
            
            CheckAliveServer(server, user, password);

            if (bServer1Exist)
            {
                _controlVisible(groupBoxProperties, false);

                sServer1 = server;
                sServer1UserName = user;
                sServer1UserPassword = password;

                mailServer = sMailServer;
                mailServerUserName = sMailUser;
                mailServerUserPassword = sMailUserPassword;

                mysqlServer = sMySqlServer;
                 mysqlServerUserName= sMySqlServerUser;
                 mysqlServerUserPassword = sMySqlServerUserPassword;

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

                        EvUserKey.SetValue("MySQLServer",mysqlServer, Microsoft.Win32.RegistryValueKind.String);
                        EvUserKey.SetValue("MySQLUser", mysqlServerUserName, Microsoft.Win32.RegistryValueKind.String);
                        try { EvUserKey.SetValue("MySQLUserPassword", EncryptStringToBase64Text(mysqlServerUserPassword, btsMess1, btsMess2), Microsoft.Win32.RegistryValueKind.String); } catch { }
                    }

                } catch { logger.Error("CreateSubKey. Ошибки с доступом на запись в реестр. Данные сохранены не корректно."); }

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

                        using (SQLiteCommand sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'EquipmentSettings' (EquipmentParameterName, EquipmentParameterValue, EquipmentParameterServer, Reserv1, Reserv2)" +
                                " VALUES (@EquipmentParameterName, @EquipmentParameterValue, @EquipmentParameterServer, @Reserv1, @Reserv2)", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@EquipmentParameterName", DbType.String).Value = "MySQLUser";
                            sqlCommand.Parameters.Add("@EquipmentParameterValue", DbType.String).Value = "MySQLServer";
                            sqlCommand.Parameters.Add("@EquipmentParameterServer", DbType.String).Value = mysqlServer;
                            sqlCommand.Parameters.Add("@Reserv1", DbType.String).Value = mysqlServerUserName;
                            try { sqlCommand.Parameters.Add("@Reserv2", DbType.String).Value = EncryptStringToBase64Text(mysqlServerUserPassword, btsMess1, btsMess2); } catch { }
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



        //--- Start. Behaviour Controls ---//

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

        private void comboBoxFio_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)//если нажата Enter
            {
                comboBoxFio.Items.Add(comboBoxFio.Text);
            }
        }


        private void EnableMainMenuItems(bool enabled)
        {
            _MenuItemEnabled(SettingsMenuItem, enabled);
            _MenuItemEnabled(FunctionMenuItem, enabled);
            _MenuItemEnabled(GroupsMenuItem, enabled);
            _MenuItemEnabled(AnualDatesMenuItem, enabled);

            CheckBoxesFiltersAll_Enable(enabled);
        }

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

        private void NumUpDown_ValueChanged(object sender, EventArgs e) //numUpDownValueChanged()
        { NumUpDownValueChanged(); }

        private void NumUpDownValueChanged()
        {
            numUpHourStart = _numUpDownReturn(numUpDownHourStart);
            numUpMinuteStart = _numUpDownReturn(numUpDownMinuteStart);
            numUpHourEnd = _numUpDownReturn(numUpDownHourEnd);
            numUpMinuteEnd = _numUpDownReturn(numUpDownMinuteEnd);
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
                _MenuItemTextSet(PersonOrGroupItem, "Работа с группой");
                nameOfLastTableFromDB = "PersonGroup";
                isPerson = false;
            }
            else
            {
                _MenuItemTextSet(PersonOrGroupItem, "Работа с одной персоной");
                nameOfLastTableFromDB = "PersonRegistered";
                isPerson = true;
            }
        }

        //--- End. Behaviour Controls ---//



        //---  Start.  DatagridView functions ---//

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
                            groupBoxFilterReport.BackColor = SystemColors.Control;
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

                            groupBoxFilterReport.BackColor = SystemColors.Control;

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
            _ProgressBar1Start();

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
                    int IndexCurrentColumn = _dataGridView1CurrentColumnIndex();

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
                    DataGridViewSeekValuesInSelectedRow dgSeek = new DataGridViewSeekValuesInSelectedRow();
                    dgSeek.FindValueInCells(dataGridView1, new string[] { @"Получатель", @"Отправитель", @"Отчет по группам", @"Наименование", @"Описание", @"Период", @"Статус" });
                    _toolStripStatusLabelSetText(StatusLabel2, "Готовлю отчет " + dgSeek.values[3]);
                    stimerPrev = "";
                    MailingAction("sendEmail", dgSeek.values[0], dgSeek.values[0], dgSeek.values[2], dgSeek.values[3], dgSeek.values[4], dgSeek.values[5], dgSeek.values[6]);

                    //  ShowDataTableQuery(databasePerson, "Mailing", "SELECT SenderEmail AS 'Отправитель', RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', " +
                    //      "NameReport AS 'Наименование', Description AS 'Описание', Period AS 'Период', DateCreated AS 'Дата создания/модификации', Status AS 'Статус'  ", " ORDER BY RecipientEmail asc; ");
                }
            } catch
            {
                try
                {
                    if (nameOfLastTableFromDB == "PersonGroup")
                    {
                        int IndexCurrentRow = _dataGridView1CurrentRowIndex();
                        int IndexCurrentColumn = _dataGridView1CurrentColumnIndex();

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

            _ProgressBar1Stop();
        }

        //Show help to Edit on some columns DataGridView
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

        private void dataGridView1_MouseClick(object sender, MouseEventArgs e) //right click of mouse on the datagridview
        {
            if (e.Button == MouseButtons.Right)
            {
                int currentMouseOverRow = dataGridView1.HitTest(e.X, e.Y).RowIndex;
                if ((nameOfLastTableFromDB == "PersonGroupDesciption") && currentMouseOverRow > -1)
                {
                    ContextMenu mRightClick = new ContextMenu();
                    mRightClick.MenuItems.Add(new MenuItem("Удалить выделенную группу", DeleteCurrentRow));
                    mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                }
                else if ((nameOfLastTableFromDB == "Mailing") && currentMouseOverRow > -1)
                {
                    string mailing = "";
                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        if (dataGridView1.Columns[i].HeaderText == "Наименование")
                        {
                            mailing = dataGridView1.Rows[currentMouseOverRow].Cells[i].Value.ToString();
                        }
                    }
                    ContextMenu mRightClick = new ContextMenu();
                    mRightClick.MenuItems.Add(new MenuItem("Удалить рассылку: " + mailing, DeleteCurrentRow));
                    mRightClick.MenuItems.Add(new MenuItem("Выполнить рассылку: " + mailing, DoMainAction));

                    mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                }
            }
        }

        private void DoMainAction(object sender, EventArgs e) //DoMainAction()
        { DoMainAction(); }

        private void DoMainAction()
        {
            _ProgressBar1Start();

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
                case "Mailing": //send report by e-mail
                    {
                        //текущий режим работы приложения
                        currentAction = "sendEmail";

                        DataGridViewSeekValuesInSelectedRow dgSeek = new DataGridViewSeekValuesInSelectedRow();
                        dgSeek.FindValueInCells(dataGridView1, new string[] { @"Получатель", @"Отправитель", @"Отчет по группам", @"Наименование", @"Описание", @"Период", @"Статус" });
                        _toolStripStatusLabelSetText(StatusLabel2, "Готовлю отчет " + dgSeek.values[3]);
                        stimerPrev = "";

                        MailingAction("sendEmail", dgSeek.values[0], dgSeek.values[1], dgSeek.values[2], dgSeek.values[3], dgSeek.values[4], dgSeek.values[5], dgSeek.values[6]);
                        break;
                    }
                default:
                    break;
            }

            _ProgressBar1Stop();
        }

        private void DeleteCurrentRow(object sender, EventArgs e) //DeleteCurrentRow()
        { DeleteCurrentRow(); }

        private void DeleteCurrentRow()
        {
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
                                if (dataGridView1.Columns[i].HeaderText == "GroupPerson" || dataGridView1.Columns[i].HeaderText == "Группа")
                                {
                                    selectedGroup = dataGridView1.Rows[IndexCurrentRow].Cells[i].Value.ToString().Trim();
                                    DeleteDataTableQueryParameters(databasePerson, "PersonGroup", "GroupPerson", selectedGroup, "", "", "", "");
                                    DeleteDataTableQueryParameters(databasePerson, "PersonGroupDesciption", "GroupPerson", selectedGroup, "", "", "", "");
                                }
                            }

                            PersonOrGroupItem.Text = "Работа с одной персоной";

                            ListGroups();
                            MembersGroupItem.Enabled = true;
                            _toolStripStatusLabelSetText(StatusLabel2, "Удалена группа: " + selectedGroup + "| Всего групп: " + _dataGridView1RowsCount());
                            break;
                        }
                    case "PersonGroup" when textBoxGroup.Text.Trim().Length > 0:
                        {
                            DeleteDataTableQueryParameters(databasePerson, "PersonGroup", "GroupPerson", textBoxGroup.Text.Trim(), "", "", "", "");
                            DeleteDataTableQueryParameters(databasePerson, "PersonGroupDesciption", "GroupPerson", textBoxGroup.Text.Trim(), "", "", "", "");
                            textBoxGroup.BackColor = Color.White;
                            PersonOrGroupItem.Text = "Работа с одной персоной";

                            MembersGroupItem.Enabled = true;
                            ListGroups();
                            break;
                        }
                    case "Mailing":
                        {
                            DataGridViewSeekValuesInSelectedRow dgSeek = new DataGridViewSeekValuesInSelectedRow();
                            dgSeek.FindValueInCells(dataGridView1, new string[] { @"Получатель", @"Отправитель", @"Отчет по группам", @"Наименование", @"Описание", @"Период", @"Статус" });
                            DeleteDataTableQueryParameters(databasePerson, "Mailing",
                                "GroupsReport", dgSeek.values[2],
                                "NameReport", dgSeek.values[3],
                                "Description", dgSeek.values[4]);
                            _toolStripStatusLabelSetText(StatusLabel2, "Удалена рассылка отчета " + dgSeek.values[2] + "| Всего рассылок: " + _dataGridView1RowsCount());
                            ShowDataTableQuery(databasePerson, "Mailing", "SELECT SenderEmail AS 'Отправитель', RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', " +
                                "NameReport AS 'Наименование', Description AS 'Описание', Period AS 'Период', DateCreated AS 'Дата создания/модификации', Status AS 'Статус' ", " ORDER BY RecipientEmail asc; ");
                            break;
                        }
                    default:
                        break;
                }
            }
        }

        //---  End.  DatagridView functions ---//



        //---  Start. Schedule Functions ---//

        private async void ModeAppItem_Click(object sender, EventArgs e) //InitScheduleTask(); 
        {
            await Task.Run(() => InitScheduleTask());
        }

        /*
        // SetUpTimer(new TimeSpan(1, 1, 0, 0)); //runs on 1st at 1:00:00 
        private void SetUpTimer(TimeSpan alertTime)
        {
            DateTime current = DateTime.Now;
            TimeSpan timeToGo = alertTime - current.TimeOfDay;
            if (timeToGo < TimeSpan.Zero)
            {
                return;//time already passed 
            }
            timer = new System.Threading.Timer(x =>
             {
                 SelectMailingDoAction();
             }, null, timeToGo, System.Threading.Timeout.InfiniteTimeSpan);
        }
        */

        static System.Threading.Timer timer;
        static object synclock = new object();
        static bool sent = false;

        public void InitScheduleTask() //ScheduleTask()
        {
            long interval = 1 * 60 * 1000; //1 minute
            if (currentModeAppManual)
            {
                currentModeAppManual = false;
                currentAction = "sendEmail";
                _MenuItemTextSet(modeItem, "Сменить режим на интеррактивный");
                _menuItemTooltipSet(modeItem, "Включен автоматический режим. Выполняются Активные рассылки из БД.");
                _MenuItemBackColorChange(modeItem, Color.DarkOrange);

                _toolStripStatusLabelSetText(StatusLabel2, "Режим авторассылки отчетов");
                _toolStripStatusLabelBackColor(StatusLabel2, Color.PaleGreen);

                timer = new System.Threading.Timer(new System.Threading.TimerCallback(ScheduleTask), null, 0, interval);
            }
            else
            {
                currentModeAppManual = true;
                _MenuItemTextSet(modeItem, "Сменить на автоматический режим");
                _menuItemTooltipSet(modeItem, "Включен интеррактивный режим. Все рассылки остановлены.");
                _MenuItemBackColorChange(modeItem, SystemColors.Control);

                _toolStripStatusLabelSetText(StatusLabel2, "Интеррактивный режим");
                _toolStripStatusLabelBackColor(StatusLabel2, SystemColors.Control);
                timer.Dispose();
            }
        }

        private void ScheduleTask(object obj) //SelectMailingDoAction()
        {
            lock (synclock)
            {
                DateTime dd = DateTime.Now;
                if (dd.Minute == 5 && sent == false) //do something at Hour 1
                {
                    _toolStripStatusLabelSetText(StatusLabel2, "Ведется работа по подготовке отчетов " + DateTime.Now.ToString());
                    _toolStripStatusLabelBackColor(StatusLabel2, Color.LightPink);
                    SelectMailingDoAction();
                    sent = true;
                }
                else if (dd.Minute != 5)
                {
                    sent = false;
                    _toolStripStatusLabelSetText(StatusLabel2, "Режим почтовых рассылок. " + DateTime.Now.ToString());
                    _toolStripStatusLabelBackColor(StatusLabel2, Color.LightCyan);
                }
            }
        }

        private void SelectMailingDoAction() //MailingAction()
        {
            string sender = "";
            string recipient = "";
            string gproupsReport = "";
            string nameReport = "";
            string descriptionReport = "";
            string period = "";
            string status = "";
            List<MailingStructure> mailingList = new List<MailingStructure>();

            using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
            {
                sqlConnection.Open();

                using (var sqlCommand = new SQLiteCommand("SELECT SenderEmail, RecipientEmail, GroupsReport, NameReport, Description, Period, Status, DateCreated FROM Mailing;", sqlConnection))
                {
                    using (var reader = sqlCommand.ExecuteReader())
                    {
                        foreach (DbDataRecord record in reader)
                        {
                            try
                            {
                                if (record != null && record["SenderEmail"].ToString().Length > 0 && record["SenderEmail"].ToString().Length > 0)
                                {
                                    sender = record["SenderEmail"].ToString();
                                    recipient = record["RecipientEmail"].ToString();
                                    gproupsReport = record["GroupsReport"].ToString();
                                    nameReport = record["NameReport"].ToString();
                                    descriptionReport = record["Description"].ToString();
                                    period = record["Period"].ToString();
                                    status = record["Status"].ToString().Trim().ToLower();

                                    if (status == "активная")
                                    {
                                        mailingList.Add(new MailingStructure()
                                        {
                                            _sender = sender,
                                            _recipient = recipient,
                                            _groupsReport = gproupsReport,
                                            _nameReport = nameReport,
                                            _descriptionReport = descriptionReport,
                                            _period = period,
                                            _status = status
                                        });
                                    }
                                }
                            } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                        }
                    }
                }
            }

            //текущий режим работы приложения
            currentAction = "sendEmail";

            foreach (MailingStructure mailng in mailingList)
            {
                _toolStripStatusLabelSetText(StatusLabel2, "Готовлю отчет " + mailng._nameReport);
                stimerPrev = "";

                MailingAction("sendEmail", mailng._recipient, mailng._sender, mailng._groupsReport, mailng._nameReport, mailng._descriptionReport, mailng._period, mailng._status);
            }

            sender = null; recipient = null; gproupsReport = null; nameReport = null; descriptionReport = null; period = null; status = null;
            mailingList = null;
        }

        private void MailingAction(string mainAction, string recipientEmail, string senderEmail, string groupsReport, string nameReport, string description, string period, string status)
        {
            switch (mainAction)
            {
                case "saveEmail":
                    {
                        MailingSave(recipientEmail, senderEmail, groupsReport, nameReport, description, period, status);
                        ShowDataTableQuery(databasePerson, "Mailing", "SELECT SenderEmail AS 'Отправитель', RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', " +
                            "NameReport AS 'Наименование', Description AS 'Описание', Period AS 'Период', DateCreated AS 'Дата создания/модификации', Status AS 'Статус' ", " ORDER BY RecipientEmail asc, DateCreated desc; ");
                        break;
                    }
                case "sendEmail":
                    {
                        //Check making report
                        CheckAliveServer(sServer1, sServer1UserName, sServer1UserPassword);

                        if (bServer1Exist)
                        {
                            DataTable dtTempIntermediate = dtPeople.Clone();
                            Person personCheck = new Person();

                            DeleteAllDataInTableQuery(databasePerson, "PersonTemp");
                            DeleteAllDataInTableQuery(databasePerson, "PersonRegistered");
                            DeleteAllDataInTableQuery(databasePerson, "PersonRegisteredFull");

                            GetNamePoints();  //Get names of the points

                            string startDay = selectPeriodMonth().Split('|')[0];
                            string lastDay = selectPeriodMonth().Split('|')[1];

                            //periodComboParameters.Add("Текущая неделя");
                            //periodComboParameters.Add("Текущий месяц");
                            //periodComboParameters.Add("Предыдущая неделя");
                            //periodComboParameters.Add("Предыдущий месяц");
                            if (period.ToLower().Contains("текущий"))
                            {
                                startDay = selectPeriodMonth().Split('|')[0];
                                lastDay = selectPeriodMonth().Split('|')[1];
                            }

                            string[] nameGroups = groupsReport.Split('+');
                            string name = "";
                            string selectedPeriod = startDay.Split(' ')[0] + " - " + lastDay.Split(' ')[0];
                            string bodyOfMail = "";
                            string titleOfbodyMail = "";
                            foreach (string nameGroup in nameGroups)
                            {
                                try { System.IO.File.Delete(filePathExcelReport); } catch { }
                                name = nameGroup.Trim();
                                dtPersonRegisteredFull.Clear();
                                GetRegistrations(name, startDay, lastDay, "sendEmail");//typeReport== only one group

                                dtTempIntermediate = dtPeople.Clone();
                                personCheck = new Person();
                                dtPersonTemp?.Clear();

                                LoadGroupMembersFromDbToDataTable(dtPersonGroup, name); //result will be in dtPersonGroup  //"Select * FROM PersonGroup where GroupPerson like '" + _textBoxReturnText(textBoxGroup) + "';"

                                foreach (DataRow row in dtPersonGroup.Rows)
                                {
                                    if (row[1].ToString().Length > 0 && row[3].ToString() == name)
                                    {
                                        personCheck = new Person();

                                        personCheck.FIO = row[1].ToString();
                                        personCheck.NAV = row[2].ToString();
                                        personCheck.GroupPerson = name;
                                        personCheck.Department = name;

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

                                dtPersonTemp = GetDistinctRecords(dtTempIntermediate, orderColumnsFinacialReport);
                                dtPersonTemp.SetColumnsOrder(orderColumnsFinacialReport);

                                if (dtPersonTemp.Rows.Count > 0)
                                {
                                    filePathExcelReport = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(filePathApplication), nameReport + "|" + startDay.Split(' ')[0] + "-" + lastDay.Split(' ')[0] + " " + name + " " + DateTime.Now.ToString("yy-MM-dd HH-mm") + @".xlsx");
                                    try { System.IO.File.Delete(filePathExcelReport); } catch (Exception expt)
                                    {
                                        logger.Error("Ошибка удаления файла " + filePathExcelReport + " " + expt.ToString());
                                    }
                                    ExportDatatableSelectedColumnsToExcel(dtPersonTemp, nameReport, filePathExcelReport);

                                    bodyOfMail = "Отчет \"" + nameReport + "\" " + " по группе \"" + name + "\"" + "\nВыполнен " + DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                                    titleOfbodyMail = "Отчет за период: с " + startDay.Split(' ')[0] + " по " + lastDay.Split(' ')[0];

                                    SendEmailAsync(senderEmail, recipientEmail, titleOfbodyMail, bodyOfMail, filePathExcelReport, Properties.Resources.LogoRYIK);
                                    _toolStripStatusLabelBackColor(StatusLabel2, Color.PaleGreen);
                                    _toolStripStatusLabelSetText(StatusLabel2, DateTime.Now.ToString("HH:mm") + " Отчет " + nameReport + "(" + name + ") подготовлен и отправлен " + recipientEmail);
                                }
                                else
                                {
                                    _toolStripStatusLabelSetText(StatusLabel2, "Ошибка получения данных по отчету " + nameReport);
                                    _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                                }
                                //destroy temporary variables
                                personCheck = null;
                                dtPersonTemp?.Dispose();
                                dtTempIntermediate?.Dispose();
                            }
                            //destroy temporary variables
                            personCheck = null; dtTempIntermediate.Dispose();
                            startDay = null; lastDay = null; selectedPeriod = null;
                            bodyOfMail = null; titleOfbodyMail = null; nameGroups = null; name = null;
                        }
                        break;
                    }
                default:
                    break;
            }
        }

        private string selectPeriodMonth(bool current = false) //firstDay + "|" + lastDay
        {
            string result = "";
            var today = DateTime.Today;
            var lastDayPrevMonth = new DateTime(today.Year, today.Month, 1).AddDays(-1);
            if (current)
            { lastDayPrevMonth = new DateTime(today.Year, today.Month, today.Day); }

            result =
                lastDayPrevMonth.ToString("yyyy-MM") + "-01" + " 00:00:00" +
                "|" +
                lastDayPrevMonth.ToString("yyyy-MM-dd") + " 23:59:59";

            return result;
        }

        private async void SendEmailAsync(string sender, string recipient, string title, string message, string pathToFile, Bitmap myLogo) //Compose and send e-mail
        {
            // string startupPath = AppDomain.CurrentDomain.RelativeSearchPath;
            // string path = System.IO.Path.Combine(startupPath, "HtmlTemplates", "NotifyTemplate.html");
            // string body = System.IO.File.ReadAllText(path);

            // адрес smtp-сервера и порт, с которого будем отправлять письмо
            using (System.Net.Mail.SmtpClient smtpClient = new System.Net.Mail.SmtpClient(mailServer, 587))
            {
                smtpClient.EnableSsl = true;
                smtpClient.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                smtpClient.UseDefaultCredentials = false;
                smtpClient.Timeout = 50000;
                // создаем объект сообщения
                using (System.Net.Mail.MailMessage newMail = new System.Net.Mail.MailMessage())
                {
                    // письмо представляет код html
                    newMail.IsBodyHtml = true;
                    newMail.BodyEncoding = UTF8Encoding.UTF8;

                    newMail.AlternateViews.Add(getEmbeddedImage(title, message, myLogo));
                    // отправитель - устанавливаем адрес и отображаемое в письме имя
                    newMail.From = new System.Net.Mail.MailAddress(sender, productName);
                    // кому отправляем
                    newMail.To.Add(new System.Net.Mail.MailAddress(recipient));
                    // тема письма
                    newMail.Subject = title;
                    //добавляем вложение
                    newMail.Attachments.Add(new System.Net.Mail.Attachment(pathToFile));
                    // логин и пароль
                    smtpClient.Credentials = new System.Net.NetworkCredential(sender.Split('@')[0], "");
                    //if error - show error
                    newMail.DeliveryNotificationOptions = System.Net.Mail.DeliveryNotificationOptions.OnFailure;
                    //async send email
                    await smtpClient.SendMailAsync(newMail);
                }
            }
        }

        private System.Net.Mail.AlternateView getEmbeddedImage(string title, string message, Bitmap bmp)
        {
            //convert embedded resources into memorystream
            Bitmap b = new Bitmap(bmp);
            ImageConverter ic = new ImageConverter();
            Byte[] ba = (Byte[])ic.ConvertTo(b, typeof(Byte[]));
            System.IO.MemoryStream logo = new System.IO.MemoryStream(ba);

            System.Net.Mail.LinkedResource res = new System.Net.Mail.LinkedResource(logo, "image/jpeg");
            res.ContentId = Guid.NewGuid().ToString();

            // текст письма
            string htmlBody = @"<p>" + title + @"</p>" + @"<p>" + message + @"</p>" + @"<img src='cid:" + res.ContentId + @"'/>";

            System.Net.Mail.AlternateView alternateView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(htmlBody, null, System.Net.Mime.MediaTypeNames.Text.Html);
            alternateView.LinkedResources.Add(res);
            return alternateView;
        }

        //---  End. Schedule Functions ---//



        //---  Start. Block Encryption-Decryption ---//  

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

        //---  End. Block Encryption-Decryption ---//  



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
        public string RealDayOfWeek = "";

        public bool Late = false;
        public bool Early = false;

        public string serverSKD = "";
        public string namePassPoint = "";
        public string directionPass = "";
    }

    public class DataGridViewSeekValuesInSelectedRow
    {
        public string[] values = new string[10];
        public bool correctData { get; set; }

        public void FindValueInCells(DataGridView DGV, params string[] columnsName)
        {
            int IndexCurrentRow = DGV.CurrentRow.Index > -1 ? DGV.CurrentRow.Index : -1;
            correctData = IndexCurrentRow > -1 ? true : false;

            if (correctData)
            {
                for (int i = 0; i < DGV.ColumnCount; i++)
                {
                    if (DGV.Columns[i].HeaderText == columnsName[0])
                    {
                        values[0] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 1 && DGV.Columns[i].HeaderText == columnsName[1])
                    {
                        values[1] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 2 && DGV.Columns[i].HeaderText == columnsName[2])
                    {
                        values[2] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 3 && DGV.Columns[i].HeaderText == columnsName[3])
                    {
                        values[3] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 4 && DGV.Columns[i].HeaderText == columnsName[4])
                    {
                        values[4] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 5 && DGV.Columns[i].HeaderText == columnsName[5])
                    {
                        values[5] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 6 && DGV.Columns[i].HeaderText == columnsName[6])
                    {
                        values[6] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 7 && DGV.Columns[i].HeaderText == columnsName[7])
                    {
                        values[7] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 8 && DGV.Columns[i].HeaderText == columnsName[8])
                    {
                        values[8] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                    else if (columnsName.Length > 9 && DGV.Columns[i].HeaderText == columnsName[9])
                    {
                        values[9] = DGV.Rows[IndexCurrentRow].Cells[i].Value.ToString();
                    }
                }
            }

        }
    }

    public static class DataTableExtensions
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

    public class MailingStructure
    {
        public string _sender = "";
        public string _recipient = "";
        public string _groupsReport = "";
        public string _nameReport = "";
        public string _descriptionReport = "";
        public string _period = "";
        public string _status = "";
    }

    public class OutReasons
    {
        public int _id = 0;
        public string _name = "";
        public string _visibleName = "";
        public int _hourly = 0;
    }

    public class OutPerson
    {
        public int _reason_id = 0;
        public string _reason_Name = "";
        public string _nav = "";
        public string _date = "";
        public int _from = 0;
        public int _to = 0;
        public int _hourly = 0;
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
