using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
// for Crypography
using System.Security.Cryptography;

//using NLog;
//Project\Control NuGet\console 
//install-package nlog
//install-package nlog.config


namespace ASTA
{
    public partial class WinFormASTA :Form
    {
        private NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        private readonly string myRegKey = @"SOFTWARE\RYIK\ASTA";
        private readonly System.IO.FileInfo databasePerson = new System.IO.FileInfo(@".\main.db");
        private readonly byte[] btsMess1 = Convert.FromBase64String(@"OCvesunvXXsxtt381jr7vp3+UCwDbE4ebdiL1uinGi0="); //Key Encrypt
        private readonly byte[] btsMess2 = Convert.FromBase64String(@"NO6GC6Zjl934Eh8MAJWuKQ=="); //Key Decrypt
        private string nameOfLastTableFromDB = "PersonRegistrationsList";

        private string currentAction = "";
        private bool currentModeAppManual = true;
        private System.Diagnostics.FileVersionInfo myFileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath);
        private string guid = System.Runtime.InteropServices.Marshal.GetTypeLibGuidForAssembly(System.Reflection.Assembly.GetExecutingAssembly()).ToString(); // получаем GIUD приложения
        private string pcname = Environment.MachineName + "|" + Environment.OSVersion;
        private string poname = "";
        private string poversion = "";
        private string productName = "";
        private string strVersion;

        private ContextMenu contextMenu1;
        private bool buttonAboutForm;

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

        //Shows visual of registration
        private PictureBox pictureBox1 = new PictureBox();
        private Bitmap bmp = new Bitmap(1, 1);

        private List<string> selectedDates = new List<string>();
        private string[] myBoldedDates;
        private List<string> boldeddDates = new List<string>();
        private string[] workSelectedDays;
        private static string reportStartDay = "";
        private static string reportLastDay = "";
        private bool reportExcelReady = false;

        //Page of "Settings' Programm"
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

        //Page of Mailing
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

        private Label listComboLabel;
        private ComboBox listCombo = new ComboBox();

        private Label periodComboLabel;
        private ComboBox periodCombo = new ComboBox();

        private Label labelSettings9;
        private ComboBox comboSettings9 = new ComboBox();

        private Color clrRealRegistration = Color.PaleGreen;
        private string sLastSelectedElement = "MainForm";

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
                                  new DataColumn(@"Учетное время прихода ЧЧ:ММ",typeof(string)),//22
                                  new DataColumn(@"Учетное время ухода ЧЧ:ММ",typeof(string)),//23
                                  new DataColumn(@"Фактич. время прихода ЧЧ:ММ",typeof(string)),//24
                                  new DataColumn(@"Фактич. время ухода ЧЧ:ММ",typeof(string)), //25
                                  new DataColumn(@"Реальное отработанное время",typeof(decimal)), //26
                                  new DataColumn(@"Отработанное время ЧЧ:ММ",typeof(string)), //27
                                  new DataColumn(@"Опоздание ЧЧ:ММ",typeof(string)),                    //28
                                  new DataColumn(@"Ранний уход ЧЧ:ММ",typeof(string)),                 //29
                                  new DataColumn(@"Отпуск",typeof(string)),                 //30
                                  //new DataColumn(@"Отпуск (отгул)",typeof(string)),                 //30
                                  new DataColumn(@"Командировка",typeof(string)),                 //31
                                  new DataColumn(@"День недели",typeof(string)),                 //32
                                  new DataColumn(@"Больничный",typeof(string)),                 //33
                                  new DataColumn(@"Отсутствовал на работе",typeof(string)),     //34
                                  new DataColumn(@"Код",typeof(string)),                        //35
                                  new DataColumn(@"Вышестоящая группа",typeof(string)),         //36
                                  new DataColumn(@"Описание группы",typeof(string)),            //37
                                  new DataColumn(@"Комментарии",typeof(string)),                 //38
                                  new DataColumn(@"Должность",typeof(string)),                 //39
                                  new DataColumn(@"График",typeof(string)),                 //40
                                  new DataColumn(@"Отгул",typeof(string))                 //41
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
                                  @"Учетное время прихода ЧЧ:ММ",//22
                                  @"Учетное время ухода ЧЧ:ММ",//23
                                  @"Фактич. время прихода ЧЧ:ММ",//24
                                  @"Фактич. время ухода ЧЧ:ММ", //25
                                  @"Реальное отработанное время", //26
                                  @"Отработанное время ЧЧ:ММ", //27
                                  @"Опоздание ЧЧ:ММ",                    //28
                                  @"Ранний уход ЧЧ:ММ",                 //29
                                  @"Отпуск",                 //30
                                  @"Командировка",                 //31
                                  @"День недели",                    //32
                                  @"Больничный",                    //33
                                  @"Отсутствовал на работе",      //34
                                  @"Код",                           //35
                                  @"Вышестоящая группа",            //36
                                  @"Описание группы",                //37
                                  @"Комментарии",                    //38
                                  @"Должность",                    //39
                                  @"График",                    //40
                                  @"Отгул"                      //41
        };
        public readonly string[] orderColumnsFinacialReport =
            {
                                  @"Фамилия Имя Отчество",//1
                                  @"NAV-код",//2
                                  @"Отдел",//11
                                  @"Дата регистрации",//12
                                  @"День недели",                    //32
                                  @"Учетное время прихода ЧЧ:ММ",//22
                                  @"Учетное время ухода ЧЧ:ММ",//23
                                  @"Фактич. время прихода ЧЧ:ММ",//24
                                  @"Фактич. время ухода ЧЧ:ММ", //25
                                  @"Отработанное время ЧЧ:ММ", //27
                                  @"Опоздание ЧЧ:ММ",                    //28
                                  @"Ранний уход ЧЧ:ММ",                 //29
                                  @"Отпуск",                 //30
                                  @"Отгул",                      //41
                                  @"Командировка",                 //31
                                  @"Больничный",                    //33
                                  @"Комментарии",                    //38
                                  @"Отсутствовал на работе",      //34
                                  @"Должность",                    //39
                                  @"График"                    //40
        };
        public readonly string[] arrayHiddenColumnsFIO =
            {
                            @"№ п/п",//0
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
                            @"Фактич. время прихода ЧЧ:ММ",//24
                            @"Фактич. время ухода ЧЧ:ММ", //25
                            @"Реальное отработанное время", //26
                            @"Отработанное время ЧЧ:ММ", //27
                            @"Опоздание ЧЧ:ММ",                   //28
                            @"Ранний уход ЧЧ:ММ",              //29
                            @"Отпуск",           //30
                            @"Командировка",                 //31
                            @"День недели",                    //32
                            @"Больничный",                    //33
                            @"Отсутствовал на работе",      //34
                            @"Код",                           //35
                            @"Вышестоящая группа",            //36
                            @"Описание группы",                //37
                            @"Отгул"                      //41
        };
        public readonly string[] nameHidenColumnsArray =
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
                @"Сервер СКД", //19
                @"Имя точки прохода", //20
                @"Направление прохода", //21
                @"Реальное отработанное время", //26
                @"Отработанное время ЧЧ:ММ", //27
                @"Опоздание ЧЧ:ММ",                    //28
                @"Ранний уход ЧЧ:ММ",                 //29
                @"Отпуск",              //30
                @"Командировка",                 //31
                @"День недели",  //32
                @"Больничный",  //33
                @"Отсутствовал на работе",      //34
                @"Код",         //35
                @"Вышестоящая группа",            //36
                @"Описание группы",                //37
                @"Отгул",                      //41

        };
        private static List<OutReasons> outResons = new List<OutReasons>();

        private DataTable dtPersonTemp = new DataTable("PersonTemp");
        private DataTable dtPersonTempAllColumns = new DataTable("PersonTempAllColumns");
        private DataTable dtPersonRegistrationsFullList = new DataTable("PersonRegistrationsFullList");
        private DataTable dtPeopleGroup = new DataTable("PeopleGroup");
        private DataTable dtPeopleListLoaded = new DataTable("PeopleLoaded");


        //Color of Person's Control elements which depend on the selected MenuItem  
        private Color labelGroupCurrentBackColor;
        private Color textBoxGroupCurrentBackColor;
        private Color labelGroupDescriptionCurrentBackColor;
        private Color textBoxGroupDescriptionCurrentBackColor;
        private Color comboBoxFioCurrentBackColor;
        private Color textBoxFIOCurrentBackColor;
        private Color textBoxNavCurrentBackColor;


        public WinFormASTA()
        { InitializeComponent(); }

        private void Form1_Load(object sender, EventArgs e)
        {
            logger.Info("");
            logger.Info("-------------------------------");
            logger.Info("");
            logger.Info("");
            logger.Info("-= Загрузка ПО =-");
            logger.Info("");

            bool existed;
            // получаем GIUD приложения
            string guid = System.Runtime.InteropServices.Marshal.GetTypeLibGuidForAssembly(System.Reflection.Assembly.GetExecutingAssembly()).ToString();
            System.Threading.Mutex mutexObj = new System.Threading.Mutex(true, guid, out existed);

            if (existed)
            { Form1Load(); }
            else
            {
                logger.Warn("-=  Запущена вторая копия программы. Завершение работы.  =-");
                ApplicationExit(); //Bolck other this app to run
                System.Threading.Thread.Sleep(200);
                return;
            }
        }

        private void Form1Load()
        {
            logger.Info("");
            logger.Info("Test Info message");
            logger.Trace("Test1 Trace message");
            logger.Debug("Test2 Debug message");
            logger.Warn("Test3 Warn message");
            logger.Error("Test4 Error message");
            logger.Fatal("Test5 Fatal message");
            logger.Info("");

            currentModeAppManual = true;
            _MenuItemTextSet(modeItem, "Сменить на автоматический режим");
            _menuItemTooltipSet(modeItem, "Включен интерактивный режим. Все рассылки остановлены.");

            myFileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath);
            strVersion = myFileVersionInfo.Comments + " ver." + myFileVersionInfo.FileVersion + " " + myFileVersionInfo.LegalCopyright;
            productName = myFileVersionInfo.ProductName;

            StatusLabel1.Text = myFileVersionInfo.ProductName + " ver." + myFileVersionInfo.FileVersion + " " + myFileVersionInfo.LegalCopyright;
            StatusLabel1.Alignment = ToolStripItemAlignment.Right;
            StatusLabel2.Text = " Начните работу с кнопки - \"Получить ФИО\"";
            contextMenu1 = new ContextMenu();  //Context Menu on notify Icon
            notifyIcon1.ContextMenu = contextMenu1;
            contextMenu1.MenuItems.Add("About", AboutSoft);
            contextMenu1.MenuItems.Add("Exit", ApplicationExit);
            notifyIcon1.Text = myFileVersionInfo.ProductName + "\nv." + myFileVersionInfo.FileVersion + "\n" + myFileVersionInfo.CompanyName;
            this.Text = myFileVersionInfo.Comments;

            logger.Info("Настраиваю интерфейс....");
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
            TableModeItem.Visible = false;
            VisualModeItem.Visible = false;
            VisualSelectColorMenuItem.Visible = false;
            TableExportToExcelItem.Visible = false;
            listFioItem.Visible = false;
            dataGridView1.ShowCellToolTips = true;
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
            //DateTime.Now.ToString("yyyy-MM-dd HH:mm")

            numUpDownHourStart.Value = 9;
            numUpDownMinuteStart.Value = 0;
            numUpDownHourEnd.Value = 18;
            numUpDownMinuteEnd.Value = 0;

            PersonOrGroupItem.Text = "Работа с одной персоной";
            toolTip1.SetToolTip(textBoxGroup, "Создать группу");
            toolTip1.SetToolTip(textBoxGroupDescription, "Изменить описание группы");
            StatusLabel2.Text = "";

            logger.Info("Проверяю настройки на локальную БД...");
            TryMakeDB();
            UpdateTableOfDB();
            SetTechInfoIntoDB();
            BoldAnualDates();

            //read last saved parameters from db and Registry and set their into variables
            logger.Info("Загружаю настройки программы...");
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


            logger.Info("Настраиваю внутренние переменные....");

            //Prepare DataTables
            dtPeople.Columns.AddRange(dcPeople);
            dtPeople.DefaultView.Sort = "[Группа], [Фамилия Имя Отчество], [Дата регистрации], [Время регистрации], [Фактич. время прихода ЧЧ:ММ], [Фактич. время ухода ЧЧ:ММ] ASC";

            //Clone default column name and structure from 'dtPeople' to other DataTables
            dtPersonTemp = dtPeople.Clone();  //Copy only structure(Name of columns)
            dtPersonRegistrationsFullList = dtPeople.Clone();  //Copy only structure(Name of columns)
            dtPeopleGroup = dtPeople.Clone();  //Copy only structure(Name of columns)

            logger.Info(productName + " полностью загружена....");

            if (currentModeAppManual)
            {
                SeekAndShowMembersOfGroup("");
            }
            else
            {
                logger.Info(productName + " включен автоматический режим....");

                ShowDataTableQuery(databasePerson, "Mailing", "SELECT SenderEmail AS 'Отправитель', RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', " +
                "NameReport AS 'Наименование', Description AS 'Описание', Period AS 'Период', DateCreated AS 'Дата создания/модификации', SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус' ",
                " ORDER BY RecipientEmail asc, DateCreated desc; ");

                nameOfLastTableFromDB = "Mailing";
                ExecuteAutoMode(true);
            }
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
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'PeopleGroupDesciption' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, GroupPerson TEXT, GroupPersonDescription TEXT, Late TEXT, Early TEXT,  Vacancy TEXT, BusinesTrip TEXT, Reserv1 TEXT, Reserv2 TEXT, " +
                    "UNIQUE ('GroupPerson') ON CONFLICT REPLACE);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'PeopleGroup' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, FIO TEXT, NAV TEXT, GroupPerson TEXT, " +
                    "HourControlling TEXT, MinuteControlling TEXT, Controlling REAL, ControllingHHMM TEXT, HourControllingOut TEXT, MinuteControllingOut TEXT, ControllingOut REAL, ControllingOUTHHMM TEXT, " +
                    "Shift TEXT, Comment TEXT, Reserv1 TEXT, Reserv2 TEXT, UNIQUE ('FIO', 'NAV', 'GroupPerson') ON CONFLICT REPLACE);", databasePerson); //Reserv1=department //Reserv2=position
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'ListOfWorkTimeShifts' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, FIO TEXT, NAV TEXT, DayStartShift TEXT, " +
                    "MoStart REAL,MoEnd REAL, TuStart REAL,TuEnd REAL, WeStart REAL,WeEnd REAL, ThStart REAL,ThEnd REAL, FrStart REAL,FrEnd REAL, " +
                    "SaStart REAL,SaEnd REAL, SuStart REAL,SuEnd REAL, Status Text, Comment TEXT, DayInputed TEXT, Reserv1 TEXT, Reserv2 TEXT, UNIQUE ('NAV', 'DayStartShift') ON CONFLICT REPLACE);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'TechnicalInfo' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, PCName TEXT, POName TEXT, POVersion TEXT, LastDateStarted TEXT, Reserv1 TEXT, " +
                    "Reserv2 TEXT, Reserverd3 TEXT);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'BoldedDates' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, BoldedDate TEXT, NAV TEXT, Groups TEXT, Reserv1 TEXT, Reserv2 TEXT);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'ProgramSettings' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, PoParameterName TEXT, PoParameterValue TEXT, Reserv1 TEXT, Reserv2 TEXT, " +
                   "UNIQUE (PoParameterName) ON CONFLICT REPLACE);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'EquipmentSettings' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, EquipmentParameterName TEXT, EquipmentParameterValue TEXT, EquipmentParameterServer TEXT," +
                    "Reserv1, Reserv2, UNIQUE ('EquipmentParameterName', 'EquipmentParameterServer') ON CONFLICT REPLACE);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'LastTakenPeopleComboList' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, ComboList TEXT, " +
                    "Reserv1 TEXT, Reserv2 TEXT, UNIQUE ('ComboList', Reserv1) ON CONFLICT REPLACE);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'Mailing' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, SenderEmail TEXT, " +
                    "RecipientEmail TEXT, GroupsReport TEXT, NameReport TEXT, Description TEXT, DateCreated TEXT, Period TEXT, Status TEXT, SendingLastDate TEXT, Reserv1 TEXT, Reserv2 TEXT);", databasePerson);
        }

        private void UpdateTableOfDB()
        {
            TryUpdateStructureSqlDB("PeopleGroupDesciption", "GroupPerson TEXT, GroupPersonDescription TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
            TryUpdateStructureSqlDB("PeopleGroup", "FIO TEXT, NAV TEXT, GroupPerson TEXT, HourControlling TEXT, MinuteControlling TEXT, Controlling REAL, " +
                    "HourControllingOut TEXT, MinuteControllingOut TEXT, ControllingOut REAL, ControllingHHMM TEXT, ControllingOUTHHMM TEXT, " +
                    "Shift TEXT, Comment TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
            TryUpdateStructureSqlDB("ListOfWorkTimeShifts", "FIO TEXT, NAV TEXT, DayStartShift TEXT, MoStart REAL,MoEnd REAL, TuStart REAL,TuEnd REAL, WeStart REAL,WeEnd REAL, ThStart REAL,ThEnd REAL, FrStart REAL,FrEnd REAL, " +
                    "SaStart REAL,SaEnd REAL, SuStart REAL,SuEnd REAL, Status Text, Comment TEXT, DayInputed TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);

            TryUpdateStructureSqlDB("TechnicalInfo", "PCName TEXT, POName TEXT, POVersion TEXT, LastDateStarted TEXT, Reserv1 TEXT, Reserv2 TEXT, Reserverd3 TEXT", databasePerson);
            TryUpdateStructureSqlDB("BoldedDates", "BoldedDate TEXT, NAV TEXT, Groups TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
            TryUpdateStructureSqlDB("ProgramSettings", " PoParameterName TEXT, PoParameterValue TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
            TryUpdateStructureSqlDB("EquipmentSettings", "EquipmentParameterName TEXT, EquipmentParameterValue TEXT, EquipmentParameterServer TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
            TryUpdateStructureSqlDB("LastTakenPeopleComboList", "ComboList TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
            TryUpdateStructureSqlDB("Mailing", "SenderEmail TEXT, RecipientEmail TEXT, GroupsReport TEXT, NameReport TEXT, Description TEXT, DateCreated TEXT, Period TEXT, Status TEXT, SendingLastDate TEXT,Reserv1 TEXT, Reserv2 TEXT", databasePerson);

        }

        private void SetTechInfoIntoDB() //Write into DB Technical Info
        {
            guid = System.Runtime.InteropServices.Marshal.GetTypeLibGuidForAssembly(System.Reflection.Assembly.GetExecutingAssembly()).ToString(); // получаем GIUD приложения
            pcname = Environment.MachineName + "(" + Environment.OSVersion + ")";
            poname = myFileVersionInfo.FileName + "(" + myFileVersionInfo.ProductName + ")";
            poversion = myFileVersionInfo.FileVersion;
            string LastDateStarted = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            string Reserv1 = "Current user: " + Environment.UserName;
            string Reserv2 = "RAM: " + Environment.WorkingSet.ToString();

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
            LastDateStarted = null; Reserv1 = null; Reserv2 = null;
        }

        private void LoadPrevioslySavedParameters()   //Select Previous Data from DB and write it into the combobox and Parameters
        {
            string modeApp = "";
            int iCombo = 0;
            int numberOfFio = 0;
            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();

                    using (var sqlCommand = new SQLiteCommand("SELECT ComboList FROM LastTakenPeopleComboList;", sqlConnection))
                    {
                        using (var reader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in reader)
                            {
                                try
                                {
                                    if (record["ComboList"]?.ToString()?.Length > 0)
                                    {
                                        _comboBoxAdd(comboBoxFio, record["ComboList"].ToString().Trim());
                                        iCombo++;
                                    }
                                } catch (Exception expt) { logger.Info(expt.ToString()); }
                            }
                        }
                    }

                    using (var sqlCommand = new SQLiteCommand("SELECT FIO FROM PeopleGroup;", sqlConnection))
                    {
                        using (var reader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in reader)
                            {
                                try
                                {
                                    if (record["FIO"]?.ToString()?.Length > 0)
                                    { numberOfFio++; }
                                } catch { logger.Info("SELECT FIO FROM PeopleGroup - error"); }
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
                                    if (record["PoParameterName"]?.ToString()?.Length > 0)
                                    {
                                        if (record["PoParameterName"]?.ToString()?.Trim() == "clrRealRegistration")
                                            clrRealRegistration = Color.FromName(record["PoParameterValue"].ToString());
                                    }
                                } catch (Exception expt) { logger.Info(expt.ToString()); }
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
                                    if (record["EquipmentParameterValue"]?.ToString()?.Length > 0)
                                    {
                                        if (record["EquipmentParameterValue"]?.ToString()?.Trim() == "SKDServer" && record["EquipmentParameterName"]?.ToString()?.Trim() == "SKDUser")
                                        {
                                            sServer1DB = record["EquipmentParameterServer"]?.ToString();

                                            //todo - check to need try
                                            if (record["Reserv1"]?.ToString()?.Length > 0)
                                                try { sServer1UserNameDB = DecryptBase64ToString(record["Reserv1"].ToString(), btsMess1, btsMess2); } catch { }
                                            //todo - check to need try
                                            if (record["Reserv2"]?.ToString()?.Length > 0)
                                                try { sServer1UserPasswordDB = DecryptBase64ToString(record["Reserv2"].ToString(), btsMess1, btsMess2); } catch { }
                                        }
                                        else if (record["EquipmentParameterValue"]?.ToString()?.Trim() == "MailServer" && record["EquipmentParameterName"]?.ToString()?.Trim() == "MailUser")
                                        {
                                            mailServerDB = record["EquipmentParameterServer"]?.ToString();
                                            mailServerUserNameDB = record["Reserv1"]?.ToString();

                                            //todo - check to need try
                                            if (record["Reserv2"]?.ToString()?.Length > 0)
                                                try { mailServerUserPasswordDB = DecryptBase64ToString(record["Reserv2"].ToString(), btsMess1, btsMess2); } catch { }
                                        }
                                        else if (record["EquipmentParameterValue"]?.ToString()?.Trim() == "MySQLServer" && record["EquipmentParameterName"]?.ToString()?.Trim() == "MySQLUser")
                                        {
                                            mysqlServerDB = record["EquipmentParameterServer"]?.ToString();
                                            mysqlServerUserNameDB = record["Reserv1"]?.ToString();

                                            //todo - check to need try
                                            if (record["Reserv2"]?.ToString()?.Length > 0)
                                                try { mysqlServerUserPasswordDB = DecryptBase64ToString(record["Reserv2"].ToString(), btsMess1, btsMess2); } catch { }
                                        }
                                    }
                                } catch (Exception expt) { logger.Info(expt.ToString()); }
                            }
                        }
                    }
                }
            }
            iFIO = iCombo;
            try { comboBoxFio.SelectedIndex = 0; } catch { }

            using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(myRegKey, Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree, System.Security.AccessControl.RegistryRights.ReadKey))
            {
                try { sServer1Registry = EvUserKey.GetValue("SKDServer").ToString().Trim(); } catch { logger.Warn("Registry GetValue SKDServer"); }
                try { sServer1UserNameRegistry = DecryptBase64ToString(EvUserKey.GetValue("SKDUser").ToString(), btsMess1, btsMess2).ToString().Trim(); } catch { }
                try { sServer1UserPasswordRegistry = DecryptBase64ToString(EvUserKey.GetValue("SKDUserPassword").ToString(), btsMess1, btsMess2).ToString().Trim(); } catch { }

                try { mailServerRegistry = EvUserKey.GetValue("MailServer").ToString().Trim(); } catch { logger.Warn("Registry GetValue Mail"); }
                try { mailServerUserNameRegistry = EvUserKey.GetValue("MailUser").ToString().Trim(); } catch { }
                try { mailServerUserPasswordRegistry = DecryptBase64ToString(EvUserKey.GetValue("MailUserPassword").ToString(), btsMess1, btsMess2).ToString().Trim(); } catch { }

                try { mysqlServerRegistry = EvUserKey.GetValue("MySQLServer").ToString().Trim(); } catch { logger.Warn("Registry GetValue MySQL"); }
                try { mysqlServerUserNameRegistry = EvUserKey.GetValue("MySQLUser").ToString().Trim(); } catch { }
                try { mysqlServerUserPasswordRegistry = DecryptBase64ToString(EvUserKey.GetValue("MySQLUserPassword").ToString(), btsMess1, btsMess2).ToString().Trim(); } catch { }

                try { modeApp = EvUserKey.GetValue("ModeApp").ToString().Trim(); } catch { logger.Warn("Registry GetValue ModeApp"); }
            }
            switch (modeApp)
            {
                case "0":
                    currentModeAppManual = false;
                    break;
                case "1":
                    currentModeAppManual = true;
                    break;
                default:
                    currentModeAppManual = true;
                    break;
            }
            logger.Info("modeApp - " + modeApp + ", currentModeAppManual - " + currentModeAppManual);

            if (numberOfFio > 0)
            { _MenuItemVisible(listFioItem, true); }
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
        private void ShowDataTableQuery(System.IO.FileInfo databasePerson, string myTable, string mySqlQuery, string mySqlWhere) //Query data from the Table of the DB
        {
            DataTable dt = new DataTable();
            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    using (var sqlDA = new SQLiteDataAdapter(mySqlQuery + " FROM '" + myTable + "' " + mySqlWhere + "; ", sqlConnection))
                    {
                        dt = new DataTable();
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
                    {
                        try
                        {
                            sqlCommand.ExecuteNonQuery();
                            logger.Info("-= Удалена таблица " + myTable + " =-");
                        } catch { }
                    }
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
        {
            GetFIO();
        }

        private async void GetFIO()  // CheckAliveServer()   GetFioFromServers()  ImportTablePeopleToTableGroupsInLocalDB()
        {
            _ProgressBar1Start();
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
                dataGridView1.Visible = true;
                pictureBox1.Visible = false;

                await Task.Run(() => GetFioFromServers(dtTempIntermediate));

                dtPeopleListLoaded?.Clear();
                dtPeopleListLoaded = dtTempIntermediate.Copy();

                await Task.Run(() => ImportTablePeopleToTableGroupsInLocalDB(databasePerson.ToString(), dtTempIntermediate));

                //show selected data     
                //distinct Records                
                var namesDistinctColumnsArray = arrayAllColumnsDataTablePeople.Except(arrayHiddenColumnsFIO).ToArray(); //take distinct data
                dtPersonTemp = GetDistinctRecords(dtTempIntermediate, namesDistinctColumnsArray);

                await Task.Run(() => ShowDatatableOnDatagridview(dtPersonTemp, arrayHiddenColumnsFIO));

                await Task.Run(() => panelViewResize(numberPeopleInLoading));
                listFioItem.Visible = true;
                nameOfLastTableFromDB = "ListFIO";
            }
            else { GetInfoSetup(); }

            _MenuItemEnabled(SettingsMenuItem, true);
            _ProgressBar1Stop();
        }

        private void GetFioFromServers(DataTable dataTablePeople) //Get the list of registered users
        {
            Person personFromServer = new Person();
            dataTablePeople.Clear();
            iFIO = 0;
            DataRow row;
            string stringConnection;
            string query;
            string fio = "";
            string groupName = "";
            string depName = "";
            string depDescr = "";

            string timeStart = ConvertDecimalTimeToStringHHMM(_numUpDownReturn(numUpDownHourStart), _numUpDownReturn(numUpDownMinuteStart));
            string timeEnd = ConvertDecimalTimeToStringHHMM(_numUpDownReturn(numUpDownHourEnd), _numUpDownReturn(numUpDownMinuteEnd));
            string dayStartShift = "";

            string timeStart_ = ConvertDecimalTimeToStringHHMM(_numUpDownReturn(numUpDownHourStart), _numUpDownReturn(numUpDownMinuteStart));
            string timeEnd_ = ConvertDecimalTimeToStringHHMM(_numUpDownReturn(numUpDownHourEnd), _numUpDownReturn(numUpDownMinuteEnd));
            string dayStartShift_ = "";

            string nav = "";
            List<string> ListFIOTemp = new List<string>();
            listFIO = new List<string>();
            List<string> listCodesWithIdCard = new List<string>(); //NAV-codes people who have idCard

            // import users and group from www.ais
            //   List<string> listCodesFromWeb = new List<string>();
            List<PeopleShift> peopleShifts = new List<PeopleShift>();
            List<PeopleDepartment> departments = new List<PeopleDepartment>();
            List<PeopleDepartment> groups = new List<PeopleDepartment>();

            try
            {
                _comboBoxClr(comboBoxFio);
                _toolStripStatusLabelSetText(StatusLabel2, "Запрашиваю данные с " + sServer1 + ". Ждите окончания процесса...");
                stimerPrev = "Запрашиваю списки персонала с " + sServer1 + ". Ждите окончания процесса...";
                stringConnection = "Data Source=" + sServer1 + "\\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID=" + sServer1UserName + ";Password=" + sServer1UserPassword + "; Connect Timeout=60";
                logger.Info(stringConnection);

                // import users and group from SCA server
                using (var sqlConnection = new System.Data.SqlClient.SqlConnection(stringConnection))
                {
                    sqlConnection.Open();

                    //import group from SCA server
                    query = "SELECT id,level_id,name,owner_id,parent_id,region_id,schedule_id  FROM OBJ_DEPARTMENT";
                    logger.Info(query);
                    using (var sqlCommand = new System.Data.SqlClient.SqlCommand(query, sqlConnection))
                    {
                        using (var sqlReader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in sqlReader)
                            {
                                try
                                {
                                    if (record?["name"].ToString().Trim().Length > 0)
                                    {
                                        groups.Add(new PeopleDepartment()
                                        {
                                            _parent_id = record["parent_id"].ToString(),
                                            _departmentName = record["id"].ToString(),
                                            _departmentDescription = record["name"].ToString()
                                        });
                                    }
                                } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                                _ProgressWork1Step(1);
                            }
                        }
                    }

                    //import users from с SCA server
                    query = "SELECT id, name, surname, patronymic, post, tabnum, parent_id FROM OBJ_PERSON ";
                    logger.Info(query);
                    using (var sqlCommand = new System.Data.SqlClient.SqlCommand(query, sqlConnection))
                    {
                        using (var sqlReader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in sqlReader)
                            {
                                try
                                {
                                    string id = record?["id"].ToString();
                                    if (record?["name"].ToString().Trim().Length > 0)
                                    {
                                        iFIO++;
                                        row = dataTablePeople.NewRow();
                                        fio = "";
                                        //todo
                                        //simplify - record["name"]?.ToString()?.Trim()
                                        try { fio = record["name"].ToString().Trim(); } catch { }
                                        try { fio += " " + record["surname"].ToString().Trim(); fio = fio.Trim(); } catch { }
                                        try { fio += " " + record["patronymic"].ToString().Trim(); fio = fio.Trim(); } catch { }
                                        groupName = record["parent_id"].ToString().Trim();
                                        nav = record["tabnum"].ToString().ToUpper().Trim();
                                        try
                                        {
                                            depName = groups.Find((x) => x._departmentName == groupName)?._departmentDescription;
                                        } catch (Exception expt) { logger.Warn(expt.Message); }

                                        row[@"№ п/п"] = iFIO;
                                        row[@"№ пропуска"] = Convert.ToInt32(record["id"].ToString().Trim());
                                        row[@"Фамилия Имя Отчество"] = fio;
                                        row[@"NAV-код"] = nav;

                                        row[@"Группа"] = groupName;
                                        row[@"Отдел"] = depName;
                                        row[@"Должность"] = record["post"].ToString().Trim();

                                        row[@"Учетное время прихода ЧЧ:ММ"] = timeStart;
                                        row[@"Учетное время ухода ЧЧ:ММ"] = timeEnd;

                                        dataTablePeople.Rows.Add(row);

                                        listFIO.Add(record["name"].ToString().Trim() + "|" + record["surname"].ToString().Trim() + "|" + record["patronymic"].ToString().Trim() + "|" + record["id"].ToString().Trim() + "|" +
                                                    record["tabnum"].ToString().Trim() + "|" + sServer1);

                                        ListFIOTemp.Add(personFromServer.FIO + "|" + personFromServer.NAV);
                                        listCodesWithIdCard.Add(nav);

                                        _ProgressWork1Step(1);
                                    }
                                } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                            }
                        }
                    }
                }

                // import users and group from web DB
                int tmpSeconds = 0;
                _toolStripStatusLabelSetText(StatusLabel2, "Запрашиваю данные с " + mysqlServer + ". Ждите окончания процесса...");
                stimerPrev = "Запрашиваю данные с " + mysqlServer + ". Ждите окончания процесса...";

                groupName = "web"; //People imported from web DB
                groups.Add(new PeopleDepartment()
                {
                    _parent_id = mysqlServer,
                    _departmentName = groupName,
                    _departmentDescription = "Stuff from the web server"
                });

                stringConnection = @"server=" + mysqlServer + @";User=" + mysqlServerUserName + @";Password=" + mysqlServerUserPassword + @";database=wwwais;convert zero datetime=True;Connect Timeout=60"; //Allow Zero Datetime=true;
                logger.Info(stringConnection);
                using (var sqlConnection = new MySql.Data.MySqlClient.MySqlConnection(stringConnection))
                {
                    sqlConnection.Open();

                    // import departments from web DB
                    query = "SELECT id, parent_id, name, stat, boss_code, head_code FROM dep_struct ORDER by id";
                    logger.Info(query);
                    using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(query, sqlConnection))
                    {
                        using (MySql.Data.MySqlClient.MySqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                if (reader?.GetString(@"name") != null && reader?.GetString(@"name")?.Length > 0)
                                {
                                    departments.Add(new PeopleDepartment()
                                    {
                                        _id = reader?.GetString(@"id"),
                                        _departmentName = reader?.GetString(@"name")
                                    });
                                    _ProgressWork1Step(1);
                                }
                            }
                        }
                    }

                    // import individual shifts of people from web DB
                    query = "Select code,start_date,mo_start,mo_end,tu_start,tu_end,we_start,we_end,th_start,th_end,fr_start,fr_end, " +
                                    "sa_start,sa_end,su_start,su_end,comment FROM work_time ORDER by start_date";
                    logger.Info(query);
                    using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(query, sqlConnection))
                    {
                        using (MySql.Data.MySqlClient.MySqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                if (reader?.GetString(@"code") != null && reader?.GetString(@"code").Length > 0)
                                {
                                    try { dayStartShift = DateTime.Parse(reader?.GetMySqlDateTime(@"start_date").ToString()).ToString("yyyy-MM-dd"); } catch
                                    { dayStartShift = DateTime.Parse("1980-01-01").ToString("yyyy-MM-dd"); }

                                    peopleShifts.Add(new PeopleShift()
                                    {
                                        _nav = reader.GetString(@"code")?.Replace('C', 'S'),
                                        _dayStartShift = dayStartShift,
                                        _MoStart = Convert.ToInt32(reader.GetString(@"mo_start")),
                                        _MoEnd = Convert.ToInt32(reader.GetString(@"mo_end")),
                                        _TuStart = Convert.ToInt32(reader.GetString(@"tu_start")),
                                        _TuEnd = Convert.ToInt32(reader.GetString(@"tu_end")),
                                        _WeStart = Convert.ToInt32(reader.GetString(@"we_start")),
                                        _WeEnd = Convert.ToInt32(reader.GetString(@"we_end")),
                                        _ThStart = Convert.ToInt32(reader.GetString(@"th_start")),
                                        _ThEnd = Convert.ToInt32(reader.GetString(@"th_end")),
                                        _FrStart = Convert.ToInt32(reader.GetString(@"fr_start")),
                                        _FrEnd = Convert.ToInt32(reader.GetString(@"fr_end")),
                                        _SaStart = Convert.ToInt32(reader.GetString(@"sa_start")),
                                        _SaEnd = Convert.ToInt32(reader.GetString(@"sa_end")),
                                        _SuStart = Convert.ToInt32(reader.GetString(@"su_start")),
                                        _SuEnd = Convert.ToInt32(reader.GetString(@"su_end")),
                                        _Status = "",
                                        _Comment = reader.GetString(@"comment")
                                    });
                                    _ProgressWork1Step(1);
                                }
                            }
                        }
                    }

                    try
                    {
                        dayStartShift_ = "Общий график с " + peopleShifts.FindLast((x) => x._nav == "0")._dayStartShift;

                        tmpSeconds = peopleShifts.FindLast((x) => x._nav == "0")._MoStart;
                        timeStart_ = ConvertSecondsTimeToStringHHMMArray(tmpSeconds)[2];

                        tmpSeconds = peopleShifts.FindLast((x) => x._nav == "0")._MoEnd;
                        timeEnd_ = ConvertSecondsTimeToStringHHMMArray(tmpSeconds)[2];
                    } catch { }

                    // import people from web DB
                    query = "Select code, family_name,first_name,last_name,vacancy,department FROM personal"; //where hidden=0
                    logger.Info(query);
                    using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(query, sqlConnection))
                    {
                        using (MySql.Data.MySqlClient.MySqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                if (reader.GetString(@"family_name") != null && reader.GetString(@"family_name").Length > 0)
                                {
                                    row = dataTablePeople.NewRow();
                                    iFIO++;

                                    fio = "";
                                    try { fio = reader.GetString(@"family_name").Trim(); } catch { }
                                    try { fio += " " + reader.GetString(@"first_name").Trim(); fio = fio.Trim(); } catch { }
                                    try { fio += " " + reader.GetString(@"last_name").Trim(); fio = fio.Trim(); } catch { }

                                    personFromServer = new Person();
                                    personFromServer.FIO = fio.Replace("&acute;", "'");
                                    nav = reader.GetString(@"code").Trim().ToUpper().Replace('C', 'S');
                                    personFromServer.NAV = nav;

                                    depName = departments.FindLast((x) => x._id == reader.GetString(@"department").Trim())?._departmentName;
                                    personFromServer.Department = depName ?? reader?.GetString(@"department")?.Trim();
                                    personFromServer.PositionInDepartment = reader.GetString(@"vacancy")?.Trim();
                                    personFromServer.GroupPerson = groupName;

                                    personFromServer.Shift = dayStartShift_;

                                    personFromServer.ControlInHHMM = timeStart_;
                                    personFromServer.ControlOutHHMM = timeEnd_;

                                    try
                                    {
                                        personFromServer.Shift = "Индивидуальный график с " + peopleShifts.FindLast((x) => x._nav == personFromServer.NAV)._dayStartShift;

                                        tmpSeconds = peopleShifts.FindLast((x) => x._nav == personFromServer.NAV)._MoStart;
                                        personFromServer.ControlInHHMM = ConvertSecondsTimeToStringHHMMArray(tmpSeconds)[2];

                                        tmpSeconds = peopleShifts.FindLast((x) => x._nav == personFromServer.NAV)._MoEnd;
                                        personFromServer.ControlOutHHMM = ConvertSecondsTimeToStringHHMMArray(tmpSeconds)[2];
                                    } catch { }

                                    row[@"№ п/п"] = iFIO;
                                    row[@"Фамилия Имя Отчество"] = personFromServer.FIO;
                                    row[@"NAV-код"] = personFromServer.NAV;

                                    row[@"Группа"] = personFromServer.GroupPerson;

                                    row[@"Отдел"] = personFromServer.Department;
                                    row[@"Должность"] = personFromServer.PositionInDepartment;

                                    row[@"График"] = personFromServer.Shift;

                                    row[@"Учетное время прихода ЧЧ:ММ"] = personFromServer.ControlInHHMM;
                                    row[@"Учетное время ухода ЧЧ:ММ"] = personFromServer.ControlOutHHMM;

                                    ListFIOTemp.Add(personFromServer.FIO + "|" + personFromServer.NAV);

                                    if (listCodesWithIdCard.IndexOf(nav) != -1)
                                    {
                                        // listCodesFromWeb.Add(nav);
                                        dataTablePeople.Rows.Add(row);
                                    }

                                    _ProgressWork1Step(1);
                                }
                            }
                        }
                    }
                    sqlConnection.Close();
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
                DeleteAllDataInTableQuery(databasePerson, "LastTakenPeopleComboList");

                for (int indDep = 0; indDep < groups.Count; indDep++)
                {
                    try
                    {
                        depName = groups[indDep]?._departmentName;
                        DeleteDataTableQueryParameters(databasePerson, "PeopleGroup", "GroupPerson", depName, "", "", "", "");
                    } catch { }
                }

                for (int indDep = 0; indDep < groups.Count; indDep++)
                {
                    try
                    {
                        depName = groups[indDep]?._departmentName;
                        depDescr = groups[indDep]?._departmentDescription;
                        CreateGroupInDB(databasePerson, depName, depDescr);
                    } catch { }
                }

                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();

                    var sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                    foreach (var shift in peopleShifts.ToArray())
                    {
                        if (shift._nav != null && shift._nav.Length > 0)
                        {
                            using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'ListOfWorkTimeShifts' (NAV, DayStartShift, MoStart, MoEnd, " +
                                "TuStart, TuEnd, WeStart, WeEnd, ThStart, ThEnd, FrStart, FrEnd, SaStart, SaEnd, SuStart, SuEnd, Status, Comment, DayInputed) " +
                            " VALUES (@NAV, @DayStartShift, @MoStart, @MoEnd, @TuStart, @TuEnd, @WeStart, @WeEnd, @ThStart, @ThEnd, @FrStart, @FrEnd, " +
                            " @SaStart, @SaEnd, @SuStart, @SuEnd, @Status, @Comment, @DayInputed)", sqlConnection))
                            {
                                sqlCommand.Parameters.Add("@NAV", DbType.String).Value = shift._nav;
                                sqlCommand.Parameters.Add("@DayStartShift", DbType.String).Value = shift._dayStartShift;
                                sqlCommand.Parameters.Add("@MoStart", DbType.Int32).Value = shift._MoStart;
                                sqlCommand.Parameters.Add("@MoEnd", DbType.Int32).Value = shift._MoEnd;
                                sqlCommand.Parameters.Add("@TuStart", DbType.Int32).Value = shift._TuStart;
                                sqlCommand.Parameters.Add("@TuEnd", DbType.Int32).Value = shift._TuEnd;
                                sqlCommand.Parameters.Add("@WeStart", DbType.Int32).Value = shift._WeStart;
                                sqlCommand.Parameters.Add("@WeEnd", DbType.Int32).Value = shift._WeEnd;
                                sqlCommand.Parameters.Add("@ThStart", DbType.Int32).Value = shift._ThStart;
                                sqlCommand.Parameters.Add("@ThEnd", DbType.Int32).Value = shift._ThEnd;
                                sqlCommand.Parameters.Add("@FrStart", DbType.Int32).Value = shift._FrStart;
                                sqlCommand.Parameters.Add("@FrEnd", DbType.Int32).Value = shift._FrEnd;
                                sqlCommand.Parameters.Add("@SaStart", DbType.Int32).Value = shift._SaStart;
                                sqlCommand.Parameters.Add("@SaEnd", DbType.Int32).Value = shift._SaEnd;
                                sqlCommand.Parameters.Add("@SuStart", DbType.Int32).Value = shift._SuStart;
                                sqlCommand.Parameters.Add("@SuEnd", DbType.Int32).Value = shift._SuEnd;
                                sqlCommand.Parameters.Add("@Status", DbType.String).Value = shift._Status;
                                sqlCommand.Parameters.Add("@Comment", DbType.String).Value = shift._Comment;
                                sqlCommand.Parameters.Add("@DayInputed", DbType.String).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                                try { sqlCommand.ExecuteNonQuery(); } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                            }
                        }
                    }
                    sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();

                    sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                    foreach (string str in ListFIOCombo.ToArray())
                    {
                        if (str != null && str.Length > 1)
                        {
                            using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'LastTakenPeopleComboList' (ComboList, Reserv1, Reserv2) " +
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

                    sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                    foreach (var dr in dataTablePeople.AsEnumerable())
                    {
                        if (dr[@"Фамилия Имя Отчество"] != null && dr[@"NAV-код"] != null && dr[@"NAV-код"].ToString().Length > 0)
                        {
                            using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'PeopleGroup' (FIO, NAV, GroupPerson, ControllingHHMM, ControllingOUTHHMM, Shift, Reserv1, Reserv2) " +
                                    " VALUES (@FIO, @NAV, @GroupPerson, @ControllingHHMM, @ControllingOUTHHMM, @Shift, @Reserv1, @Reserv2)", sqlConnection))
                            {
                                sqlCommand.Parameters.Add("@FIO", DbType.String).Value = dr[@"Фамилия Имя Отчество"].ToString();
                                sqlCommand.Parameters.Add("@NAV", DbType.String).Value = dr[@"NAV-код"].ToString();
                                sqlCommand.Parameters.Add("@GroupPerson", DbType.String).Value = dr[@"Группа"].ToString();
                                sqlCommand.Parameters.Add("@ControllingHHMM", DbType.String).Value = dr[@"Учетное время прихода ЧЧ:ММ"].ToString();
                                sqlCommand.Parameters.Add("@ControllingOUTHHMM", DbType.String).Value = dr[@"Учетное время ухода ЧЧ:ММ"].ToString();
                                sqlCommand.Parameters.Add("@Shift", DbType.String).Value = dr[@"График"].ToString();
                                // sqlCommand.Parameters.Add("@Comment", DbType.String).Value = dr[@"Комментарии"].ToString();
                                sqlCommand.Parameters.Add("@Reserv1", DbType.String).Value = dr[@"Отдел"].ToString();
                                sqlCommand.Parameters.Add("@Reserv2", DbType.String).Value = dr[@"Должность"].ToString();
                                try
                                {
                                    sqlCommand.ExecuteNonQuery();
                                } catch (Exception expt)
                                {
                                    logger.Warn("ошибка записи в локальную базу PeopleGroup - " + dr[@"Фамилия Имя Отчество"] + "\n" + dr[@"NAV-код"] + "\n" + expt.ToString());
                                }
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
            listCodesWithIdCard = null;

            if (!bServer1Exist)
            {
                _toolStripStatusLabelSetText(StatusLabel2, "Ошибка доступа к БД СКД-сервера");
                _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                _MenuItemEnabled(QuickLoadDataItem, false);
                MessageBox.Show("Проверьте правильность написания серверов,\nимя и пароль sa-администратора,\nа а также доступность серверов и их баз!");
            }
            _MenuItemEnabled(GetFioItem, true);
            _MenuItemEnabled(FunctionMenuItem, true);
            _MenuItemEnabled(SettingsMenuItem, true);
            _MenuItemEnabled(AnualDatesMenuItem, true);
            _MenuItemEnabled(GroupsMenuItem, true);
        }

        private void listFioItem_Click(object sender, EventArgs e) //ListFioReturn()
        {
            SeekAndShowMembersOfGroup("");
            nameOfLastTableFromDB = "ListFIO";
        }

        /*
        private void ExportDatagridToExcel()  //Export to Excel from DataGridView
        {
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

            stimerPrev = "";
            _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);
            sLastSelectedElement = "ExportExcel";
            iDGColumns = 0; iDGRows = 0;
            _toolStripStatusLabelSetText(StatusLabel2, "Готово!");
        }
        */

        private string filePathApplication = Application.ExecutablePath;
        private string filePathExcelReport;

        private void ExportDatatableSelectedColumnsToExcel(DataTable dataTable, string nameReport, string filePath)  //Export DataTable to Excel 
        {
            reportExcelReady = false;
            dataTable.SetColumnsOrder(orderColumnsFinacialReport);
            DataView viewExport = new DataView(dataTable);
            viewExport.Sort = "[Фамилия Имя Отчество], [Дата регистрации] ASC";
            DataTable dtExport = viewExport.ToTable();

            _toolStripStatusLabelSetText(StatusLabel2, "В таблице " + dataTable.TableName + " столбцов всего - " + dtExport.Columns.Count + ", строк - " + dtExport.Rows.Count);
            _toolStripStatusLabelSetText(StatusLabel2, "Записываю в Excel-файл отчет - " + nameReport);

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

                //colourize of columns
                sheet.Columns[columnsInTable].Interior.Color = Color.Silver;  // последняя колонка

                try
                {
                    sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(@"Фамилия Имя Отчество")) + 1)]
                    .Interior.Color = Color.DarkSeaGreen;
                } catch { } //"Фамилия Имя Отчество"

                try
                {
                    sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(@"Опоздание ЧЧ:ММ")) + 1)]
                    .Interior.Color = Color.SandyBrown;
                } catch { } //"Опоздание ЧЧ:ММ"

                try
                {
                    sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(@"Ранний уход ЧЧ:ММ")) + 1)]
                    .Interior.Color = Color.SandyBrown;
                } catch { } //"Ранний уход ЧЧ:ММ"

                _ProgressWork1Step(1);

                for (int column = 0; column < columnsInTable; column++)
                {
                    sheet.Cells[column + 1].WrapText = true;
                    sheet.Cells[1, column + 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    sheet.Cells[1, column + 1].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
                    sheet.Cells[1, column + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    sheet.Cells[1, column + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    sheet.Cells[1, column + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    sheet.Cells[1, column + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    sheet.Cells[1, column + 1].Value = nameColumns[column];
                    sheet.Cells[1, column + 1].Interior.Color = Color.Silver; //colourize the first row
                    sheet.Columns[column + 1].Font.Size = 8;
                    sheet.Columns[column + 1].Font.Name = "Tahoma";
                    _ProgressWork1Step(1);
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
                        _ProgressWork1Step(1);
                    }
                }

                //Область сортировки          
                Microsoft.Office.Interop.Excel.Range range = sheet.Range["A2", GetExcelColumnName(columnsInTable) + (rows - 1)];

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

                _ProgressWork1Step(1);

                releaseObject(range);
                //  releaseObject(rangeKey);
                releaseObject(sheet);
                releaseObject(sheets);
                releaseObject(workbook);
                releaseObject(workbooks);
                excel.Quit();
                releaseObject(excel);
                indexColumns = null;
                nameColumns = null;
                //
                _ProgressWork1Step(1);

                stimerPrev = "";
                _toolStripStatusLabelSetText(StatusLabel2, "Путь к отчету: " + filePath);
                _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);
                logger.Info("Отчет сгенерирован " + nameReport + " и сохранен в " + filePath);
                reportExcelReady = true;
            } catch (Exception expt)
            {
                _toolStripStatusLabelSetText(StatusLabel2, "Ошибка генерации файла. Проверьте наличие установленного Excel");
                logger.Error("ExportDatatableSelectedColumnsToExcel - " + expt.ToString());
            }
            viewExport?.Dispose();
            dtExport?.Dispose();

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

            // _toolStripStatusLabelSetText(StatusLabel2, "Генерирую Excel-файл");
            // stimerPrev = "Наполняю файл данными из текущей таблицы";


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

            if (nameOfLastTableFromDB == "PeopleGroup")
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
            nameOfLastTableFromDB = "PersonRegistrationsList";
        }

        private void CreateGroupItem_Click(object sender, EventArgs e)
        {
            if (textBoxGroup.Text.Trim().Length > 0)
            {
                CreateGroupInDB(databasePerson, textBoxGroup.Text.Trim(), textBoxGroupDescription.Text.Trim());
            }

            PersonOrGroupItem.Text = "Работа с одной персоной";
            nameOfLastTableFromDB = "PeopleGroup";
            ListGroups();
        }

        private void CreateGroupInDB(System.IO.FileInfo fileInfo, string nameGroup, string descriptionGroup)
        {
            if (nameGroup.Length > 0)
            {
                using (var connection = new SQLiteConnection($"Data Source={fileInfo};Version=3;"))
                {
                    connection.Open();
                    using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PeopleGroupDesciption' (GroupPerson, GroupPersonDescription) " +
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

            ShowDataTableQuery(databasePerson, "PeopleGroupDesciption", "SELECT GroupPerson AS 'Группа', GroupPersonDescription AS 'Описание группы', Reserv1 AS 'Колличество в группе' ", " group by GroupPerson ORDER BY GroupPerson asc; ");

            try
            {
                DataGridViewSeekValuesInSelectedRow dgSeek = new DataGridViewSeekValuesInSelectedRow();
                dgSeek.FindValueInCells(dataGridView1, new string[] { @"Группа", @"Описание группы" });
                textBoxGroup.Text = dgSeek.values[0];
                textBoxGroupDescription.Text = dgSeek.values[1];
                StatusLabel2.Text = @"Выбрана группа: " + dgSeek.values[0] + "(" + dgSeek.values[1] + ")" + @" |  Всего групп: " + _dataGridView1RowsCount();

                QuickLoadDataItem.BackColor = Color.PaleGreen;
                groupBoxPeriod.BackColor = Color.PaleGreen;
                groupBoxTimeStart.BackColor = Color.PaleGreen;
                groupBoxTimeEnd.BackColor = Color.PaleGreen;
                groupBoxFilterReport.BackColor = SystemColors.Control;
            } catch { }

            DeleteGroupItem.Visible = true;
            dataGridView1.Visible = true;
            MembersGroupItem.Enabled = true;
            nameOfLastTableFromDB = "PeopleGroupDesciption";
            PersonOrGroupItem.Text = "Работа с одной персоной";
        }

        private void MembersGroupItem_Click(object sender, EventArgs e)//SearchMembersSelectedGroup()
        { SearchMembersSelectedGroup(); }

        private void SearchMembersSelectedGroup()
        {
            if (nameOfLastTableFromDB == "PeopleGroup" || nameOfLastTableFromDB == "PeopleGroupDesciption")
            {
                DataGridViewSeekValuesInSelectedRow dgSeek = new DataGridViewSeekValuesInSelectedRow();
                dgSeek.FindValueInCells(dataGridView1, new string[] { @"Группа" });
                SeekAndShowMembersOfGroup(dgSeek.values[0]);
            }
        }

        private void SeekAndShowMembersOfGroup(string nameGroup)
        {
            dtPeopleListLoaded?.Dispose();
            dtPeopleListLoaded = dtPeople.Clone();
            var dtTemp = dtPeople.Clone();

            numberPeopleInLoading = 0;
            DataRow dataRow;
            string d1 = "", dprtmnt = "", pstn = "", shift = "", query = ""; ;

            query = "Select FIO, NAV, GroupPerson, ControllingHHMM, ControllingOUTHHMM, Shift, Comment, Reserv1, Reserv2 FROM PeopleGroup ";
            if (nameGroup.Length > 0)
            { query += " where GroupPerson like '" + nameGroup + "'"; }
            query += ";";
            logger.Info("SeekAndShowMembersOfGroup: " + query);
            using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
            {
                sqlConnection.Open();
                using (var sqlCommand = new SQLiteCommand(query, sqlConnection))
                {
                    using (var sqlReader = sqlCommand.ExecuteReader())
                    {
                        foreach (DbDataRecord record in sqlReader)
                        {
                            d1 = "0"; dprtmnt = ""; pstn = ""; shift = "";
                            try { d1 = record["FIO"].ToString().Trim(); } catch { }

                            if (record != null && d1.Length > 0)
                            {
                                try { dprtmnt = record[@"Reserv1"].ToString().Trim(); } catch { dprtmnt = record[@"GroupPerson"]?.ToString(); }
                                try { pstn = record[@"Reserv2"].ToString().Trim(); } catch { }
                                try { shift = record[@"Shift"].ToString().Trim(); } catch { }

                                dataRow = dtTemp.NewRow();
                                dataRow[@"Фамилия Имя Отчество"] = d1;
                                dataRow[@"NAV-код"] = record[@"NAV"]?.ToString();

                                dataRow[@"Группа"] = record[@"GroupPerson"]?.ToString();
                                dataRow[@"Отдел"] = dprtmnt;
                                dataRow[@"Должность"] = pstn;

                                dataRow[@"Учетное время прихода ЧЧ:ММ"] = record[@"ControllingHHMM"]?.ToString();
                                dataRow[@"Учетное время ухода ЧЧ:ММ"] = record[@"ControllingOUTHHMM"]?.ToString();

                                dataRow[@"Комментарии"] = record["Comment"]?.ToString();
                                dataRow[@"График"] = record[@"Shift"]?.ToString();

                                dtTemp.Rows.Add(dataRow);
                                numberPeopleInLoading++;
                            }
                        }
                        d1 = null;
                    }
                }
            }

            if (numberPeopleInLoading > 0)
            {
                var namesDistinctCollumnsArray = arrayAllColumnsDataTablePeople.Except(arrayHiddenColumnsFIO).ToArray(); //take distinct data
                dtPeopleListLoaded = dtTemp.Copy();
                dtPersonTemp = GetDistinctRecords(dtTemp, namesDistinctCollumnsArray);

                ShowDatatableOnDatagridview(dtPersonTemp, arrayHiddenColumnsFIO);
                _MenuItemVisible(DeletePersonFromGroupItem, true);
                nameOfLastTableFromDB = "PeopleGroup";
            }

            query = null;
            dataRow = null;
            dtTemp.Dispose();
        }


        private void importPeopleInLocalDBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dtPeopleListLoaded?.Clear();
            dtPeopleListLoaded = dtPeople.Copy();

            ImportTextToTable(dtPeopleListLoaded);
            ImportTablePeopleToTableGroupsInLocalDB(databasePerson.ToString(), dtPeopleListLoaded);
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
                    DataRow row = dt.NewRow();

                    foreach (string strRow in listRows)
                    {
                        string[] cell = strRow.Split('\t');
                        if (cell.Length == 7)
                        {
                            row[@"Фамилия Имя Отчество"] = cell[0];
                            row[@"NAV-код"] = cell[1];
                            row[@"Группа"] = cell[2];
                            row[@"Отдел"] = cell[2];
                            listGroups.Add(cell[2]);

                            checkHourS = cell[3];
                            if (TryParseStringToDecimal(checkHourS) == 0)
                            { checkHourS = numUpHourStart.ToString(); }
                            row[@"Время прихода"] = ConvertStringsTimeToDecimal(checkHourS, cell[4]);
                            row[@"Учетное время прихода ЧЧ:ММ"] = ConvertStringsTimeToStringHHMM(checkHourS, cell[4]);

                            checkHourE = cell[5];
                            if (TryParseStringToDecimal(checkHourE) == 0)
                            { checkHourE = numUpDownHourEnd.ToString(); }
                            row[@"Время ухода"] = ConvertStringsTimeToDecimal(checkHourE, cell[6]);
                            row[@"Учетное время ухода ЧЧ:ММ"] = ConvertStringsTimeToStringHHMM(checkHourE, cell[6]);

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

        private void ImportTablePeopleToTableGroupsInLocalDB(string pathToPersonDB, DataTable dtSource) //use listGroups /add reserv1 reserv2
        {
            using (var sqlConnection = new SQLiteConnection($"Data Source={pathToPersonDB};Version=3;"))
            {
                sqlConnection.Open();

                //import groups
                SQLiteCommand commandTransaction = new SQLiteCommand("begin", sqlConnection);
                commandTransaction.ExecuteNonQuery();
                foreach (var dr in dtSource.AsEnumerable())
                {
                    if (dr[@"Фамилия Имя Отчество"] != null && dr[@"NAV-код"] != null && dr[@"NAV-код"]?.ToString().Length > 0)
                    {
                        using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'PeopleGroup' (FIO, NAV, GroupPerson, ControllingHHMM, ControllingOUTHHMM, Shift, Comment, Reserv1, Reserv2) " +
                                " VALUES (@FIO, @NAV, @GroupPerson, @ControllingHHMM, @ControllingOUTHHMM, @Shift, @Comment, @Reserv1, @Reserv2)", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@FIO", DbType.String).Value = dr[@"Фамилия Имя Отчество"].ToString();
                            sqlCommand.Parameters.Add("@NAV", DbType.String).Value = dr[@"NAV-код"].ToString();

                            sqlCommand.Parameters.Add("@GroupPerson", DbType.String).Value = dr[@"Группа"].ToString();
                            sqlCommand.Parameters.Add("@Reserv1", DbType.String).Value = dr[@"Отдел"].ToString();
                            sqlCommand.Parameters.Add("@Reserv2", DbType.String).Value = dr[@"Должность"].ToString();

                            sqlCommand.Parameters.Add("@ControllingHHMM", DbType.String).Value = dr[@"Учетное время прихода ЧЧ:ММ"].ToString();
                            sqlCommand.Parameters.Add("@ControllingOUTHHMM", DbType.String).Value = dr[@"Учетное время ухода ЧЧ:ММ"].ToString();

                            sqlCommand.Parameters.Add("@Shift", DbType.String).Value = dr[@"График"].ToString();
                            sqlCommand.Parameters.Add("@Comment", DbType.String).Value = dr[@"Комментарии"].ToString();

                            try
                            { sqlCommand.ExecuteNonQuery(); } catch (Exception expt)
                            {
                                logger.Info("GetFIO: ошибка записи в базу: " + dr[@"Фамилия Имя Отчество"] + "\n" + dr[@"NAV-код"] + "\n" + expt.ToString());
                            }
                        }
                    }
                }

                commandTransaction = new SQLiteCommand("end", sqlConnection);
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
                using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PeopleGroupDesciption' (GroupPerson, GroupPersonDescription) " +
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
            string group = _textBoxReturnText(textBoxGroup);
            string groupDescription = _textBoxReturnText(textBoxGroupDescription);
            string[] timeIn = { "00", "00", "00:00" };
            string[] timeOut = { "00", "00", "00:00" };
            decimal[] timeInDecimal = { 0, 0, 0, 0 };
            decimal[] timeOutDecimal = { 0, 0, 0, 0 };

            if (_dataGridView1CurrentRowIndex() > -1)
            {
                DataGridViewSeekValuesInSelectedRow dgSeek = new DataGridViewSeekValuesInSelectedRow();
                dgSeek.FindValueInCells(dataGridView1, new string[] {
                 @"Фамилия Имя Отчество", @"NAV-код", @"Отдел", @"Должность", @"Учетное время прихода ЧЧ:ММ", @"Учетное время ухода ЧЧ:ММ"
                            });

                timeInDecimal = ConvertStringTimeHHMMToDecimalArray(dgSeek.values[4]);
                timeOutDecimal = ConvertStringTimeHHMMToDecimalArray(dgSeek.values[5]);

                timeIn = ConvertDecimalTimeToStringHHMMArray(timeInDecimal[2]);
                timeOut = ConvertDecimalTimeToStringHHMMArray(timeOutDecimal[2]);

                using (var connection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    connection.Open();
                    if (group.Length > 0)
                    {
                        using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PeopleGroupDesciption' (GroupPerson, GroupPersonDescription) " +
                                                "VALUES (@GroupPerson, @GroupPersonDescription )", connection))
                        {
                            command.Parameters.Add("@GroupPerson", DbType.String).Value = group;
                            command.Parameters.Add("@GroupPersonDescription", DbType.String).Value = groupDescription;
                            try { command.ExecuteNonQuery(); } catch { }
                        }
                    }

                    if (group.Length > 0 && textBoxNav.Text.Trim().Length > 0)
                    {
                        timeInDecimal = ConvertStringTimeHHMMToDecimalArray(dgSeek.values[4]);
                        timeOutDecimal = ConvertStringTimeHHMMToDecimalArray(dgSeek.values[5]);

                        timeIn = ConvertDecimalTimeToStringHHMMArray(timeInDecimal[2]);
                        timeOut = ConvertDecimalTimeToStringHHMMArray(timeOutDecimal[2]);

                        using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PeopleGroup' (FIO, NAV, GroupPerson, ControllingHHMM, ControllingOUTHHMM, Reserv1, Reserv2) " +
                                                    "VALUES (@FIO, @NAV, @GroupPerson, @ControllingHHMM, @ControllingOUTHHMM, @Reserv1, @Reserv2)", connection))
                        {
                            command.Parameters.Add("@FIO", DbType.String).Value = dgSeek.values[0];
                            command.Parameters.Add("@NAV", DbType.String).Value = dgSeek.values[1];
                            command.Parameters.Add("@GroupPerson", DbType.String).Value = group;
                            command.Parameters.Add("@Reserv1", DbType.String).Value = dgSeek.values[2];
                            command.Parameters.Add("@Reserv2", DbType.String).Value = dgSeek.values[3];

                            command.Parameters.Add("@ControllingHHMM", DbType.String).Value = timeIn[2];
                            command.Parameters.Add("@ControllingOUTHHMM", DbType.String).Value = timeOut[2];
                            try { command.ExecuteNonQuery(); } catch { }
                        }
                        StatusLabel2.Text = "\"" + ShortFIO(dgSeek.values[0]) + "\"" + " добавлен в группу \"" + group + "\"";
                        _toolStripStatusLabelBackColor(StatusLabel2, SystemColors.Control);
                    }
                    else if (group.Length > 0 && dgSeek.values[1].Length == 0)
                        try
                        {
                            StatusLabel2.Text = "Отсутствует NAV-код у:" + ShortFIO(textBoxFIO.Text);
                            _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                        } catch { }
                    else if (group.Length == 0 && dgSeek.values[1].Length > 0)
                        try
                        {
                            StatusLabel2.Text = "Не указана группа, в которую нужно добавить!";
                            _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                        } catch { }
                }
            }
            SeekAndShowMembersOfGroup(group);

            PersonOrGroupItem.Text = "Работа с одной персоной";

            nameOfLastTableFromDB = "PeopleGroup";
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
        {
            string txtGroupBox = _textBoxReturnText(textBoxGroup);
            string txtFIO = _textBoxReturnText(textBoxFIO);
            await Task.Run(() => GetData(txtGroupBox, txtFIO));
        }

        private async void GetData(string groups, string fio)
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
                GetNamePoints();  //Get names of the points

                if ((nameOfLastTableFromDB == "PeopleGroupDesciption" || nameOfLastTableFromDB == "PeopleGroup") && groups.Length > 0)
                {
                    _toolStripStatusLabelSetText(StatusLabel2, "Получаю данные по группе " + groups);
                    stimerPrev = "Получаю данные по группе " + groups;
                }
                else if (nameOfLastTableFromDB == "Mailing")
                {
                    _toolStripStatusLabelSetText(StatusLabel2, "Загружаю данные по группе " + groups);
                    stimerPrev = "Загружаю регистрации с СКД сервера";
                }
                else
                {
                    _toolStripStatusLabelSetText(StatusLabel2, "Получаю данные по \"" + ShortFIO(fio) + "\" ");
                    stimerPrev = "Получаю данные по \"" + ShortFIO(fio) + "\"";
                    nameOfLastTableFromDB = "PersonRegistrationsList";
                }

                dtPersonRegistrationsFullList.Clear();
                logger.Info("GetData: " + groups);

                reportStartDay = _dateTimePickerStart().Split(' ')[0];
                reportLastDay = _dateTimePickerEnd().Split(' ')[0];

                GetRegistrations(groups, _dateTimePickerStart(), _dateTimePickerEnd(), "");
                logger.Info("GetData: " + groups + " " + _dateTimePickerStart() + " " + _dateTimePickerEnd());

                dtPersonTemp = dtPersonRegistrationsFullList.Copy();
                dtPersonTempAllColumns = dtPersonRegistrationsFullList.Copy(); //store all columns

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
            string fio, nav, group, dep, pos, timein, timeout, comment, shift;
            if ((nameOfLastTableFromDB == "PeopleGroupDesciption" || nameOfLastTableFromDB == "PeopleGroup" || nameOfLastTableFromDB == "Mailing" ||
                nameOfLastTableFromDB == "ListFIO" || doPostAction == "sendEmail") && selectedGroup.Length > 0)
            {
                //  if (doPostAction != "sendEmail")
                { LoadGroupMembersFromDbToDataTable(selectedGroup); } //result will be in dtPeopleGroup

                logger.Info("GetRegistrations, DT - " + dtPeopleGroup.TableName + " , всего записей - " + dtPeopleGroup.Rows.Count);
                foreach (DataRow row in dtPeopleGroup.Rows)
                {
                    if (row[@"Фамилия Имя Отчество"]?.ToString().Length > 0 && row[@"Группа"]?.ToString() == selectedGroup)
                    {
                        person = new Person();
                        if (!(doPostAction == "sendEmail"))
                        {
                            _textBoxSetText(textBoxFIO, row[@"Фамилия Имя Отчество"]?.ToString());   //иммитируем выбор данных
                            _textBoxSetText(textBoxNav, row[@"NAV-код"]?.ToString());   //Select person                  
                        }

                        // string fio, nav, group, dep, pos, timein, timeout, comment, shift;
                        fio = row[@"Фамилия Имя Отчество"]?.ToString();
                        nav = row[@"NAV-код"]?.ToString();

                        group = row[@"Группа"]?.ToString();
                        dep = row[@"Отдел"]?.ToString();
                        pos = row[@"Должность"]?.ToString();

                        //todo
                        //import all of day's shift for the current person
                        timein = row[@"Учетное время прихода ЧЧ:ММ"]?.ToString();
                        timeout = row[@"Учетное время ухода ЧЧ:ММ"]?.ToString();

                        comment = row[@"Комментарии"]?.ToString();
                        shift = row[@"График"]?.ToString();

                        person.FIO = fio;
                        person.NAV = nav;

                        person.GroupPerson = group;
                        person.Department = dep;
                        person.PositionInDepartment = pos;

                        person.ControlInHHMM = timein;
                        person.ControlOutHHMM = timeout;

                        person.ControlInDecimal = ConvertStringTimeHHMMToDecimalArray(timein)[2];
                        person.ControlOutDecimal = ConvertStringTimeHHMMToDecimalArray(timeout)[2];

                        person.Comment = comment;
                        person.Shift = shift;

                        GetPersonRegistrationFromServer(dtPersonRegistrationsFullList, person, startDate, endDate);     //Search Registration at checkpoints of the selected person
                    }
                }
                nameOfLastTableFromDB = "PeopleGroup";
                _toolStripStatusLabelSetText(StatusLabel2, "Данные по группе \"" + selectedGroup + "\" получены");
            }
            else
            {
                person = new Person();
                person.NAV = _textBoxReturnText(textBoxNav);
                person.FIO = _textBoxReturnText(textBoxFIO);
                logger.Info("GetRegistrations " + person.FIO);
                person.GroupPerson = "";
                person.Department = "";
                person.Shift = "";
                person.Comment = "";
                person.PositionInDepartment = "Сотрудник";

                person.ControlInHHMM = ConvertStringsTimeToStringHHMM(_numUpDownReturn(numUpDownHourStart).ToString(), _numUpDownReturn(numUpDownMinuteStart).ToString());
                person.ControlOutHHMM = ConvertStringsTimeToStringHHMM(_numUpDownReturn(numUpDownHourEnd).ToString(), _numUpDownReturn(numUpDownMinuteEnd).ToString());

                GetPersonRegistrationFromServer(dtPersonRegistrationsFullList, person, startDate, endDate);

                _toolStripStatusLabelSetText(StatusLabel2, "Данные с СКД по \"" + ShortFIO(_textBoxReturnText(textBoxFIO)) + "\" получены!");
            }

            person = null;
        }

        private void GetPersonRegistrationFromServer(DataTable dtTarget, Person person, string startDay, string endDay)
        {
            int[] startPeriod = ConvertStringDateToIntArray(startDay);
            int[] endPeriod = ConvertStringDateToIntArray(endDay);

            SeekAnualDays(dtTarget, person, false, startPeriod[0], startPeriod[1], startPeriod[2], endPeriod[0], endPeriod[1], endPeriod[2]);

            DataRow rowPerson;
            string stringConnection = "";
            string query = "";
            HashSet<string> personWorkedDays = new HashSet<string>();

            string stringIdCardIntellect = "";
            string personNAVTemp = "";
            string[] stringSelectedFIO = { "", "", "" };
            try { stringSelectedFIO[0] = Regex.Split(person.FIO, "[ ]")[0]; } catch { }
            try { stringSelectedFIO[1] = Regex.Split(person.FIO, "[ ]")[1]; } catch { }
            try { stringSelectedFIO[2] = Regex.Split(person.FIO, "[ ]")[2]; } catch { }

            //is looking for a NAV and an idCard
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
                                    if (record?["tabnum"] != null && record?["tabnum"].ToString().Trim() == person.NAV)
                                    {
                                        stringIdCardIntellect = record["id"].ToString().Trim();
                                        person.idCard = Convert.ToInt32(record["id"].ToString().Trim());
                                        break;
                                    }
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
                        if (strRowWithNav.Contains(person.NAV) && person.NAV.Length > 0)
                            try
                            {
                                stringIdCardIntellect = Regex.Split(strRowWithNav, "[|]")[3].Trim();
                                person.idCard = Convert.ToInt32(stringIdCardIntellect);
                                break;
                            } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                    }

                    if (stringIdCardIntellect.Length == 0)
                    {
                        foreach (var strRowWithNav in listFIO.ToArray())
                        {
                            if (strRowWithNav.ToLower().Contains(stringSelectedFIO[0].ToLower().Trim()) &&
                                strRowWithNav.ToLower().Contains(stringSelectedFIO[1].ToLower().Trim()) &&
                                strRowWithNav.ToLower().Contains(stringSelectedFIO[2].ToLower().Trim()))
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
                                    _ProgressWork1Step(1); break;
                                }
                            }
                        }
                    }
                }
            } catch (Exception expt) { MessageBox.Show(expt.ToString()); }

            string[] cellData;
            string namePoint = "";
            string direction = "";

            try
            {
                stringConnection = @"Data Source=" + sServer1 + @"\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID=" + sServer1UserName + ";Password=" + sServer1UserPassword + "; Connect Timeout=120";
                query = "SELECT param0, param1, objid, CONVERT(varchar, date, 120) AS date, CONVERT(varchar, PROTOCOL.time, 114) AS time FROM protocol " +
                   " where action like 'ACCESS_IN' AND param1 like '" + stringIdCardIntellect + "' AND date >= '" + startDay + "' AND date <= '" + endDay + "' " +
                   " ORDER BY date ASC";
                string stringDataNew = "";
                int idCardIntellect = 0;
                string fullPointName = "";
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
                                        if (record["param0"] != null && record["param0"]?.ToString().Trim().Length > 0)
                                        {
                                            stringDataNew = Regex.Split(record["date"].ToString().Trim(), "[ ]")[0];
                                            hourManaging = Convert.ToDecimal(Regex.Split(record["time"].ToString().Trim(), "[:]")[0]);
                                            minuteManaging = Convert.ToDecimal(Regex.Split(record["time"].ToString().Trim(), "[:]")[1]);
                                            managingHours = ConvertDecimalSeparatedTimeToDecimal(hourManaging, minuteManaging);
                                            person.idCard = Convert.ToInt32(record["param1"].ToString().Trim());

                                            namePoint = "";
                                            direction = "";
                                            fullPointName = record["objid"]?.ToString().Trim();

                                            foreach (var aPointPass in listPoints.ToArray())
                                            {
                                                if (aPointPass != null && sServer1 != null && fullPointName != null &&
                                                    aPointPass.Contains(sServer1) && aPointPass.Contains(fullPointName) &&
                                                    aPointPass.Contains("|") && fullPointName.Length ==
                                                    Regex.Split(aPointPass, "[|]")[1].Length)
                                                    try
                                                    {
                                                        namePoint = Regex.Split(aPointPass, "[|]")[2];
                                                        if (namePoint.ToLower().Contains("выход"))
                                                            direction = "Выход";
                                                        else if (namePoint.ToLower().Contains("вход"))
                                                            direction = "Вход";
                                                        break;
                                                    } catch { }
                                            }
                                            personWorkedDays.Add(stringDataNew);

                                            rowPerson = dtTarget.NewRow();
                                            rowPerson[@"Фамилия Имя Отчество"] = record["param0"]?.ToString().Trim();
                                            rowPerson[@"NAV-код"] = person.NAV;

                                            rowPerson[@"№ пропуска"] = record["param1"]?.ToString().Trim();
                                            rowPerson[@"Группа"] = person.GroupPerson;
                                            rowPerson[@"Отдел"] = person.Department;
                                            rowPerson[@"Должность"] = person.PositionInDepartment;

                                            rowPerson[@"График"] = person.Shift;

                                            //day of registration //real data
                                            rowPerson[@"Дата регистрации"] = stringDataNew;
                                            rowPerson[@"Время регистрации"] = TryParseStringToDecimal(managingHours.ToString("#.###"));
                                            rowPerson[@"Сервер СКД"] = sServer1;
                                            rowPerson[@"Имя точки прохода"] = namePoint;
                                            rowPerson[@"Направление прохода"] = direction;

                                            rowPerson[@"Учетное время прихода ЧЧ:ММ"] = person.ControlInHHMM;
                                            rowPerson[@"Учетное время ухода ЧЧ:ММ"] = person.ControlOutHHMM;
                                            rowPerson[@"Фактич. время прихода ЧЧ:ММ"] = ConvertStringsTimeToStringHHMM(hourManaging.ToString(), minuteManaging.ToString());

                                            dtTarget.Rows.Add(rowPerson);

                                            _ProgressWork1Step(1);
                                        }
                                    } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                                }
                            }
                        }
                    }
                }

                stringDataNew = null; query = null; stringConnection = null;

                _ProgressWork1Step(1);
            } catch (Exception Expt)
            { MessageBox.Show(Expt.ToString(), @"Сервер не доступен, или неправильная авторизация", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            if (listPoints.Count > 0)
            { bLoaded = true; }

            // рабочие дни в которые отсутствовал данная персона
            foreach (string day in workSelectedDays.Except(personWorkedDays))
            {
                rowPerson = dtTarget.NewRow();
                rowPerson[@"Фамилия Имя Отчество"] = person.FIO;
                rowPerson[@"NAV-код"] = person.NAV;
                rowPerson[@"Группа"] = person.GroupPerson;
                rowPerson[@"№ пропуска"] = person.idCard;

                rowPerson[@"Отдел"] = person.Department;
                rowPerson[@"Должность"] = person.PositionInDepartment;

                rowPerson[@"График"] = person.Shift;

                rowPerson[@"Время регистрации"] = "0";
                rowPerson[@"Дата регистрации"] = day;
                rowPerson[@"День недели"] = DayOfWeekRussian((DateTime.Parse(day)).DayOfWeek.ToString());
                rowPerson[@"Учетное время прихода ЧЧ:ММ"] = person.ControlInHHMM;
                rowPerson[@"Учетное время ухода ЧЧ:ММ"] = person.ControlOutHHMM;
                rowPerson[@"Отсутствовал на работе"] = "Да";

                dtTarget.Rows.Add(rowPerson);//добавляем рабочий день в который  сотрудник не выходил на работу
                _ProgressWork1Step(1);
            }

            //look for late and absence of data in www's DB
            outResons = new List<OutReasons>();

            List<OutPerson> outPerson = new List<OutPerson>();
            outResons.Add(new OutReasons() { _id = "0", _hourly = 1, _name = @"Unknow", _visibleName = @"Неизвестная" });

            stringConnection = @"server=" + mysqlServer + @";User=" + mysqlServerUserName + @";Password=" + mysqlServerUserPassword + @";database=wwwais;pooling = false; convert zero datetime=True;Connect Timeout=60";
            logger.Info(stringConnection);
            using (var sqlConnection = new MySql.Data.MySqlClient.MySqlConnection(stringConnection))
            {
                sqlConnection.Open();
                query = "Select id,name,hourly,visibled_name FROM out_reasons";
                logger.Trace(query);

                using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(query, sqlConnection))
                {
                    using (MySql.Data.MySqlClient.MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (reader.GetString(@"name") != null && reader.GetString(@"name").Length > 0)
                            {
                                outResons.Add(new OutReasons()
                                {
                                    _id = reader.GetString(@"id"),
                                    _hourly = Convert.ToInt32(reader.GetString(@"hourly")),
                                    _name = reader.GetString(@"name"),
                                    _visibleName = reader.GetString(@"visibled_name")
                                });
                            }
                        }
                    }
                }
                _ProgressWork1Step(1);

                string date = "";
                string resonId = "";
                query = "Select * FROM out_users where user_code='" + person.NAV + "' AND reason_date >= '" + startDay.Split(' ')[0] + "' AND reason_date <= '" + endDay.Split(' ')[0] + "' ";
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
                                resonId = outResons.Find((x) => x._id == reader?.GetString(@"reason_id"))?._id;
                                try { date = DateTime.Parse(reader.GetString(@"reason_date")).ToString("yyyy-MM-dd"); } catch { date = ""; }

                                outPerson.Add(new OutPerson()
                                {
                                    _reason_id = resonId,
                                    _nav = reader.GetString(@"user_code").ToUpper().Replace('C', 'S'),
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
            logger.Info(person.NAV + " - на сайте всего записей с отсутствиями: " + outPerson.Count);
            _ProgressWork1Step(1);

            string nav = "";
            foreach (var row in dtTarget.AsEnumerable())
            {
                if (row[@"NAV-код"]?.ToString() == person.NAV)
                {
                    row[@"Фамилия Имя Отчество"] = person.FIO;
                    row[@"NAV-код"] = person.NAV;
                    row[@"Группа"] = person.GroupPerson;
                    row[@"№ пропуска"] = person.idCard;
                    row[@"Отдел"] = person.Department;
                    row[@"Должность"] = person.PositionInDepartment;
                    row[@"График"] = person.Shift;
                }
            }

            foreach (DataRow dr in dtTarget.Rows) // search whole table
            {
                nav = dr[@"NAV-код"]?.ToString();
                foreach (string day in workSelectedDays)
                {
                    if (dr[@"Дата регистрации"].ToString() == day)
                    {
                        // todo - remove try and check
                        try
                        {
                            //set Comments
                            dr[@"Комментарии"] = outPerson.Find((x) => x._date == day && x._nav == nav)?._reason_id;
                        } catch { }
                        break;
                    }
                }
            }

            _ProgressWork1Step(1);

            listRegistrations.Clear(); rowPerson = null;
            namePoint = null; direction = null;
            stringIdCardIntellect = null; personNAVTemp = null; stringSelectedFIO = new string[1]; cellData = new string[1];
        }

        private string DayOfWeekRussian(string dayEnglish) //return a day of week as the same short name in Russian 
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
        private void LoadGroupMembersFromDbToDataTable(string namePointedGroup) // dtPeopleGroup //"Select * FROM PeopleGroup where GroupPerson like '" + _textBoxReturnText(textBoxGroup) + "';"
        {
            dtPeopleGroup.Clear();
            //  dtPeopleGroup = dtPeople.Clone();
            DataRow dataRow;

            string query = "Select FIO, NAV, GroupPerson, ControllingHHMM, ControllingOUTHHMM, Shift, Comment, Reserv1, Reserv2 FROM PeopleGroup ";
            if (namePointedGroup.Length > 0)
            { query += "where GroupPerson like '" + namePointedGroup + "'; "; }
            else { query += ";"; }

            using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
            {
                sqlConnection.Open();
                using (var sqlCommand = new SQLiteCommand(query, sqlConnection))
                {
                    using (var sqlReader = sqlCommand.ExecuteReader())
                    {
                        foreach (DbDataRecord record in sqlReader)
                        {
                            try
                            {
                                if (record[@"FIO"] != null && record[@"NAV"] != null)
                                {
                                    dataRow = dtPeopleGroup.NewRow();

                                    dataRow[@"Фамилия Имя Отчество"] = record[@"FIO"].ToString();
                                    dataRow[@"NAV-код"] = record[@"NAV"].ToString();

                                    dataRow[@"Группа"] = record[@"GroupPerson"].ToString();
                                    dataRow[@"Отдел"] = record[@"Reserv1"].ToString();
                                    dataRow[@"Должность"] = record[@"Reserv2"].ToString();

                                    dataRow[@"Комментарии"] = record[@"Comment"].ToString();
                                    dataRow[@"График"] = record[@"Shift"].ToString();

                                    dataRow[@"Учетное время прихода ЧЧ:ММ"] = record[@"ControllingHHMM"].ToString();
                                    dataRow[@"Учетное время ухода ЧЧ:ММ"] = record[@"ControllingOUTHHMM"].ToString();

                                    dtPeopleGroup.Rows.Add(dataRow);
                                }
                            } catch { }
                        }
                    }
                }
            }
            query = null; dataRow = null;

            logger.Info("LoadGroupMembersFromDbToDataTable, всего записей в dtPeopleGroup " + dtPeopleGroup.Rows.Count + ", группа - " + namePointedGroup);
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
                DataGridViewSeekValuesInSelectedRow dgSeek = new DataGridViewSeekValuesInSelectedRow();
                dgSeek.FindValueInCells(dataGridView1, new string[] { @"Группа", @"NAV-код" });
                DeleteDataTableQueryNAV(databasePerson, "PeopleGroup", "NAV", dgSeek.values[1], "GroupPerson", dgSeek.values[0]);
                SeekAndShowMembersOfGroup(dgSeek.values[0]);
            }

            nameOfLastTableFromDB = "PeopleGroup";
            _controlVisible(dataGridView1, true);
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
            int[] startPeriod = _dateTimePickerReturnArray(dateTimePickerStart);
            int[] endPeriod = _dateTimePickerReturnArray(dateTimePickerEnd);
            DataTable dtEmpty = new DataTable();
            Person emptyPerson = new Person();
            SeekAnualDays(dtEmpty, emptyPerson, false, startPeriod[0], startPeriod[1], startPeriod[2], endPeriod[0], endPeriod[1], endPeriod[2]);

            dtEmpty.Dispose();
            emptyPerson = null;

            CheckBoxesFiltersAll_Enable(false);
            _controlVisible(dataGridView1, false);
            _controlVisible(pictureBox1, false);

            string nameGroup = _textBoxReturnText(textBoxGroup);

            DataTable dtTempIntermediate = dtPeople.Clone();
            dtPersonTempAllColumns = dtPeople.Clone();
            Person personCheck = new Person();
            personCheck.FIO = _textBoxReturnText(textBoxFIO);
            personCheck.NAV = _textBoxReturnText(textBoxNav);
            personCheck.GroupPerson = nameGroup;
            personCheck.Department = nameGroup;

            personCheck.ControlInDecimal = ConvertDecimalSeparatedTimeToDecimal(numUpHourStart, numUpMinuteStart);
            personCheck.ControlOutDecimal = ConvertDecimalSeparatedTimeToDecimal(numUpHourEnd, numUpMinuteEnd);
            personCheck.ControlInHHMM = ConvertDecimalTimeToStringHHMM(numUpHourStart, numUpMinuteStart);
            personCheck.ControlOutHHMM = ConvertDecimalTimeToStringHHMM(numUpHourEnd, numUpMinuteEnd);
            dtPersonTemp?.Clear();

            if ((nameOfLastTableFromDB == "PeopleGroupDesciption" || nameOfLastTableFromDB == "PeopleGroup") && nameGroup.Length > 0)
            {
                LoadGroupMembersFromDbToDataTable(nameGroup); //result will be in dtPeopleGroup  //"Select * FROM PeopleGroup where GroupPerson like '" + _textBoxReturnText(textBoxGroup) + "';"

                if (_CheckboxCheckedStateReturn(checkBoxReEnter))
                {
                    foreach (DataRow row in dtPeopleGroup.Rows)
                    {
                        if (row[@"Фамилия Имя Отчество"]?.ToString().Length > 0 && row[@"Группа"]?.ToString() == nameGroup)
                        {
                            personCheck = new Person();

                            personCheck.FIO = row[@"Фамилия Имя Отчество"].ToString();
                            personCheck.NAV = row[@"NAV-код"].ToString();

                            personCheck.GroupPerson = row[@"Группа"].ToString();
                            personCheck.Department = row[@"Отдел"].ToString();
                            personCheck.PositionInDepartment = row[@"Должность"].ToString();

                            personCheck.ControlInDecimal = ConvertStringTimeHHMMToDecimalArray(row[@"Учетное время прихода ЧЧ:ММ"].ToString())[2];
                            personCheck.ControlOutDecimal = ConvertStringTimeHHMMToDecimalArray(row[@"Учетное время ухода ЧЧ:ММ"].ToString())[2];
                            personCheck.ControlInHHMM = row[@"Учетное время прихода ЧЧ:ММ"].ToString();
                            personCheck.ControlOutHHMM = row[@"Учетное время ухода ЧЧ:ММ"].ToString();

                            personCheck.Comment = row[@"Комментарии"].ToString();
                            personCheck.Shift = row[@"График"].ToString();
                            FilterDataByNav(personCheck, dtPersonRegistrationsFullList, dtTempIntermediate);
                        }
                    }
                }
                else
                { dtTempIntermediate = dtPersonRegistrationsFullList.Select("[Группа] = '" + nameGroup + "'").Distinct().CopyToDataTable(); }
            }
            else
            {
                if (!_CheckboxCheckedStateReturn(checkBoxReEnter))
                { dtTempIntermediate = dtPersonRegistrationsFullList.Copy(); }
                else
                { FilterDataByNav(personCheck, dtPersonRegistrationsFullList, dtTempIntermediate); }
            }

            //store all columns
            dtPersonTempAllColumns = dtTempIntermediate.Copy();

            //show selected data     
            //distinct Records        
            var namesDistinctColumnsArray = arrayAllColumnsDataTablePeople.Except(nameHidenColumnsArray).ToArray(); //take distinct data
            dtPersonTemp = GetDistinctRecords(dtTempIntermediate, namesDistinctColumnsArray);

            ShowDatatableOnDatagridview(dtPersonTemp, nameHidenColumnsArray);

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

        //Copy Data from dtPersonRegistrationsList into dtPersonTemp by Filter(NAV and anual dates or minimalTime or dayoff)
        private void FilterDataByNav(Person person, DataTable dataTableSource, DataTable dataTableForStoring)
        {
            logger.Info("FilterDataByNav: " + person.NAV + "| dataTableSource: " + dataTableSource.Rows.Count);
            DataRow rowDtStoring;
            DataTable dtTemp = dataTableSource.Clone();

            HashSet<string> hsDays = new HashSet<string>();
            DataTable dtAllRegistrationsInSelectedDay = dataTableSource.Clone(); //All registrations in the selected day

            decimal decimalFirstRegistrationInDay;
            string[] stringHourMinuteFirstRegistrationInDay = new string[2];
            decimal decimalLastRegistrationInDay;
            string[] stringHourMinuteLastRegistrationInDay = new string[2];
            decimal workedHours = 0;
            string dayOfWeek = "";
            string exceptReason = "";

            try
            {
                var allWorkedDaysPerson = dataTableSource.Select("[NAV-код] = '" + person.NAV + "'");

                if (_CheckboxCheckedStateReturn(checkBoxReEnter) || currentAction == "sendEmail") //checkBoxReEnter.Checked
                {
                    foreach (DataRow dataRowDate in allWorkedDaysPerson) //make the list of worked days
                    { hsDays.Add(dataRowDate[@"Дата регистрации"].ToString()); }

                    foreach (var workedDay in hsDays.ToArray())
                    {
                        dtAllRegistrationsInSelectedDay = allWorkedDaysPerson.Distinct().CopyToDataTable().Select("[Дата регистрации] = '" + workedDay + "'").CopyToDataTable();

                        //Select only one row with selected NAV for the selected workedDay

                        rowDtStoring = dtAllRegistrationsInSelectedDay.Select("[Дата регистрации] = '" + workedDay + "'").First();
                        dayOfWeek = DayOfWeekRussian(DateTime.Parse(workedDay).DayOfWeek.ToString());
                        rowDtStoring[@"День недели"] = dayOfWeek;

                        //find first registration within the during selected workedDay
                        decimalFirstRegistrationInDay = Convert.ToDecimal(dtAllRegistrationsInSelectedDay.Compute("MIN([Время регистрации])", string.Empty));

                        //find last registration within the during selected workedDay
                        decimalLastRegistrationInDay = Convert.ToDecimal(dtAllRegistrationsInSelectedDay.Compute("MAX([Время регистрации])", string.Empty));

                        //take and convert a real time coming into a string timearray
                        stringHourMinuteFirstRegistrationInDay = ConvertDecimalTimeToStringHHMMArray(decimalFirstRegistrationInDay);
                        rowDtStoring[@"Время регистрации"] = decimalFirstRegistrationInDay;              //("Время регистрации", typeof(decimal)), //15
                        rowDtStoring[@"Фактич. время прихода ЧЧ:ММ"] = stringHourMinuteFirstRegistrationInDay[2];  //("Фактич. время прихода ЧЧ:ММ", typeof(string)),//24

                        stringHourMinuteLastRegistrationInDay = ConvertDecimalTimeToStringHHMMArray(decimalLastRegistrationInDay);
                        rowDtStoring[@"Реальное время ухода"] = decimalLastRegistrationInDay;                 //("Реальное время ухода", typeof(decimal)), //18
                        rowDtStoring[@"Фактич. время ухода ЧЧ:ММ"] = stringHourMinuteLastRegistrationInDay[2];     //("Фактич. время ухода ЧЧ:ММ", typeof(string)), //25

                        //worked out times
                        workedHours = decimalLastRegistrationInDay - decimalFirstRegistrationInDay;
                        rowDtStoring[@"Реальное отработанное время"] = workedHours;                                  // ("Реальное отработанное время", typeof(decimal)), //26
                        rowDtStoring[@"Отработанное время ЧЧ:ММ"] = ConvertDecimalTimeToStringHHMMArray(workedHours)[2];  //("Отработанное время ЧЧ:ММ", typeof(string)), //27

                        //todo 
                        //will calculate if day of week different
                        if (decimalFirstRegistrationInDay > person.ControlInDecimal && decimalFirstRegistrationInDay != 0) // "Опоздание ЧЧ:ММ", typeof(bool)),           //28
                        {
                            rowDtStoring[@"Опоздание ЧЧ:ММ"] = ConvertDecimalTimeToStringHHMM(decimalFirstRegistrationInDay - person.ControlInDecimal);
                        }
                        else { rowDtStoring[@"Опоздание ЧЧ:ММ"] = ""; }

                        if (decimalLastRegistrationInDay < person.ControlOutDecimal && decimalLastRegistrationInDay != 0)  // "Ранний уход ЧЧ:ММ", typeof(bool)),                 //29
                        {
                            rowDtStoring[@"Ранний уход ЧЧ:ММ"] = ConvertDecimalTimeToStringHHMM(person.ControlOutDecimal - decimalLastRegistrationInDay);
                        }
                        else { rowDtStoring[@"Ранний уход ЧЧ:ММ"] = ""; }


                        exceptReason = rowDtStoring[@"Комментарии"].ToString();

                        rowDtStoring[@"Комментарии"] = outResons.Find((x) => x._id == exceptReason)?._visibleName;

                        switch (exceptReason)
                        {
                            case "2":
                            case "10":
                            case "11":
                            case "12":
                            case "18":
                                rowDtStoring[@"Отпуск"] = outResons.Find((x) => x._id == exceptReason)?._visibleName;
                                    break;
                            case "3":
                            case "21":
                                rowDtStoring[@"Больничный"] = outResons.Find((x) => x._id == exceptReason)?._visibleName;
                                rowDtStoring[@"Комментарии"] = "";
                                break;
                            case "1":
                            case "9":
                            case "20":
                                rowDtStoring[@"Отгул"] = outResons.Find((x) => x._id == exceptReason)?._visibleName;
                                break;
                            case "4":
                                rowDtStoring[@"Командировка"] = outResons.Find((x) => x._id == exceptReason)?._visibleName;
                                rowDtStoring[@"Комментарии"] = "";
                                break;
                            default:
                                break;
                        }


                        /*
                            @"Отпуск",                 //30
                            @"Командировка",                 //31
                            @"Больничный",                    //33
                            @"Комментарии",                    //38
                            @"Отгул"                      //41
                            */

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
                    int[] startPeriod = ConvertStringDateToIntArray(reportStartDay);
                    int[] endPeriod = ConvertStringDateToIntArray(reportLastDay);
                    DataTable dtEmpty = new DataTable();
                    Person emptyPerson = new Person();

                    SeekAnualDays(dtTemp, person, true, startPeriod[0], startPeriod[1], startPeriod[2], endPeriod[0], endPeriod[1], endPeriod[2]);
                    dtEmpty.Dispose();
                    emptyPerson = null;
                }

                if (_CheckboxCheckedStateReturn(checkBoxTimeViolations)) //checkBoxStartWorkInTime Checking
                { QueryDeleteDataFromDataTable(dtTemp, "[Опоздание ЧЧ:ММ]='' AND [Ранний уход ЧЧ:ММ]=''", person.NAV); }

                foreach (DataRow dr in dtTemp.AsEnumerable())
                { dataTableForStoring.ImportRow(dr); }

                allWorkedDaysPerson = null;
            } catch (Exception expt) { MessageBox.Show(expt.ToString()); }

            stringHourMinuteFirstRegistrationInDay = null; stringHourMinuteLastRegistrationInDay = null; hsDays = null;
            rowDtStoring = null; dtTemp = null; dtAllRegistrationsInSelectedDay = null;
        }





        /// <summary>
        /// ///////////////////////////////////////////////////////////////
        /// </summary>
        //check dubled function!!!!!!!!

        private void SeekAnualDays(DataTable dt, Person person, bool delRow, int startYear, int startMonth, int startDay, int endYear, int endMonth, int endDay) //Exclude Anual Days from the table "PersonTemp" DB
        {
            //test
            List<string> daysBolded = new List<string>();
            boldeddDates.Clear();
            selectedDates.Clear();
            workSelectedDays = new string[1];
            myBoldedDates = new string[1];

            var oneDay = TimeSpan.FromDays(1);
            var twoDays = TimeSpan.FromDays(2);


            var mySelectedStartDay = new DateTime(startYear, startMonth, startDay);
            var mySelectedEndDay = new DateTime(endYear, endMonth, endDay);
            //  int myYearNow = DateTime.Now.Year;
            var myMonthCalendar = new MonthCalendar();

            myMonthCalendar.MaxSelectionCount = 60;
            myMonthCalendar.SelectionRange = new SelectionRange(mySelectedStartDay, mySelectedEndDay);
            myMonthCalendar.FirstDayOfWeek = Day.Monday;

            for (int year = -1; year < 3; year++)
            {
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear + year, 1, 1));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear + year, 1, 2));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear + year, 1, 7));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear + year, 3, 8));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear + year, 5, 1));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear + year, 5, 2));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear + year, 5, 9));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear + year, 6, 28));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear + year, 8, 24));    // (plavayuschaya data)
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear + year, 10, 16));   // (plavayuschaya data)
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
                case (int)Day.Saturday:
                    myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear, 8, 24) + twoDays);    // (plavayuschaya data)
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
                case (int)Day.Saturday:
                    myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startYear, 10, 16) + twoDays);    // (plavayuschaya data)
                    break;
                default:
                    break;
            }

            string singleDate = null;


            List<string> daysSelected = new List<string>();
            for (var myDate = myMonthCalendar.SelectionStart; myDate <= myMonthCalendar.SelectionEnd; myDate += oneDay)
            {
                //todo - will do simplify "singleDate"
                singleDate = Regex.Split(myDate.ToString("yyyy-MM-dd"), " ")[0].Trim();
                if (myDate.DayOfWeek == DayOfWeek.Saturday || myDate.DayOfWeek == DayOfWeek.Sunday)
                {
                    daysBolded.Add(singleDate);
                    if (delRow)
                    {
                        QueryDeleteDataFromDataTable(dt, "[Дата регистрации]='" + singleDate + "'", person.NAV); // ("Дата регистрации",typeof(string)),//12
                    }
                }
                daysSelected.Add(singleDate);
            }
            foreach (var myAnualDate in myMonthCalendar.AnnuallyBoldedDates)
            {
                for (var myDate = myMonthCalendar.SelectionStart; myDate <= myMonthCalendar.SelectionEnd; myDate += oneDay)
                {
                    if (myDate == myAnualDate)
                    {
                        //todo - will do simplify "singleDate"
                        singleDate = Regex.Split(myDate.ToString("yyyy-MM-dd"), " ")[0].Trim();
                        daysBolded.Add(singleDate);
                        if (delRow)
                        {
                            QueryDeleteDataFromDataTable(dt, "[Дата регистрации]='" + singleDate + "'", person.NAV); // ("Дата регистрации",typeof(string)),//12
                        }
                    }
                }
            }
            myBoldedDates = daysBolded.ToArray();
            workSelectedDays = daysSelected.Except(daysBolded).ToArray();

            dt.AcceptChanges();

            myMonthCalendar.Dispose();
            daysBolded = null;
            daysSelected = null;
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
            monthCalendar.Refresh();
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



        //----- Clear ---------//
        private void ClearFilesInApplicationFolders(string maskFiles, string discribeFiles)
        {
            try
            {
                System.IO.FileInfo[] filesPath = new System.IO.DirectoryInfo(System.IO.Path.GetDirectoryName(filePathApplication)).GetFiles(maskFiles, System.IO.SearchOption.AllDirectories);
                foreach (System.IO.FileInfo file in filesPath)
                {
                    if (file.Length == 0) { continue; }
                    else
                    {
                        try
                        {
                            System.IO.File.Delete(file.FullName);
                            logger.Info("Удален файл: \"" + file.FullName + "\"");
                        } catch { logger.Warn("Файл не удален: \"" + file.FullName + "\""); }
                    }
                }
            } catch { logger.Warn("Ошибка удаления: " + discribeFiles); }

        }

        private void ClearReportItem_Click(object sender, EventArgs e) //ReCreatePersonTables()
        { CleaReportsRecreateTables(); }

        private void CleaReportsRecreateTables()
        {
            logger.Info("-= Очистика от отчетов =-");

            ClearFilesInApplicationFolders(@"*.xlsx", "Excel-файлов");

            _textBoxSetText(textBoxFIO, "");
            _textBoxSetText(textBoxGroup, "");
            _textBoxSetText(textBoxGroupDescription, "");
            _textBoxSetText(textBoxNav, "");

            GC.Collect();

            TryMakeDB();
            UpdateTableOfDB();

            DataTable dt = new DataTable();
            _dataGridViewSource(dt);

            _toolStripStatusLabelSetText(StatusLabel2, @"Временные таблицы удалены");
        }

        private void ClearDataItem_Click(object sender, EventArgs e) //ReCreateAllPeopleTables()
        { ClearGotReportsRecreateTables(); }

        private void ClearGotReportsRecreateTables()
        {
            logger.Info("-= Очистика от отчетов и полученных данных =-");
            DeleteTable(databasePerson, "LastTakenPeopleComboList");

            ClearFilesInApplicationFolders(@"*.xlsx", "Excel-файлов");
            ClearFilesInApplicationFolders(@"*.log", "логов");

            _textBoxSetText(textBoxFIO, "");
            _textBoxSetText(textBoxGroup, "");
            _textBoxSetText(textBoxGroupDescription, "");
            _textBoxSetText(textBoxNav, "");

            _comboBoxClr(comboBoxFio);
            listFIO.Clear();

            GC.Collect();
            iFIO = 0;
            TryMakeDB();
            UpdateTableOfDB();

            DataTable dt = new DataTable();
            _dataGridViewSource(dt);

            _toolStripStatusLabelSetText(StatusLabel2, @"База очищена от загруженных данных. Остались только созданные группы");
        }

        private void ClearAllItem_Click(object sender, EventArgs e) //ReCreate DB
        { ReCreateDB(); }

        private void ReCreateDB()
        {
            if (databasePerson.Exists)
            {
                logger.Info("-= Очистика локальной базы от всех полученных, сгенерированных, сохраненных и введенных данных =-");

                DeleteTable(databasePerson, "PeopleGroup");
                DeleteTable(databasePerson, "PeopleGroupDesciption");
                DeleteTable(databasePerson, "TechnicalInfo");
                DeleteTable(databasePerson, "BoldedDates");
                DeleteTable(databasePerson, "ProgramSettings");
                DeleteTable(databasePerson, "EquipmentSettings");
                DeleteTable(databasePerson, "ProgramSettings");
                DeleteTable(databasePerson, "LastTakenPeopleComboList");
                GC.Collect();

                ClearFilesInApplicationFolders(@"*.xlsx", "Excel-файлов");
                ClearFilesInApplicationFolders(@"*.log", "логов");

                _textBoxSetText(textBoxFIO, "");
                _textBoxSetText(textBoxGroup, "");
                _textBoxSetText(textBoxGroupDescription, "");
                _textBoxSetText(textBoxNav, "");

                _comboBoxClr(comboBoxFio);
                listFIO.Clear();

                iFIO = 0;
                TryMakeDB();
            }
            else
            { TryMakeDB(); }

            DataTable dt = new DataTable();
            _dataGridViewSource(dt);

            _toolStripStatusLabelSetText(StatusLabel2, @"Все таблицы очищены");
        }


        private void SelectPersonFromDataGrid(Person personSelected)
        {
            decimal[] timeIn = new decimal[4];
            decimal[] timeOut = new decimal[4];

            try
            {
                int IndexCurrentRow = _dataGridView1CurrentRowIndex();

                DataGridViewSeekValuesInSelectedRow dgSeek = new DataGridViewSeekValuesInSelectedRow();
                dgSeek.FindValueInCells(dataGridView1, new string[] { @"Фамилия Имя Отчество", @"NAV-код", @"Группа", @"Учетное время прихода ЧЧ:ММ", @"Учетное время ухода ЧЧ:ММ" });

                if (IndexCurrentRow > -1)
                {
                    if (nameOfLastTableFromDB == "PeopleGroup")
                    {
                        personSelected.FIO = dgSeek.values[0];
                        personSelected.NAV = dgSeek.values[1]; //Take the name of selected group
                        personSelected.GroupPerson = dgSeek.values[2]; //Take the name of selected group

                        personSelected.ControlInHHMM = dgSeek.values[3]; //Take the name of selected group
                        timeIn = ConvertStringTimeHHMMToDecimalArray(personSelected.ControlInHHMM);
                        personSelected.ControlInDecimal = timeIn[2];

                        personSelected.ControlOutHHMM = dgSeek.values[4]; //Take the name of selected group
                        timeOut = ConvertStringTimeHHMMToDecimalArray(personSelected.ControlOutHHMM);
                        personSelected.ControlOutDecimal = timeOut[2];

                        _numUpDownSet(numUpDownHourStart, timeIn[0]);
                        _numUpDownSet(numUpDownMinuteStart, timeIn[1]);

                        _numUpDownSet(numUpDownHourEnd, timeOut[0]);
                        _numUpDownSet(numUpDownMinuteEnd, timeOut[1]);

                        StatusLabel2.Text = @"Выбрана группа: " + personSelected.GroupPerson + @" | Курсор на: " + personSelected.FIO;

                        groupBoxPeriod.BackColor = Color.PaleGreen;
                    }
                    else if (nameOfLastTableFromDB == "PersonRegistrationsList")
                    {
                        personSelected.FIO = dgSeek.values[0];
                        personSelected.NAV = dgSeek.values[1]; //Take the name of selected group
                        personSelected.GroupPerson = _textBoxReturnText(textBoxGroup);

                        personSelected.ControlInHHMM = dgSeek.values[3]; //Take the name of selected group
                        timeIn = ConvertStringTimeHHMMToDecimalArray(personSelected.ControlInHHMM);
                        personSelected.ControlInDecimal = timeIn[2];

                        personSelected.ControlOutHHMM = dgSeek.values[4]; //Take the name of selected group
                        timeOut = ConvertStringTimeHHMMToDecimalArray(personSelected.ControlOutHHMM);
                        personSelected.ControlOutDecimal = timeOut[2];

                        _numUpDownSet(numUpDownHourStart, timeIn[0]);
                        _numUpDownSet(numUpDownMinuteStart, timeIn[1]);

                        _numUpDownSet(numUpDownHourEnd, timeOut[0]);
                        _numUpDownSet(numUpDownMinuteEnd, timeOut[1]);

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

                personSelected.ControlInHHMM = ConvertDecimalTimeToStringHHMM(_numUpDownReturn(numUpDownHourStart), _numUpDownReturn(numUpDownMinuteStart));
                personSelected.ControlOutHHMM = ConvertDecimalTimeToStringHHMM(_numUpDownReturn(numUpDownHourEnd), _numUpDownReturn(numUpDownMinuteEnd));
            } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
        }



        //---- Start. Drawing ---//

        private void VisualItem_Click(object sender, EventArgs e) //FindWorkDaysInSelected() , DrawFullWorkedPeriodRegistration()
        {
            Person personVisual = new Person();
            if (bLoaded)
            {
                SelectPersonFromDataGrid(personVisual);
                dataGridView1.Visible = false;
                //FindWorkDaysInSelected(dateTimePickerStart.Value.Year, dateTimePickerStart.Value.Month, dateTimePickerStart.Value.Day, dateTimePickerEnd.Value.Year, dateTimePickerEnd.Value.Month, dateTimePickerEnd.Value.Day);
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
                { hsNAV.Add(row[@"NAV-код"].ToString()); }
                string[] arrayNAVs = hsNAV.ToArray(); //unique NAV-codes
                int countNAVs = arrayNAVs.Count(); //the number of selected people
                numberPeopleInLoading = countNAVs;

                //count max number of events in-out all of selected people (the group or a single person)
                //It needs to prevent the error "index of scope"
                // SeekAnualDays();
                foreach (DataRow row in rowsPersonRegistrationsForDraw)
                {
                    for (int k = 0; k < workSelectedDays.Length; k++)
                    {
                        if (workSelectedDays[k].Length == 10 && row[@"Дата регистрации"].ToString().Contains(workSelectedDays[k]))  //если приход пришел в рабочий день - считаем
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
                                nav = row[@"NAV-код"].ToString();
                                dayRegistration = row[@"Дата регистрации"].ToString();

                                if (singleNav.Contains(nav) && dayRegistration.Contains(workDay))
                                {
                                    fio = row[@"Фамилия Имя Отчество"].ToString();
                                    minutesIn = (int)ConvertStringTimeHHMMToDecimalArray(row[@"Учетное время прихода ЧЧ:ММ"].ToString())[3];
                                    minutesInFact = (int)ConvertStringTimeHHMMToDecimalArray(row[@"Фактич. время прихода ЧЧ:ММ"].ToString())[3];
                                    minutesOut = (int)ConvertStringTimeHHMMToDecimalArray(row[@"Учетное время ухода ЧЧ:ММ"].ToString())[3];
                                    minutesOutFact = (int)ConvertStringTimeHHMMToDecimalArray(row[@"Фактич. время ухода ЧЧ:ММ"].ToString())[3];
                                    directionPass = row[@"Направление прохода"].ToString().ToLower();

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
            { hsNAV.Add(row[@"NAV-код"].ToString()); }
            string[] arrayNAVs = hsNAV.ToArray(); //unique NAV-codes
            int countNAVs = arrayNAVs.Count(); //the number of selected people
            numberPeopleInLoading = countNAVs;

            //count max number of events in-out all of selected people (the group or a single person)
            //It needs to prevent the error "index of scope"
            foreach (DataRow row in rowsPersonRegistrationsForDraw)
            {
                for (int k = 0; k < workSelectedDays.Length; k++)
                {
                    if (workSelectedDays[k].Length == 10 && row[@"Дата регистрации"].ToString().Contains(workSelectedDays[k]))  //если приход пришел в рабочий день - считаем
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
                            nav = row[@"NAV-код"].ToString();
                            dayRegistration = row[@"Дата регистрации"].ToString();

                            if (singleNav.Contains(nav) && dayRegistration.Contains(workDay))
                            {
                                fio = row[@"Фамилия Имя Отчество"].ToString();
                                minutesIn = (int)ConvertStringTimeHHMMToDecimalArray(row[@"Учетное время прихода ЧЧ:ММ"].ToString())[3];
                                minutesInFact = (int)ConvertStringTimeHHMMToDecimalArray(row[@"Фактич. время прихода ЧЧ:ММ"].ToString())[3];
                                minutesOut = (int)ConvertStringTimeHHMMToDecimalArray(row[@"Учетное время ухода ЧЧ:ММ"].ToString())[3];
                                minutesOutFact = (int)ConvertStringTimeHHMMToDecimalArray(row[@"Фактич. время ухода ЧЧ:ММ"].ToString())[3];
                                directionPass = row[@"Направление прохода"].ToString().ToLower();

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
                    "SELECT GroupPerson, GroupPersonDescription FROM PeopleGroupDesciption;", sqlConnection))
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

        private void MailingSave(string recipientEmail, string senderEmail, string groupsReport, string nameReport, string description, string period, string status, string localDate, string SendingDate)
        {
            bool recipientValid = false;
            bool senderValid = false;

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

                    using (SQLiteCommand sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'Mailing' (SenderEmail, RecipientEmail, GroupsReport, NameReport, Description, Period, Status, DateCreated, SendingLastDate)" +
                               " VALUES (@SenderEmail, @RecipientEmail, @GroupsReport, @NameReport, @Description, @Period, @Status, @DateCreated, @SendingLastDate)", sqlConnection))
                    {
                        sqlCommand.Parameters.Add("@SenderEmail", DbType.String).Value = senderEmail;
                        sqlCommand.Parameters.Add("@RecipientEmail", DbType.String).Value = recipientEmail;
                        sqlCommand.Parameters.Add("@GroupsReport", DbType.String).Value = groupsReport;
                        sqlCommand.Parameters.Add("@NameReport", DbType.String).Value = nameReport;
                        sqlCommand.Parameters.Add("@Description", DbType.String).Value = description;
                        sqlCommand.Parameters.Add("@Period", DbType.String).Value = period;
                        sqlCommand.Parameters.Add("@Status", DbType.String).Value = status;
                        sqlCommand.Parameters.Add("@DateCreated", DbType.String).Value = localDate;
                        sqlCommand.Parameters.Add("@SendingLastDate", DbType.String).Value = SendingDate;

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
                    "NameReport AS 'Наименование', Description AS 'Описание', Period AS 'Период', DateCreated AS 'Дата создания/модификации', SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус' ",
                    " ORDER BY RecipientEmail asc, DateCreated desc; ");
                nameOfLastTableFromDB = "Mailing";
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
            {
                PropertiesSave();
            }
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

                MailingSave(recipientEmail, senderEmail, report, nameReport, description, period, status, DateTime.Now.ToString(), "");
                ShowDataTableQuery(databasePerson, "Mailing", "SELECT SenderEmail AS 'Отправитель', RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', " +
                    "NameReport AS 'Наименование', Description AS 'Описание', Period AS 'Период', DateCreated AS 'Дата создания/модификации', SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус' ",
                    " ORDER BY RecipientEmail asc, DateCreated desc; ");
                nameOfLastTableFromDB = "Mailing";
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
                mysqlServerUserName = sMySqlServerUser;
                mysqlServerUserPassword = sMySqlServerUserPassword;

                try
                {
                    string modeApp = "";

                    if (currentModeAppManual)
                    {
                        modeApp = "1";
                    }
                    else if (!currentModeAppManual)
                    {
                        modeApp = "0";
                    }

                    using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(myRegKey))
                    {
                        try { EvUserKey.SetValue("SKDServer", server, Microsoft.Win32.RegistryValueKind.String); } catch { }
                        try { EvUserKey.SetValue("SKDUser", EncryptStringToBase64Text(user, btsMess1, btsMess2), Microsoft.Win32.RegistryValueKind.String); } catch { }
                        try { EvUserKey.SetValue("SKDUserPassword", EncryptStringToBase64Text(password, btsMess1, btsMess2), Microsoft.Win32.RegistryValueKind.String); } catch { }

                        try { EvUserKey.SetValue("MailServer", mailServer, Microsoft.Win32.RegistryValueKind.String); } catch { }
                        try { EvUserKey.SetValue("MailUser", mailServerUserName, Microsoft.Win32.RegistryValueKind.String); } catch { }
                        try { EvUserKey.SetValue("MailUserPassword", EncryptStringToBase64Text(mailServerUserPassword, btsMess1, btsMess2), Microsoft.Win32.RegistryValueKind.String); } catch { }

                        try { EvUserKey.SetValue("MySQLServer", mysqlServer, Microsoft.Win32.RegistryValueKind.String); } catch { }
                        try { EvUserKey.SetValue("MySQLUser", mysqlServerUserName, Microsoft.Win32.RegistryValueKind.String); } catch { }
                        try { EvUserKey.SetValue("MySQLUserPassword", EncryptStringToBase64Text(mysqlServerUserPassword, btsMess1, btsMess2), Microsoft.Win32.RegistryValueKind.String); } catch { }

                        logger.Info("Данные в реестре сохранены");
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
                    EvUserKey?.DeleteSubKey("MySQLServer");
                    EvUserKey?.DeleteSubKey("MySQLUser");
                    EvUserKey?.DeleteSubKey("MySQLUserPassword");
                    EvUserKey?.DeleteSubKey("ModeApp");
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
                nameOfLastTableFromDB = "PeopleGroup";
                isPerson = false;
            }
            else
            {
                _MenuItemTextSet(PersonOrGroupItem, "Работа с одной персоной");
                nameOfLastTableFromDB = "PersonRegistrationsList";
                isPerson = true;
            }
        }

        //--- End. Behaviour Controls ---//




        //---  Start.  DatagridView functions ---//

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e) //dataGridView1CellClick()
        { dataGridView1CellClick(); }

        private void dataGridView1CellClick()
        {
            DataGridViewSeekValuesInSelectedRow dgSeek;
            if (dataGridView1.Rows.Count > 0 && dataGridView1.CurrentRow.Index < dataGridView1.Rows.Count)
            {
                decimal[] timeIn = new decimal[4];
                decimal[] timeOut = new decimal[4];

                try
                {
                    int IndexCurrentRow = _dataGridView1CurrentRowIndex();
                    if (IndexCurrentRow > -1)
                    {
                        dgSeek = new DataGridViewSeekValuesInSelectedRow();

                        if (nameOfLastTableFromDB == "PeopleGroupDesciption")
                        {
                            dgSeek.FindValueInCells(dataGridView1, new string[] {
                             @"Группа",   @"Описание группы"
                            });

                            textBoxGroup.Text = dgSeek.values[0]; //Take the name of selected group
                            textBoxGroupDescription.Text = dgSeek.values[1]; //Take the name of selected group
                            groupBoxPeriod.BackColor = Color.PaleGreen;
                            groupBoxFilterReport.BackColor = SystemColors.Control;
                            StatusLabel2.Text = @"Выбрана группа: " + dgSeek.values[0] + @" |  Всего ФИО: " + iFIO;
                            if (textBoxFIO.TextLength > 3)
                            {
                                comboBoxFio.SelectedIndex = comboBoxFio.FindString(textBoxFIO.Text);
                            }
                        }
                        else if (nameOfLastTableFromDB == "ListFIO" || nameOfLastTableFromDB == "PeopleGroup" || nameOfLastTableFromDB == "PersonRegistrationsList")
                        {
                            dgSeek.FindValueInCells(dataGridView1, new string[] {
                             @"Группа", @"Фамилия Имя Отчество", @"NAV-код", @"Учетное время прихода ЧЧ:ММ", @"Учетное время ухода ЧЧ:ММ"
                            });

                            textBoxGroup.Text = dgSeek.values[0];
                            textBoxFIO.Text = dgSeek.values[1];
                            textBoxNav.Text = dgSeek.values[2];

                            StatusLabel2.Text = @"Выбрана группа: " + dgSeek.values[0] +
                                @" |Курсор на: " + ShortFIO(dgSeek.values[1]) + @" |Всего ФИО: " + iFIO;

                            groupBoxPeriod.BackColor = Color.PaleGreen;
                            groupBoxTimeStart.BackColor = Color.PaleGreen;
                            groupBoxTimeEnd.BackColor = Color.PaleGreen;
                            groupBoxFilterReport.BackColor = SystemColors.Control;
                            try
                            {
                                timeIn = ConvertStringTimeHHMMToDecimalArray(dgSeek.values[3]);
                                timeOut = ConvertStringTimeHHMMToDecimalArray(dgSeek.values[4]);
                                _numUpDownSet(numUpDownHourStart, timeIn[0]);
                                _numUpDownSet(numUpDownMinuteStart, timeIn[1]);
                                _numUpDownSet(numUpDownHourEnd, timeOut[0]);
                                _numUpDownSet(numUpDownMinuteEnd, timeOut[1]);
                            } catch { logger.Warn("dataGridView1CellClick:" + timeIn[0]); }

                            if (dgSeek.values[1].Length > 3)
                            { comboBoxFio.SelectedIndex = comboBoxFio.FindString(dgSeek.values[1]); }
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
            string group = "";

            if (_dataGridView1CurrentRowIndex() > -1)
            {
                try
                {
                    DataGridViewSeekValuesInSelectedRow dgSeek = new DataGridViewSeekValuesInSelectedRow();
                    if (nameOfLastTableFromDB == "PeopleGroup" || nameOfLastTableFromDB == "ListFIO")
                    {
                        dgSeek.FindValueInCells(dataGridView1, new string[] {
                            @"Фамилия Имя Отчество", @"NAV-код", @"Группа",
                            @"Учетное время прихода ЧЧ:ММ", @"Учетное время ухода ЧЧ:ММ",
                            @"Отдел", @"Должность", @"График", @"Комментарии"
                            });

                        fio = dgSeek.values[0];
                        textBoxFIO.Text = fio;

                        nav = dgSeek.values[1];
                        textBoxNav.Text = nav;

                        group = dgSeek.values[2];
                        textBoxGroup.Text = group;

                        using (var connection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                        {
                            connection.Open();
                            using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PeopleGroup' (FIO, NAV, GroupPerson, ControllingHHMM, ControllingOUTHHMM, Shift, Comment, Reserv1, Reserv2) " +
                                                                                        " VALUES (@FIO, @NAV, @GroupPerson, @ControllingHHMM, @ControllingOUTHHMM, @Shift, @Comment, @Reserv1, @Reserv2)", connection))
                            {
                                command.Parameters.Add("@FIO", DbType.String).Value = fio;
                                command.Parameters.Add("@NAV", DbType.String).Value = nav;

                                command.Parameters.Add("@GroupPerson", DbType.String).Value = group;
                                command.Parameters.Add("@Reserv1", DbType.String).Value = dgSeek.values[5];
                                command.Parameters.Add("@Reserv2", DbType.String).Value = dgSeek.values[6];

                                command.Parameters.Add("@ControllingHHMM", DbType.String).Value = dgSeek.values[3];
                                command.Parameters.Add("@ControllingOUTHHMM", DbType.String).Value = dgSeek.values[4];

                                command.Parameters.Add("@Shift", DbType.String).Value = dgSeek.values[7];
                                command.Parameters.Add("@Comment", DbType.String).Value = dgSeek.values[8];

                                try { command.ExecuteNonQuery(); } catch { }
                            }
                        }
                        SeekAndShowMembersOfGroup(group);
                        nameOfLastTableFromDB = "PeopleGroup";
                        StatusLabel2.Text = @"Обновлено время прихода " + ShortFIO(fio) + " в группе: " + group;
                    }
                    else if (nameOfLastTableFromDB == "PeopleGroupDesciption")
                    {
                        dgSeek.FindValueInCells(dataGridView1, new string[] {
                             @"Группа",   @"Описание группы"
                            });

                        textBoxGroup.Text = dgSeek.values[0]; //Take the name of selected group
                        textBoxGroupDescription.Text = dgSeek.values[1]; //Take the name of selected group
                        groupBoxPeriod.BackColor = Color.PaleGreen;
                        groupBoxFilterReport.BackColor = SystemColors.Control;
                        StatusLabel2.Text = @"Выбрана группа: " + dgSeek.values[0] + @" |  Всего ФИО: " + iFIO;
                    }
                    else if (nameOfLastTableFromDB == "Mailing")
                    {
                        dgSeek.FindValueInCells(dataGridView1, new string[] { @"Получатель", @"Отправитель", @"Отчет по группам", @"Наименование", @"Описание", @"Период", @"Статус" });
                        _toolStripStatusLabelSetText(StatusLabel2, "Готовлю отчет " + dgSeek.values[3]);
                        stimerPrev = "";

                        //todo 
                        //edit and update mailing

                        //   ExecuteSql("UPDATE 'Mailing' SET SendingLastDate='" + DateTime.Now.ToString() + "' WHERE RecipientEmail='" + dgSeek.values[0]
                        //  + "' AND NameReport='" + dgSeek.values[3] + "' AND GroupsReport ='" + dgSeek.values[2] + "';", databasePerson);


                        //  MailingAction("sendEmail", dgSeek.values[0], dgSeek.values[0], dgSeek.values[2], dgSeek.values[3], dgSeek.values[4], dgSeek.values[5], dgSeek.values[6]);

                        ShowDataTableQuery(databasePerson, "Mailing", "SELECT SenderEmail AS 'Отправитель', RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', " +
                            "NameReport AS 'Наименование', Description AS 'Описание', Period AS 'Период', DateCreated AS 'Дата создания/модификации', SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус' ",
                            " ORDER BY RecipientEmail asc, DateCreated desc; ");
                        nameOfLastTableFromDB = "Mailing";
                    }
                } catch { }
            }
        }

        //Show help to Edit on some columns DataGridView
        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            DataGridViewCell cell;
            if (nameOfLastTableFromDB == "PeopleGroup")
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
                    else if ((e.ColumnIndex == this.dataGridView1.Columns["Учетное время прихода ЧЧ:ММ"].Index) && e.Value != null)
                    {
                        cell = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                        cell.ToolTipText = "Для установки нового значения нажмите F2,\nвнесите новое значение,\nа затем нажмите Enter";
                    }
                    else if ((e.ColumnIndex == this.dataGridView1.Columns["Учетное время ухода ЧЧ:ММ"].Index) && e.Value != null)
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
                if ((nameOfLastTableFromDB == "PeopleGroupDesciption") && currentMouseOverRow > -1)
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
                case "PeopleGroupDesciption":
                    {
                        break;
                    }
                case "PeopleGroup" when textBoxGroup.Text.Trim().Length > 0:
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

                        logger.Info("DoMainAction: UPDATE 'Mailing' SET SendingLastDate='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "' WHERE RecipientEmail='" + dgSeek.values[0] + "' AND NameReport='" + dgSeek.values[3] + "' AND GroupsReport ='" + dgSeek.values[2] + "';");
                        ExecuteSql("UPDATE 'Mailing' SET SendingLastDate='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "' WHERE RecipientEmail='" + dgSeek.values[0]
                        + "' AND NameReport='" + dgSeek.values[3] + "' AND GroupsReport ='" + dgSeek.values[2] + "';", databasePerson);

                        logger.Info("DoMainAction, sendEmail: " + dgSeek.values[0] + "|" + dgSeek.values[1] + "|" + dgSeek.values[2] + "|" + dgSeek.values[3] + "|" + dgSeek.values[4] + "|" + dgSeek.values[5] + "|" + dgSeek.values[6]);
                        MailingAction("sendEmail", dgSeek.values[0], dgSeek.values[1], dgSeek.values[2], dgSeek.values[3], dgSeek.values[4], dgSeek.values[5], dgSeek.values[6]);

                        logger.Info("DoMainAction, ShowDataTableQuery: ");
                        ShowDataTableQuery(databasePerson, "Mailing", "SELECT SenderEmail AS 'Отправитель', RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', " +
                        "NameReport AS 'Наименование', Description AS 'Описание', Period AS 'Период', DateCreated AS 'Дата создания/модификации', SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус' ",
                        " ORDER BY RecipientEmail asc, DateCreated desc; ");
                        nameOfLastTableFromDB = "Mailing"; break;
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
            string group = _textBoxReturnText(textBoxGroup);

            if (_dataGridView1CurrentRowIndex() > -1)
            {
                switch (nameOfLastTableFromDB)
                {
                    case "PeopleGroupDesciption":
                        {
                            DataGridViewSeekValuesInSelectedRow dgSeek = new DataGridViewSeekValuesInSelectedRow();
                            dgSeek.FindValueInCells(dataGridView1, new string[] { @"Группа" });

                            DeleteDataTableQueryParameters(databasePerson, "PeopleGroup", "GroupPerson", dgSeek.values[0], "", "", "", "");
                            DeleteDataTableQueryParameters(databasePerson, "PeopleGroupDesciption", "GroupPerson", dgSeek.values[0], "", "", "", "");

                            PersonOrGroupItem.Text = "Работа с одной персоной";

                            ListGroups();
                            MembersGroupItem.Enabled = true;
                            _toolStripStatusLabelSetText(StatusLabel2, "Удалена группа: " + dgSeek.values[0] + "| Всего групп: " + _dataGridView1RowsCount());
                            break;
                        }
                    case "PeopleGroup" when group.Length > 0:
                        {
                            DeleteDataTableQueryParameters(databasePerson, "PeopleGroup", "GroupPerson", group, "", "", "", "");
                            DeleteDataTableQueryParameters(databasePerson, "PeopleGroupDesciption", "GroupPerson", group, "", "", "", "");
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
                                "NameReport AS 'Наименование', Description AS 'Описание', Period AS 'Период', DateCreated AS 'Дата создания/модификации', SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус' ",
                                " ORDER BY RecipientEmail asc, DateCreated desc; ");
                            nameOfLastTableFromDB = "Mailing"; break;
                        }
                    default:
                        break;
                }
            }
            group = null;
        }

        //---  End.  DatagridView functions ---//



        public string GetSafeFilename(string filename)
        {
            return string.Join("_", filename.Split(System.IO.Path.GetInvalidFileNameChars()));
        }

        private string selectPeriodMonth(bool current = false) //format of result: firstDay + "|" + lastDay
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


        //---  Start. Schedule Functions ---//
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

        //todo записать режим в реестр

        private void ModeAppItem_Click(object sender, EventArgs e)  //ModeApp()
        { SwitchAppMode(); }

        private void SwitchAppMode()       // ExecuteAutoMode()
        {
            if (currentModeAppManual)
            {
                _MenuItemTextSet(modeItem, "Сменить режим на интерактивный");
                _menuItemTooltipSet(modeItem, "Включен автоматический режим. Выполняются Активные рассылки из БД.");
                _MenuItemBackColorChange(modeItem, Color.DarkOrange);

                _toolStripStatusLabelSetText(StatusLabel2, "Включен режим авторассылки отчетов");
                _toolStripStatusLabelBackColor(StatusLabel2, Color.PaleGreen);

                try
                {
                    using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(myRegKey))
                    {
                        EvUserKey.SetValue("ModeApp", "0", Microsoft.Win32.RegistryValueKind.String);
                        logger.Info("Save ModeApp in Registry. Данные в реестре сохранены");
                    }

                    using (Microsoft.Win32.RegistryKey EvUserKey =
                         Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run\"))
                    {
                        EvUserKey.SetValue(productName, "\"" + Application.ExecutablePath + "\"");
                        logger.Info("Save AutoRun App in Registry. Данные в реестре сохранены");
                    }
                } catch (Exception expt) { logger.Warn("Save ModeApp,AutoRun in Registry. Последний режим работы не сохранен. " + expt); }
                ExecuteAutoMode(true);
            }
            else
            {
                _MenuItemTextSet(modeItem, "Сменить на автоматический режим");
                _menuItemTooltipSet(modeItem, "Включен интерактивный режим. Все рассылки остановлены.");
                _MenuItemBackColorChange(modeItem, SystemColors.Control);

                _toolStripStatusLabelSetText(StatusLabel2, "Интерактивный режим");
                _toolStripStatusLabelBackColor(StatusLabel2, SystemColors.Control);

                try
                {
                    using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(myRegKey))
                    {
                        EvUserKey.SetValue("ModeApp", "1", Microsoft.Win32.RegistryValueKind.String);
                        logger.Info("Change ModeApp in Registry. Данные в реестре сохранены");
                    }

                    using (Microsoft.Win32.RegistryKey EvUserKey =
                        Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", true))
                    {
                        EvUserKey.DeleteValue(productName);
                        logger.Info("Delete AutoRun App from Registry. Ключ удален");
                    }
                } catch (Exception expt) { logger.Warn("Delete ModeApp from Registry. Ошибка удаления ключа. " + expt); }
                ExecuteAutoMode(false);
            }

            //trigger mode
            if (currentModeAppManual)
            { currentModeAppManual = false; }
            else
            { currentModeAppManual = true; }
        }

        private async void ExecuteAutoMode(bool manualMode) //InitScheduleTask()
        {
            await Task.Run(() => InitScheduleTask(manualMode));
        }

        public void InitScheduleTask(bool manualMode) //ScheduleTask()
        {
            long interval = 1 * 60 * 1000; //1 minute
            if (manualMode)
            {
                _MenuItemTextSet(modeItem, "Сменить режим на интерактивный");
                _menuItemTooltipSet(modeItem, "Включен автоматический режим. Выполняются Активные рассылки из БД.");
                _MenuItemBackColorChange(modeItem, Color.DarkOrange);

                _toolStripStatusLabelSetText(StatusLabel2, "Включен режим авторассылки отчетов");
                _toolStripStatusLabelBackColor(StatusLabel2, Color.PaleGreen);

                currentAction = "sendEmail";
                timer = new System.Threading.Timer(new System.Threading.TimerCallback(ScheduleTask), null, 0, interval);
            }
            else
            {
                _MenuItemTextSet(modeItem, "Сменить на автоматический режим");
                _menuItemTooltipSet(modeItem, "Включен интерактивный режим. Все рассылки остановлены.");
                _MenuItemBackColorChange(modeItem, SystemColors.Control);

                _toolStripStatusLabelSetText(StatusLabel2, "Интерактивный режим");
                _toolStripStatusLabelBackColor(StatusLabel2, SystemColors.Control);

                timer?.Dispose();
            }
        }

        static System.Threading.Timer timer;
        static object synclock = new object();
        static bool sent = false;

        private void ScheduleTask(object obj) //SelectMailingDoAction()
        {
            lock (synclock)
            {
                DateTime dd = DateTime.Now;
                if (dd.Day == 1 && dd.Hour == 2 && dd.Minute == 5 && sent == false) //do something at Hour 3 and 5 minute //dd.Day == 1 && 
                {
                    _toolStripStatusLabelSetText(StatusLabel2, "Ведется работа по подготовке отчетов " + DateTime.Now.ToString("yyyy-MM-dd HH:mm"));
                    _toolStripStatusLabelBackColor(StatusLabel2, Color.LightPink);
                    SelectMailingDoAction();
                    sent = true;
                }
                else if (dd.Minute != 5)
                {
                    sent = false;
                    _toolStripStatusLabelSetText(StatusLabel2, "Режим почтовых рассылок. " + DateTime.Now.ToString("yyyy-MM-dd HH:mm"));
                    _toolStripStatusLabelBackColor(StatusLabel2, Color.LightCyan);
                    ClearFilesInApplicationFolders(@"*.xlsx", "Excel-файлов");
                }
            }
        }

        private void SelectMailingDoAction() //MailingAction()
        {
            _ProgressBar1Start();

            //предварительно обновить список ФИО и индивидуальные графики
            // commented only for test period!!!
            GetFIO();

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

                logger.Info("UPDATE 'Mailing' SET SendingLastDate='" + DateTime.Now.ToString() + "' WHERE RecipientEmail='" + mailng._recipient
                + "' AND NameReport='" + mailng._nameReport + "' AND GroupsReport ='" + mailng._groupsReport + "';");
                ExecuteSql("UPDATE 'Mailing' SET SendingLastDate='" + DateTime.Now.ToString() + "' WHERE RecipientEmail='" + mailng._recipient
                + "' AND NameReport='" + mailng._nameReport + "' AND GroupsReport ='" + mailng._groupsReport + "';", databasePerson);

                logger.Info("sendEmail" + "|" + mailng._recipient + "|" + mailng._sender + "|" + mailng._groupsReport + "|" + mailng._nameReport + "|" + mailng._descriptionReport + "|" + mailng._period + "|" + mailng._status);
                MailingAction("sendEmail", mailng._recipient, mailng._sender, mailng._groupsReport, mailng._nameReport, mailng._descriptionReport, mailng._period, mailng._status);

                ShowDataTableQuery(databasePerson, "Mailing", "SELECT SenderEmail AS 'Отправитель', RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', " +
                "NameReport AS 'Наименование', Description AS 'Описание', Period AS 'Период', DateCreated AS 'Дата создания/модификации', SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус' ",
                " ORDER BY RecipientEmail asc, DateCreated desc; ");
                nameOfLastTableFromDB = "Mailing";
            }

            sender = null; recipient = null; gproupsReport = null; nameReport = null; descriptionReport = null; period = null; status = null;
            mailingList = null;

            _ProgressBar1Stop();
        }

        private void MailingAction(string mainAction, string recipientEmail, string senderEmail, string groupsReport, string nameReport, string description, string period, string status)
        {
            _toolStripStatusLabelBackColor(StatusLabel2, SystemColors.Control);

            switch (mainAction)
            {
                case "saveEmail":
                    {
                        MailingSave(recipientEmail, senderEmail, groupsReport, nameReport, description, period, status, DateTime.Now.ToString(), "");

                        ShowDataTableQuery(databasePerson, "Mailing", "SELECT SenderEmail AS 'Отправитель', RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', " +
                            "NameReport AS 'Наименование', Description AS 'Описание', Period AS 'Период', DateCreated AS 'Дата создания/модификации', SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус' ",
                            " ORDER BY RecipientEmail asc, DateCreated desc; ");
                        nameOfLastTableFromDB = "Mailing";
                        break;
                    }
                case "sendEmail":
                    {
                        //Check making report
                        CheckAliveServer(sServer1, sServer1UserName, sServer1UserPassword);

                        if (bServer1Exist)
                        {
                            DataTable dtTempIntermediate = dtPeople.Clone();
                            Person person = new Person();
                            string fio, nav, group, dep, pos, timein, timeout, comment, shift;

                            GetNamePoints();  //Get names of the registration' points

                            logger.Info("MailingAction, sendEmail: " + selectPeriodMonth());
                            reportStartDay = selectPeriodMonth().Split('|')[0];
                            reportLastDay = selectPeriodMonth().Split('|')[1];

                            //periodComboParameters.Add("Текущий месяц");
                            //periodComboParameters.Add("Предыдущий месяц");
                            if (period.ToLower().Contains("текущий") || period.ToLower().Contains("текущая"))
                            {
                                reportStartDay = selectPeriodMonth(true).Split('|')[0];
                                reportLastDay = selectPeriodMonth(true).Split('|')[1];
                            }
                            logger.Info("MailingAction, startDay, lastDay: " + reportStartDay + " " + reportLastDay);

                            string nameGroup = "";
                            string selectedPeriod = reportStartDay.Split(' ')[0] + " - " + reportLastDay.Split(' ')[0];
                            string bodyOfMail = "";
                            string titleOfbodyMail = "";
                            string[] groups = groupsReport.Split('+');

                            foreach (string name in groups)
                            {
                                try { System.IO.File.Delete(filePathExcelReport); } catch { }

                                nameGroup = name.Trim();
                                if (nameGroup.Length > 0)
                                {
                                    dtPersonRegistrationsFullList.Clear();
                                    GetRegistrations(name, reportStartDay, reportLastDay, "sendEmail");//typeReport== only one group
                                    logger.Info("sendEmail:" + "dtPeopleGroup.Rows.Count - " + dtPeopleGroup.Rows.Count);
                                    logger.Info("sendEmail:" + "dtPersonRegistrationsFullList.Rows.Count - " + dtPersonRegistrationsFullList.Rows.Count);

                                    dtTempIntermediate?.Dispose();
                                    dtTempIntermediate = dtPeople.Clone();

                                    person = new Person();

                                    dtPersonTemp?.Dispose();
                                    dtPersonTemp = dtPeople.Clone();

                                    LoadGroupMembersFromDbToDataTable(nameGroup); //result will be in dtPeopleGroup  //"Select * FROM PeopleGroup where GroupPerson like '" + _textBoxReturnText(textBoxGroup) + "';"

                                    foreach (DataRow row in dtPeopleGroup.Rows)
                                    {
                                        if (row[@"Фамилия Имя Отчество"]?.ToString().Length > 0 && row[@"Группа"]?.ToString() == nameGroup)
                                        {
                                            person = new Person();

                                            fio = row[@"Фамилия Имя Отчество"].ToString();
                                            nav = row[@"NAV-код"].ToString();

                                            group = row[@"Группа"].ToString();
                                            dep = row[@"Отдел"].ToString();
                                            pos = row[@"Должность"].ToString();

                                            timein = row[@"Учетное время прихода ЧЧ:ММ"].ToString();
                                            timeout = row[@"Учетное время ухода ЧЧ:ММ"].ToString();

                                            comment = row[@"Комментарии"].ToString();
                                            shift = row[@"График"].ToString();

                                            person.FIO = fio;
                                            person.NAV = nav;

                                            person.GroupPerson = group;
                                            person.Department = dep;
                                            person.PositionInDepartment = pos;

                                            person.ControlInHHMM = timein;
                                            person.ControlOutHHMM = timeout;

                                            person.ControlInDecimal = ConvertStringTimeHHMMToDecimalArray(timein)[2];
                                            person.ControlOutDecimal = ConvertStringTimeHHMMToDecimalArray(timeout)[2];

                                            person.Comment = comment;
                                            person.Shift = shift;

                                            FilterDataByNav(person, dtPersonRegistrationsFullList, dtTempIntermediate);
                                        }
                                    }
                                    person = null;
                                    logger.Info("dtTempIntermediate: " + dtTempIntermediate.Rows.Count);
                                    dtPersonTemp = GetDistinctRecords(dtTempIntermediate, orderColumnsFinacialReport);
                                    dtPersonTemp.SetColumnsOrder(orderColumnsFinacialReport);
                                    logger.Info("dtPersonTemp: " + dtPersonTemp.Rows.Count);

                                    if (dtPersonTemp.Rows.Count > 0)
                                    {
                                        string nameFile = (nameReport + " " + reportStartDay.Split(' ')[0] + "-" + reportLastDay.Split(' ')[0] + " " + name + " от " + DateTime.Now.ToString("yy-MM-dd HH-mm"));
                                        string illegal = GetSafeFilename(nameFile) + @".xlsx";

                                        filePathExcelReport = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(filePathApplication), illegal);
                                        try { System.IO.File.Delete(filePathExcelReport); } catch (Exception expt)
                                        { logger.Error("Ошибка удаления файла " + filePathExcelReport + " " + expt.ToString()); }

                                        logger.Info("сохраняю файл: " + filePathExcelReport);
                                        ExportDatatableSelectedColumnsToExcel(dtPersonTemp, nameReport, filePathExcelReport);

                                        if (reportExcelReady)
                                        {
                                            bodyOfMail = "Отчет \"" + nameReport + "\" " + " по группе \"" + nameGroup + "\"";
                                            titleOfbodyMail = "Отчет за период: с " + reportStartDay.Split(' ')[0] + " по " + reportLastDay.Split(' ')[0];
                                            _toolStripStatusLabelSetText(StatusLabel2, "Подготавливаю отчет для отправки " + recipientEmail);

                                            SendEmailAsync(senderEmail, recipientEmail, titleOfbodyMail, bodyOfMail, filePathExcelReport, Properties.Resources.LogoRYIK, productName);

                                            _toolStripStatusLabelBackColor(StatusLabel2, Color.PaleGreen);
                                            _toolStripStatusLabelSetText(StatusLabel2, DateTime.Now.ToString("yyyy-MM-dd HH:mm") + " Отчет " + nameReport + "(" + name + ") подготовлен и отправлен " + recipientEmail);
                                        }
                                        else
                                        {
                                            _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                                            _toolStripStatusLabelSetText(StatusLabel2, DateTime.Now.ToString("yyyy-MM-dd HH:mm") + " Ошибка создания отчета: " + nameReport + "(" + name + ")");
                                        }
                                    }
                                    else
                                    {
                                        _toolStripStatusLabelSetText(StatusLabel2, "Ошибка получения данных по отчету " + nameReport);
                                        _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                                    }
                                    //destroy temporary variables
                                    person = null;
                                }
                            }
                            //destroy temporary variables
                            dtTempIntermediate?.Dispose();
                            dtPersonTemp?.Clear();
                            selectedPeriod = null;
                            bodyOfMail = null; titleOfbodyMail = null; nameGroup = null;
                        }
                        break;
                    }
                default:
                    break;
            }
        }

        private async void SendEmailAsync(string sender, string recipient, string title, string messageBeforePicture, string pathToFile, Bitmap myLogo, string messageAfterPicture) //Compose and send e-mail
        {
            // string startupPath = AppDomain.CurrentDomain.RelativeSearchPath;
            // string path = System.IO.Path.Combine(startupPath, "HtmlTemplates", "NotifyTemplate.html");
            // string body = System.IO.File.ReadAllText(path);

            // адрес smtp-сервера и порт, с которого будем отправлять письмо
            using (System.Net.Mail.SmtpClient smtpClient = new System.Net.Mail.SmtpClient(mailServer, 587))
            {
                smtpClient.EnableSsl = false; // I get error with "true"
                smtpClient.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                smtpClient.UseDefaultCredentials = false;
                smtpClient.Timeout = 50000;
                // создаем объект сообщения
                using (System.Net.Mail.MailMessage newMail = new System.Net.Mail.MailMessage())
                {
                    // письмо представляет код html
                    newMail.IsBodyHtml = true;
                    newMail.BodyEncoding = UTF8Encoding.UTF8;

                    newMail.AlternateViews.Add(getEmbeddedImage(title, messageBeforePicture, myLogo, messageAfterPicture));
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

        private System.Net.Mail.AlternateView getEmbeddedImage(string title, string messageBeforePicture, Bitmap bmp, string messageAfterPicture)
        {
            //convert embedded resources into memorystream
            Bitmap b = new Bitmap(bmp);
            ImageConverter ic = new ImageConverter();
            Byte[] ba = (Byte[])ic.ConvertTo(b, typeof(Byte[]));
            System.IO.MemoryStream logo = new System.IO.MemoryStream(ba);

            System.Net.Mail.LinkedResource res = new System.Net.Mail.LinkedResource(logo, "image/jpeg");
            res.ContentId = Guid.NewGuid().ToString();

            // текст письма
            string htmlBody =
                @"<p><font size='4' color='black' face='Arial'>" + title + @"</font></p>" +
                @"<p><h4>" + messageBeforePicture + @"</h4></p>" +
                @"<hr><img src='cid:" + res.ContentId + @"'/>" +
                @"<hr><p><h5>Отчет сгенерирован на " + pcname +
                @" с использованием ПО - <font size='2' color='blue' face='Arial'>" +
                myFileVersionInfo.ProductName + " " + myFileVersionInfo.FileVersion + @"</font> " + @" в автоматическом режиме <font size='2' color='red' face='Arial'>" +
                DateTime.Now.ToString("yyyy-MM-dd в HH:mm") + @"</font></h5></p>";

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
                        dataGridView1.Rows.Clear();
                        dataGridView1.Columns.Clear();
                        System.Collections.ArrayList Empty = new System.Collections.ArrayList();
                        dataGridView1.DataSource = Empty;
                        dataGridView1.Refresh();
                    }
                }));
            }
            else
            {
                if (dt != null && dt.Rows.Count > 0)
                { dataGridView1.DataSource = dt; }
                else
                {
                    dataGridView1.Rows.Clear();
                    dataGridView1.Columns.Clear();
                    System.Collections.ArrayList Empty = new System.Collections.ArrayList();
                    dataGridView1.DataSource = Empty;
                    dataGridView1.Refresh();
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
                Invoke(new MethodInvoker(delegate
                {
                    comboBx.Items.Clear();
                    comboBx.SelectedText = "";
                    comboBx.Text = "";
                }));
            else
            {
                comboBx.Items.Clear();
                comboBx.SelectedText = "";
                comboBx.Text = "";
            }
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
                    stringDT = dateTimePickerStart.Value.Year.ToString("0000") + "-" + dateTimePickerStart.Value.Month.ToString("00") + "-" + dateTimePickerStart.Value.Day.ToString("00") + " 00:00:00";
                }));
            else
            {
                stringDT = dateTimePickerStart.Value.Year.ToString("0000") + "-" + dateTimePickerStart.Value.Month.ToString("00") + "-" + dateTimePickerStart.Value.Day.ToString("00") + " 00:00:00";
            }
            return stringDT;
        }

        private string _dateTimePickerEnd() //add string into  from other threads
        {
            string stringDT = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    stringDT = dateTimePickerEnd.Value.Year.ToString("0000") + "-" + dateTimePickerEnd.Value.Month.ToString("00") + "-" + dateTimePickerEnd.Value.Day.ToString("00") + " 23:59:59";
                }));
            else
            {
                stringDT = dateTimePickerEnd.Value.Year.ToString("0000") + "-" + dateTimePickerEnd.Value.Month.ToString("00") + "-" + dateTimePickerEnd.Value.Day.ToString("00") + " 23:59:59";
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

        private int[] _dateTimePickerReturnArray(DateTimePicker dateTimePicker) //add string into  from other threads
        {
            int[] result = new int[3];

            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    result[0] = dateTimePicker.Value.Year;
                    result[1] = dateTimePicker.Value.Month;
                    result[2] = dateTimePicker.Value.Day;
                }
                ));
            else
            {
                result[0] = dateTimePicker.Value.Year;
                result[1] = dateTimePicker.Value.Month;
                result[2] = dateTimePicker.Value.Day;
            }
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

        /*
        string stime1 = "18:40";
        string stime2 = "19:35";
        DateTime t1 = DateTime.Parse(stime1);
        DateTime t2 = DateTime.Parse(stime2);
        TimeSpan ts = t2 - t1;
        int minutes = (int)ts.TotalMinutes;
        int seconds = (int)ts.TotalSeconds;
        */

        private decimal ConvertDecimalSeparatedTimeToDecimal(decimal decimalHour, decimal decimalMinute)
        {
            decimal result = decimalHour + TryParseStringToDecimal(TimeSpan.FromMinutes((double)decimalMinute).TotalHours.ToString());
            return result;
        }

        private decimal ConvertStringsTimeToDecimal(string hour, string minute)
        {
            decimal result = TryParseStringToDecimal(hour) + TryParseStringToDecimal(TimeSpan.FromMinutes(TryParseStringToDouble(minute)).TotalHours.ToString());
            return result;
        }

        private string[] ConvertDecimalTimeToStringHHMMArray(decimal decimalTime)
        {
            string[] result = new string[3];
            int hour = (int)(decimalTime);
            int minute = Convert.ToInt32(60 * (decimalTime - hour));

            result[0] = string.Format("{0:d2}", hour);
            result[1] = string.Format("{0:d2}", minute);
            result[2] = string.Format("{0:d2}:{1:d2}", hour, minute);
            return result;
        }

        private string ConvertDecimalTimeToStringHHMM(decimal hours, decimal minutes)
        {
            string result = string.Format("{0:d2}:{1:d2}", (int)hours, (int)minutes);
            return result;
        }



        private string ConvertDecimalTimeToStringHHMM(decimal decimalTime)
        {
            string result;
            int hour = (int)(decimalTime);
            int minute = Convert.ToInt32(60 * (decimalTime - hour));
            result = string.Format("{0:d2}:{1:d2}", hour, minute);
            return result;
        }

        private string[] ConvertSecondsTimeToStringHHMMArray(int seconds)
        {
            string[] result = new string[3];
            var ts = TimeSpan.FromSeconds(seconds);
            result[0] = String.Format("{0:d2}", (int)ts.TotalHours);
            result[1] = String.Format("{0:d2}", (int)ts.Minutes);
            result[2] = String.Format("{0:d2}:{1:d2}", (int)ts.TotalHours, (int)ts.Minutes);

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

        private decimal[] ConvertStringTimeHHMMToDecimalArray(string timeInHHMM) //time HH:MM converted to decimal value
        {
            decimal[] result = new decimal[4];
            string hour = "0";
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

        private string[] ConvertStringTimeHHMMToStringArray(string timeInHHMM) //time HH:MM converted to decimal value
        {
            string[] result = new string[4];
            decimal h = 0;
            decimal m = 0;

            if (timeInHHMM.Contains(':'))
            {
                string[] time = timeInHHMM.Split(':');
                h = TryParseStringToDecimal(time[0]);
                m = TryParseStringToDecimal(time[1]);
            }
            else
            {
                h = TryParseStringToDecimal(timeInHHMM);
                m = 0;
            }

            result[0] = String.Format("{0:d2}", h);                             // only hours        22
            result[1] = String.Format("{0:d2}", m);                             // only minutes        25
            result[2] = String.Format("{0:d2}:{1:d2}", h, m);                   // normalyze to     22:25

            return result;
        }

        private int[] ConvertStringDateToIntArray(string dateYYYYmmDD) //date YYYY-MM-DD to int array values
        {
            int[] result = new int[] { 1970, 1, 1 };

            if (dateYYYYmmDD.Contains('-'))
            {
                string[] res = dateYYYYmmDD.Split(' ')[0]?.Trim()?.Split('-');
                result[0] = Convert.ToInt32(res[0]);
                result[1] = Convert.ToInt32(res[1]);
                result[2] = Convert.ToInt32(res[2]);
            }
            else if (dateYYYYmmDD.Length == 8)
            {
                result[0] = Convert.ToInt32(dateYYYYmmDD.Remove(4));
                result[1] = Convert.ToInt32((dateYYYYmmDD.Remove(0, 2)).Remove(2));
                result[2] = Convert.ToInt32(dateYYYYmmDD.Remove(0, 5));
            }

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



    }


}
