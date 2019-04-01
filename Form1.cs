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
using System.Windows;

using System.Security.Cryptography;  // for Crypography

//using NLog;
//Project\Control NuGet\console 
//install-package nlog
//install-package nlog.config


namespace ASTA
{
    public partial class WinFormASTA :Form
    {
        static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        NotifyIcon notifyIcon = new NotifyIcon();
        ContextMenu contextMenu;
        bool buttonAboutForm;

        readonly string myRegKey = @"SOFTWARE\RYIK\ASTA";
        readonly System.IO.FileInfo databasePerson = new System.IO.FileInfo(@".\main.db");
        readonly byte[] btsMess1 = Convert.FromBase64String(@"OCvesunvXXsxtt381jr7vp3+UCwDbE4ebdiL1uinGi0="); //Key Encrypt
        readonly byte[] btsMess2 = Convert.FromBase64String(@"NO6GC6Zjl934Eh8MAJWuKQ=="); //Key Decrypt

        static string[] allParametersOfConfig = new string[] {
                "SKDServer" , "SKDUser", "SKDUserPassword",
                "MySQLServer", "MySQLUser", "MySQLUserPassword",
                "MailServer","MailUser","MailUserPassword"
            };

        string nameOfLastTableFromDB = "PersonRegistrationsList";
        string currentAction = "";
        bool currentModeAppManual = true;
        System.Diagnostics.FileVersionInfo myFileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath);
        string guid = System.Runtime.InteropServices.Marshal.GetTypeLibGuidForAssembly(System.Reflection.Assembly.GetExecutingAssembly()).ToString(); // получаем GIUD приложения
        string pcname = Environment.MachineName + "|" + Environment.OSVersion;
        string poname = "";
        string poversion = "";
        string productName = "";
        string strVersion;


        int iCounterLine = 0;

        //collecting of data
        static List<PassByPoint> passByPoints = new List<PassByPoint>(); //List byPass points 
        static List<Person> listFIO = new List<Person>(); // List of FIO and identity of data

        //Controls "NumUpDown"
        decimal numUpHourStart = 9;
        decimal numUpMinuteStart = 0;
        decimal numUpHourEnd = 18;
        decimal numUpMinuteEnd = 0;

        //Shows visual of registration
        PictureBox pictureBox1 = new PictureBox();
        Bitmap bmp = new Bitmap(1, 1);

        //making reports
        const int offsetTimeIn = 60;    // смещение времени прихода после учетного, в секундах, в течении которого не учитывается опоздание
        const int offsetTimeOut = 60;   // смещение времени ухода до учетного, в секундах в течении которого не учитывается ранний уход
        string[] myBoldedDates;
        string[] workSelectedDays;
        static string reportStartDay = "";
        static string reportLastDay = "";
        bool reportExcelReady = false;
        string filePathApplication = Application.ExecutablePath;
        string filePathExcelReport;

        //mailing
        const string NAME_OF_SENDER_REPORTS = "Отдел компенсаций и льгот";
        static System.Threading.Timer timer;
        static object synclock = new object();
        static bool sent = false;
        static string DEFAULT_RECEIVING_PORT_MAILSERVER = "587";
        static string DEFAULT_DAY_OF_SENDING_REPORT = "28";

        //Page of Mailing
        Label labelMailServerName;
        TextBox textBoxMailServerName;
        Label labelMailServerUserName;
        TextBox textBoxMailServerUserName;
        Label labelMailServerUserPassword;
        TextBox textBoxMailServerUserPassword;
        static string mailServer = "";
        static string mailServerRegistry = "";
        static string mailServerDB = "";
        static string mailServerUserName = "";
        static string mailServerUserNameRegistry = "";
        static string mailServerUserNameDB = "";
        static string mailServerUserPassword = "";
        static string mailServerUserPasswordRegistry = "";
        static string mailServerUserPasswordDB = "";

        //Page of "Settings' Programm"
        bool bServer1Exist = false;
        Label labelServer1;
        TextBox textBoxServer1;
        Label labelServer1UserName;
        TextBox textBoxServer1UserName;
        Label labelServer1UserPassword;
        TextBox textBoxServer1UserPassword;
        string sServer1 = "";
        string sServer1Registry = "";
        string sServer1DB = "";
        string sServer1UserName = "";
        string sServer1UserNameRegistry = "";
        string sServer1UserNameDB = "";
        string sServer1UserPassword = "";
        string sServer1UserPasswordRegistry = "";
        string sServer1UserPasswordDB = "";

        Label labelmysqlServer;
        TextBox textBoxmysqlServer;
        Label labelmysqlServerUserName;
        TextBox textBoxmysqlServerUserName;
        Label labelmysqlServerUserPassword;
        TextBox textBoxmysqlServerUserPassword;
        static string mysqlServer = "";
        static string mysqlServerRegistry = "";
        static string mysqlServerDB = "";
        static string mysqlServerUserName = "";
        static string mysqlServerUserNameRegistry = "";
        static string mysqlServerUserNameDB = "";
        static string mysqlServerUserPassword = "";
        static string mysqlServerUserPasswordRegistry = "";
        static string mysqlServerUserPasswordDB = "";

        Label listComboLabel;
        ComboBox listCombo;

        Label periodComboLabel;
        ListBox periodCombo;

        Label labelSettings9;
        ComboBox comboSettings9;

        Label labelSettings15; //type report
        ComboBox comboSettings15;

        Label labelSettings16;
        TextBox textBoxSettings16;

        CheckBox checkBox1;
        static List<ParameterConfig> listParameters = new List<ParameterConfig>();

        static Color clrRealRegistration = Color.PaleGreen;
        Color clrRealRegistrationRegistry = Color.PaleGreen;
        string sLastSelectedElement = "MainForm";

        OpenFileDialog openFileDialog1 = new OpenFileDialog();
        List<string> listGroups = new List<string>();

        int numberPeopleInLoading = 1;
        string stimerPrev = "";
        string stimerCurr = "Ждите!";

        //Names of collumns
        const string NPP = @"№ п/п";
        const string FIO = @"Фамилия Имя Отчество";
        const string CODE = @"NAV-код";
        const string GROUP = @"Группа";
        const string GROUP_DECRIPTION = @"Описание группы";
        const string TIMEIN = @"Время прихода";
        const string TIMEOUT = @"Время ухода";
        const string N_ID = @"№ пропуска";
        const string DEPARTMENT = @"Отдел";
        const string DEPARTMENT_ID = @"Отдел (id)";
        const string PLACE_EMPLOYEE = @"Местонахождение сотрудника";
        const string DATE_REGISTRATION = @"Дата регистрации";
        const string TIME_REGISTRATION = @"Время регистрации";
        const string SERVER_SKD = @"Сервер СКД";
        const string NAME_CHECKPOINT = @"Имя точки прохода";
        const string DIRECTION_WAY = @"Направление прохода";
        const string DESIRED_TIME_IN = @"Учетное время прихода ЧЧ:ММ";
        const string DESIRED_TIME_OUT = @"Учетное время ухода ЧЧ:ММ";
        const string REAL_TIME_IN = @"Фактич. время прихода ЧЧ:ММ";
        const string REAL_TIME_OUT = @"Фактич. время ухода ЧЧ:ММ";
        const string DAY_OF_WEEK = @"День недели";
        const string CHIEF_ID = @"Руководитель (код)";
        const string EMPLOYEE_POSITION = @"Должность";
        const string EMPLOYEE_SHIFT = @"График";
        const string EMPLOYEE_SHIFT_COMMENT = @"Комментарии (командировка, на выезде, согласованное отсутствие…….)";
        const string EMPLOYEE_HOOKY = @"Прогул (отпуск за свой счет)";
        const string EMPLOYEE_ABSENCE = @"Отсутствовал на работе";
        const string EMPLOYEE_SICK_LEAVE = @"Больничный";
        const string EMPLOYEE_TRIP = @"Командировка";
        const string EMPLOYEE_VACATION = @"Отпуск";
        const string EMPLOYEE_BEING_LATE = @"Опоздание ЧЧ:ММ";
        const string EMPLOYEE_EARLY_DEPARTURE = @"Ранний уход ЧЧ:ММ";
        const string EMPLOYEE_PLAN_TIME_WORKED = @"Отработанное время ЧЧ:ММ";
        const string EMPLOYEE_TIME_SPENT = @"Реальное отработанное время";

        //DataTables with people data
        DataTable dtPeople = new DataTable("People");
        DataColumn[] dcPeople =
           {
                                  new DataColumn(NPP,typeof(int)),//0
                                  new DataColumn(FIO,typeof(string)),//1
                                  new DataColumn(CODE,typeof(string)),//2
                                  new DataColumn(GROUP,typeof(string)),//3
                                  new DataColumn(TIMEIN,typeof(int)),//6
                                  new DataColumn(TIMEOUT,typeof(int)),//9
                                  new DataColumn(N_ID,typeof(int)), //10
                                  new DataColumn(DEPARTMENT,typeof(string)),//11
                                  new DataColumn(PLACE_EMPLOYEE,typeof(string)),//12
                                  new DataColumn(DATE_REGISTRATION,typeof(string)),//13
                                  new DataColumn(TIME_REGISTRATION,typeof(int)), //15
                               //   new DataColumn(@"Реальное время ухода",typeof(int)), //18
                                  new DataColumn(SERVER_SKD,typeof(string)), //19
                                  new DataColumn(NAME_CHECKPOINT,typeof(string)), //20
                                  new DataColumn(DIRECTION_WAY,typeof(string)), //21
                                  new DataColumn(DESIRED_TIME_IN,typeof(string)),//22
                                  new DataColumn(DESIRED_TIME_OUT,typeof(string)),//23
                                  new DataColumn(REAL_TIME_IN,typeof(string)),//24
                                  new DataColumn(REAL_TIME_OUT,typeof(string)), //25
                                  new DataColumn(EMPLOYEE_TIME_SPENT,typeof(int)), //26
                                  new DataColumn(EMPLOYEE_PLAN_TIME_WORKED,typeof(string)), //27
                                  new DataColumn(EMPLOYEE_BEING_LATE,typeof(string)),                    //28
                                  new DataColumn(EMPLOYEE_EARLY_DEPARTURE,typeof(string)),                 //29
                                  new DataColumn(EMPLOYEE_VACATION,typeof(string)),                 //30
                                  new DataColumn(EMPLOYEE_TRIP,typeof(string)),                 //31
                                  new DataColumn(DAY_OF_WEEK,typeof(string)),                 //32
                                  new DataColumn(EMPLOYEE_SICK_LEAVE,typeof(string)),                 //33
                                  new DataColumn(EMPLOYEE_ABSENCE,typeof(string)),     //34
                             //     new DataColumn(@"Код",typeof(string)),                        //35
                              //    new DataColumn(@"Вышестоящая группа",typeof(string)),         //36
                                  new DataColumn(GROUP_DECRIPTION,typeof(string)),            //37
                                  new DataColumn(EMPLOYEE_SHIFT_COMMENT,typeof(string)),                 //38
                                  new DataColumn(EMPLOYEE_POSITION,typeof(string)),                 //39
                                  new DataColumn(EMPLOYEE_SHIFT,typeof(string)),                 //40
                                  new DataColumn(EMPLOYEE_HOOKY,typeof(string)),                 //41
                                  new DataColumn(DEPARTMENT_ID,typeof(string)), //42
                                  new DataColumn(CHIEF_ID,typeof(string)) //43
                };
        readonly string[] arrayAllColumnsDataTablePeople =
            {
                                  NPP,//0
                                  FIO,//1
                                  CODE,//2
                                  GROUP,//3
                                  TIMEIN,//6
                                  TIMEOUT,//9
                                  N_ID, //10
                                  DEPARTMENT,//11
                                  PLACE_EMPLOYEE,//12
                                  DATE_REGISTRATION,//13
                                  TIME_REGISTRATION, //15
                                //  @"Реальное время ухода", //18
                                  SERVER_SKD, //19
                                  NAME_CHECKPOINT, //20
                                  DIRECTION_WAY, //21
                                  DESIRED_TIME_IN,//22
                                  DESIRED_TIME_OUT,//23
                                  REAL_TIME_IN,//24
                                  REAL_TIME_OUT, //25
                                  EMPLOYEE_TIME_SPENT, //26
                                  EMPLOYEE_PLAN_TIME_WORKED, //27
                                  EMPLOYEE_BEING_LATE,                    //28
                                  EMPLOYEE_EARLY_DEPARTURE,                 //29
                                  EMPLOYEE_VACATION,                 //30
                                  EMPLOYEE_TRIP,                 //31
                                  DAY_OF_WEEK,                    //32
                                  EMPLOYEE_SICK_LEAVE,                    //33
                                  EMPLOYEE_ABSENCE,      //34
                               //   @"Код",                           //35
                              //    @"Вышестоящая группа",            //36
                                  GROUP_DECRIPTION,                //37
                                  EMPLOYEE_SHIFT_COMMENT,      //38
                                  EMPLOYEE_POSITION,                    //39
                                  EMPLOYEE_SHIFT,                    //40
                                  EMPLOYEE_HOOKY,   //41
                                  DEPARTMENT_ID,                     //42
                                  @"Руководитель (код)"                     //43
        };
        readonly string[] orderColumnsFinacialReport =
            {
                                  FIO,//1
                                  CODE,//2
                                  DEPARTMENT,//11
                                  PLACE_EMPLOYEE,//12
                                 // DEPARTMENT_ID,                     //42
                                  DATE_REGISTRATION,//12
                                  DAY_OF_WEEK,                    //32
                                  DESIRED_TIME_IN,//22
                                  DESIRED_TIME_OUT,//23
                                  REAL_TIME_IN,//24
                                  REAL_TIME_OUT, //25
                                  EMPLOYEE_PLAN_TIME_WORKED, //27
                                  EMPLOYEE_BEING_LATE,                    //28
                                  EMPLOYEE_EARLY_DEPARTURE,                 //29
                                  EMPLOYEE_VACATION,                 //30
                                  EMPLOYEE_HOOKY,    //41
                                 //EMPLOYEE_TRIP,                 //31
                                  EMPLOYEE_SICK_LEAVE,                    //33
                                  EMPLOYEE_ABSENCE,      //34
                                  EMPLOYEE_SHIFT_COMMENT,                    //38
                                  EMPLOYEE_POSITION,                    //39
                                  EMPLOYEE_SHIFT                    //40
        };
        readonly string[] arrayHiddenColumnsFIO =
            {
                                  NPP,//0
                            TIMEIN,            //6
                            TIMEOUT,              //9
                            N_ID,               //10
                            DATE_REGISTRATION,         //12
                            TIME_REGISTRATION,        //15
                           // @"Реальное время ухода",     //18
                            SERVER_SKD,               //19
                            NAME_CHECKPOINT,        //20
                            DIRECTION_WAY,      //21
                            REAL_TIME_IN,//24
                            REAL_TIME_OUT, //25
                            EMPLOYEE_TIME_SPENT, //26
                            EMPLOYEE_PLAN_TIME_WORKED, //27
                            EMPLOYEE_BEING_LATE,                   //28
                            EMPLOYEE_EARLY_DEPARTURE,              //29
                            EMPLOYEE_VACATION,           //30
                            EMPLOYEE_TRIP,                 //31
                            DAY_OF_WEEK,                    //32
                            EMPLOYEE_SICK_LEAVE,                    //33
                            EMPLOYEE_ABSENCE,      //34
                         //   @"Код",                           //35
                        //    @"Вышестоящая группа",            //36
                            EMPLOYEE_SHIFT_COMMENT,                    //38
                            GROUP_DECRIPTION,                //37
                            EMPLOYEE_HOOKY
        };
        readonly string[] nameHidenColumnsArray =
            {
                                  NPP,//0
                TIMEIN,//6
                TIMEOUT,//9
                TIME_REGISTRATION, //15
             //   @"Реальное время ухода", //18
                SERVER_SKD, //19
                NAME_CHECKPOINT, //20
                DIRECTION_WAY, //21
                EMPLOYEE_TIME_SPENT, //26
                EMPLOYEE_PLAN_TIME_WORKED, //27
                EMPLOYEE_BEING_LATE,                    //28
                EMPLOYEE_EARLY_DEPARTURE,                 //29
                EMPLOYEE_VACATION,              //30
                EMPLOYEE_TRIP,                 //31
                DAY_OF_WEEK,  //32
                EMPLOYEE_SICK_LEAVE,  //33
                EMPLOYEE_ABSENCE,      //34
            //    @"Код",         //35
              //  @"Вышестоящая группа",            //36
                GROUP_DECRIPTION,                //37
                EMPLOYEE_HOOKY                   //43
        };
        static List<OutReasons> outResons = new List<OutReasons>();
        static List<OutPerson> outPerson = new List<OutPerson>();

        static DataTable dtPersonTemp = new DataTable("PersonTemp");
        static DataTable dtPersonTempAllColumns = new DataTable("PersonTempAllColumns");
        static DataTable dtPersonRegistrationsFullList = new DataTable("PersonRegistrationsFullList");
        static DataTable dtPeopleGroup = new DataTable("PeopleGroup");
        static DataTable dtPeopleListLoaded = new DataTable("PeopleLoaded");

        //Color of Person's Control elements which depend on the selected MenuItem  
        Color labelGroupCurrentBackColor;
        Color textBoxGroupCurrentBackColor;
        Color labelGroupDescriptionCurrentBackColor;
        Color textBoxGroupDescriptionCurrentBackColor;
        Color comboBoxFioCurrentBackColor;
        Color textBoxFIOCurrentBackColor;
        Color textBoxNavCurrentBackColor;


        public WinFormASTA()
        { InitializeComponent(); }

        private void Form1_Load(object sender, EventArgs e)
        { Form1Load(); }

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

            logger.Info("Настраиваю интерфейс....");
            Bitmap bmp = Properties.Resources.LogoRYIK;
            this.Icon = Icon.FromHandle(bmp.GetHicon());

            notifyIcon.Icon = this.Icon;
            notifyIcon.Visible = true;
            notifyIcon.BalloonTipText = @"Developed by ©Yuri Ryabchenko";
            notifyIcon.ShowBalloonTip(500);


            currentModeAppManual = true;
            _MenuItemTextSet(ModeItem, "Включить режим автоматических e-mail рассылок");
            _menuItemTooltipSet(ModeItem, "Включен интерактивный режим. Все рассылки остановлены.");

            myFileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath);
            strVersion = myFileVersionInfo.Comments + " ver." + myFileVersionInfo.FileVersion + " " + myFileVersionInfo.LegalCopyright;
            productName = myFileVersionInfo.ProductName;

            StatusLabel1.Text = myFileVersionInfo.ProductName + " ver." + myFileVersionInfo.FileVersion + " " + myFileVersionInfo.LegalCopyright;
            StatusLabel1.Alignment = ToolStripItemAlignment.Right;
            StatusLabel2.Text = " Начните работу с кнопки - \"Получить ФИО\"";

            contextMenu = new ContextMenu();  //Context Menu on notify Icon
            notifyIcon.ContextMenu = contextMenu;
            contextMenu.MenuItems.Add("About", AboutSoft);
            contextMenu.MenuItems.Add("-", AboutSoft);
            contextMenu.MenuItems.Add("Exit", ApplicationExit);

            this.Text = myFileVersionInfo.Comments;
            notifyIcon.Text = myFileVersionInfo.ProductName + "\nv." + myFileVersionInfo.FileVersion + "\n" + myFileVersionInfo.CompanyName;

            EditAnualDaysItem.Text = @"Выходные(рабочие) дни";
            EditAnualDaysItem.ToolTipText = @"Войти в режим добавления/удаления праздничных дней";

            _MenuItemEnabled(AddAnualDateItem, false);

            MembersGroupItem.Enabled = false;
            AddPersonToGroupItem.Enabled = false;
            CreateGroupItem.Enabled = false;
            DeleteGroupItem.Visible = false;
            DeletePersonFromGroupItem.Visible = false;
            CheckBoxesFiltersAll_Enable(false);
            TableModeItem.Visible = false;
            VisualModeItem.Visible = false;
            ChangeColorMenuItem.Visible = false;
            TableExportToExcelItem.Visible = false;
            listFioItem.Visible = false;
            dataGridView1.ShowCellToolTips = true;
            groupBoxProperties.Visible = false;

            comboBoxFio.DrawMode = DrawMode.OwnerDrawFixed;
            comboBoxFio.DrawItem += new DrawItemEventHandler(ComboBox_DrawItem);

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
            dateTimePickerEnd.MaxDate = DateTime.Parse("2025-12-31");
            dateTimePickerStart.Value = DateTime.Parse(DateTime.Now.Year + "-" + DateTime.Now.Month + "-01");
            dateTimePickerEnd.Value = DateTime.Now;
            //DateTime.Now.ToString("yyyy-MM-dd HH:mm")

            numUpDownHourStart.Value = 9;
            numUpDownMinuteStart.Value = 0;
            numUpDownHourEnd.Value = 18;
            numUpDownMinuteEnd.Value = 0;

            PersonOrGroupItem.Text = "Работать с одной персоной";
            toolTip1.SetToolTip(textBoxGroup, "Создать группу");
            toolTip1.SetToolTip(textBoxGroupDescription, "Изменить описание группы");
            StatusLabel2.Text = "";

            logger.Info("TryMakeDB...");
            TryMakeDB();
            logger.Info("UpdateTableOfDB...");
            UpdateTableOfDB();
            logger.Info("SetTechInfoIntoDB...");
            SetTechInfoIntoDB();

            //read last saved parameters from db and Registry and set their into variables
            logger.Info("Загружаю настройки программы...");
            BoldAnualDates();
            LoadPrevioslySavedParameters();

            sServer1 = sServer1Registry?.Length > 0 ? sServer1Registry : sServer1DB;
            sServer1UserName = sServer1UserNameRegistry?.Length > 0 ? sServer1UserNameRegistry : sServer1UserNameDB;
            sServer1UserPassword = sServer1UserPasswordRegistry?.Length > 0 ? sServer1UserPasswordRegistry : sServer1UserPasswordDB;

            mailServer = mailServerRegistry?.Length > 0 ? mailServerRegistry : mailServerDB;
            mailServerUserName = mailServerUserNameRegistry?.Length > 0 ? mailServerUserNameRegistry : mailServerUserNameDB;
            mailServerUserPassword = mailServerUserPasswordRegistry?.Length > 0 ? mailServerUserPasswordRegistry : mailServerUserPasswordDB;

            mysqlServer = mysqlServerRegistry?.Length > 0 ? mysqlServerRegistry : mysqlServerDB;
            mysqlServerUserName = mysqlServerUserNameRegistry?.Length > 0 ? mysqlServerUserNameRegistry : mysqlServerUserNameDB;
            mysqlServerUserPassword = mysqlServerUserPasswordRegistry?.Length > 0 ? mysqlServerUserPasswordRegistry : mysqlServerUserPasswordDB;

            clrRealRegistration = clrRealRegistrationRegistry != Color.PaleGreen ? clrRealRegistrationRegistry : Color.PaleGreen;


            logger.Info("Настраиваю переменные....");

            //Prepare DataTables
            dtPeople.Columns.AddRange(dcPeople);
            dtPeople.DefaultView.Sort = "[Группа], [Фамилия Имя Отчество], [Дата регистрации], [Время регистрации], [Фактич. время прихода ЧЧ:ММ], [Фактич. время ухода ЧЧ:ММ] ASC";

            //Clone default column name and structure from 'dtPeople' to other DataTables
            dtPersonTemp = dtPeople.Clone();  //Copy only structure(Name of columns)
            dtPersonRegistrationsFullList = dtPeople.Clone();  //Copy only structure(Name of columns)
            dtPeopleGroup = dtPeople.Clone();  //Copy only structure(Name of columns)

            logger.Info("Программа " + productName + " полностью загружена....");

            if (currentModeAppManual)
            {
                nameOfLastTableFromDB = "ListFIO";
                SeekAndShowMembersOfGroup("");
                logger.Info("Программа запущена в интерактивном режиме....");
            }
            else
            {
                _controlEnable(comboBoxFio, false);
                nameOfLastTableFromDB = "Mailing";
                logger.Info(productName + " включен автоматический режим....");

                ShowDataTableDbQuery(databasePerson, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
                "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
                "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
                " ORDER BY RecipientEmail asc, DateCreated desc; ");
                dataGridView1.Select();
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
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'ConfigDB' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, ParameterName TEXT, Value TEXT, Description TEXT, DateCreated TEXT, IsPassword TEXT, IsExample TEXT, UNIQUE ('ParameterName', 'IsExample') ON CONFLICT REPLACE);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'PeopleGroupDesciption' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, GroupPerson TEXT, GroupPersonDescription TEXT, AmountStaffInDepartment TEXT, Recipient TEXT, UNIQUE ('GroupPerson') ON CONFLICT REPLACE);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'PeopleGroup' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, FIO TEXT, NAV TEXT, GroupPerson TEXT, ControllingHHMM TEXT, ControllingOUTHHMM TEXT, " +
                    "Shift TEXT, Comment TEXT, Department TEXT, PositionInDepartment TEXT, DepartmentId TEXT, City TEXT, Boss TEXT, UNIQUE ('FIO', 'NAV', 'GroupPerson', 'DepartmentId') ON CONFLICT REPLACE);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'ListOfWorkTimeShifts' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, NAV TEXT, DayStartShift TEXT, " +
                    "MoStart REAL,MoEnd REAL, TuStart REAL,TuEnd REAL, WeStart REAL,WeEnd REAL, ThStart REAL,ThEnd REAL, FrStart REAL,FrEnd REAL, " +
                    "SaStart REAL,SaEnd REAL, SuStart REAL,SuEnd REAL, Status Text, Comment TEXT, DayInputed TEXT, UNIQUE ('NAV', 'DayStartShift') ON CONFLICT REPLACE);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'TechnicalInfo' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, PCName TEXT, POName TEXT, POVersion TEXT, LastDateStarted TEXT, CurrentUser TEXT, " +
                    "FreeRam TEXT, GuidAppication TEXT);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'BoldedDates' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, DayBolded TEXT, NAV TEXT, DayType TEXT, DayDesciption TEXT, DateCreated TEXT);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'EquipmentSettings' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, EquipmentParameterName TEXT, EquipmentParameterValue TEXT, EquipmentParameterServer TEXT," +
                    "Reserv1, Reserv2, UNIQUE ('EquipmentParameterName', 'EquipmentParameterServer') ON CONFLICT REPLACE);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'LastTakenPeopleComboList' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, ComboList TEXT, " +
                    "DateCreated TEXT, UNIQUE ('ComboList') ON CONFLICT REPLACE);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'Mailing' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, SenderEmail TEXT, " +
                    "RecipientEmail TEXT, GroupsReport TEXT, NameReport TEXT, Description TEXT, Period TEXT, Status TEXT, SendingLastDate TEXT, TypeReport TEXT, DayReport TEXT, DateCreated TEXT" +
                    ", UNIQUE ('RecipientEmail', 'GroupsReport', 'NameReport', 'Description', 'Period', 'TypeReport', 'DayReport') ON CONFLICT REPLACE);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'MailingException' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, RecipientEmail TEXT, NameReport TEXT, Description TEXT, DayReport TEXT, DateCreated TEXT" +
                    ", UNIQUE ('RecipientEmail', 'NameReport') ON CONFLICT REPLACE);", databasePerson);
            ExecuteSql("CREATE TABLE IF NOT EXISTS 'SelectedCityToLoadFromWeb' ('Id' INTEGER PRIMARY KEY AUTOINCREMENT, City TEXT, DateCreated TEXT" +
                    ", UNIQUE ('City') ON CONFLICT REPLACE);", databasePerson);
        }

        private void UpdateTableOfDB()
        {
            TryUpdateStructureSqlDB("ConfigDB", "ParameterName TEXT, Value TEXT, Description TEXT, DateCreated TEXT, IsPassword TEXT, IsExample TEXT", databasePerson);
            TryUpdateStructureSqlDB("PeopleGroupDesciption", "GroupPerson TEXT, GroupPersonDescription TEXT, AmountStaffInDepartment TEXT, Recipient TEXT", databasePerson);
            TryUpdateStructureSqlDB("PeopleGroup", "FIO TEXT, NAV TEXT, GroupPerson TEXT, ControllingHHMM TEXT, ControllingOUTHHMM TEXT, " +
                    "Shift TEXT, Comment TEXT, Department TEXT, PositionInDepartment TEXT, DepartmentId TEXT, City TEXT, Boss TEXT", databasePerson);
            TryUpdateStructureSqlDB("ListOfWorkTimeShifts", "NAV TEXT, DayStartShift TEXT, MoStart REAL,MoEnd REAL, TuStart REAL,TuEnd REAL, WeStart REAL,WeEnd REAL, ThStart REAL,ThEnd REAL, FrStart REAL,FrEnd REAL, " +
                    "SaStart REAL,SaEnd REAL, SuStart REAL,SuEnd REAL, Status Text, Comment TEXT, DayInputed TEXT", databasePerson);
            TryUpdateStructureSqlDB("TechnicalInfo", "PCName TEXT, POName TEXT, POVersion TEXT, LastDateStarted TEXT, CurrentUser TEXT, FreeRam TEXT, GuidAppication TEXT", databasePerson);
            TryUpdateStructureSqlDB("BoldedDates", "DayBolded TEXT, NAV TEXT, DayType TEXT, DayDesciption TEXT, DateCreated TEXT", databasePerson);
            TryUpdateStructureSqlDB("EquipmentSettings", "EquipmentParameterName TEXT, EquipmentParameterValue TEXT, EquipmentParameterServer TEXT, Reserv1 TEXT, Reserv2 TEXT", databasePerson);
            TryUpdateStructureSqlDB("LastTakenPeopleComboList", "ComboList TEXT, DateCreated TEXT", databasePerson);
            TryUpdateStructureSqlDB("Mailing", "SenderEmail TEXT, RecipientEmail TEXT, GroupsReport TEXT, NameReport TEXT, Description TEXT, Period TEXT, Status TEXT, SendingLastDate TEXT, TypeReport TEXT, DayReport TEXT, DateCreated TEXT", databasePerson);
            TryUpdateStructureSqlDB("MailingException", "RecipientEmail TEXT, NameReport TEXT, Description TEXT, DayReport TEXT, DateCreated TEXT", databasePerson);
            TryUpdateStructureSqlDB("SelectedCitytoLoadFromWeb", "City TEXT, DateCreated TEXT", databasePerson);
        }

        private void SetTechInfoIntoDB() //Write Technical Info in DB 
        {
            guid = System.Runtime.InteropServices.Marshal.GetTypeLibGuidForAssembly(System.Reflection.Assembly.GetExecutingAssembly()).ToString(); // получаем GIUD приложения
            pcname = Environment.MachineName + "(" + Environment.OSVersion + ")";
            poname = myFileVersionInfo.FileName + "(" + myFileVersionInfo.ProductName + ")";
            poversion = myFileVersionInfo.FileVersion;
            string LastDateStarted = DateTime.Now.ToYYYYMMDDHHMM();
            string CurrentUser = Environment.UserName;
            string FreeRam = "RAM: " + Environment.WorkingSet.ToString();

            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'TechnicalInfo' (PCName, POName, POVersion, LastDateStarted, CurrentUser, FreeRam, GuidAppication) " +
                        " VALUES (@PCName, @POName, @POVersion, @LastDateStarted, @CurrentUser, @FreeRam, @GuidAppication)", sqlConnection))
                    {
                        sqlCommand.Parameters.Add("@PCName", DbType.String).Value = pcname;
                        sqlCommand.Parameters.Add("@POName", DbType.String).Value = poname;
                        sqlCommand.Parameters.Add("@POVersion", DbType.String).Value = poversion;
                        sqlCommand.Parameters.Add("@LastDateStarted", DbType.String).Value = LastDateStarted;
                        sqlCommand.Parameters.Add("@CurrentUser", DbType.String).Value = CurrentUser;
                        sqlCommand.Parameters.Add("@FreeRam", DbType.String).Value = FreeRam;
                        sqlCommand.Parameters.Add("@GuidAppication", DbType.String).Value = guid;
                        try { sqlCommand.ExecuteNonQuery(); } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                    }
                }
            }
            LastDateStarted = null; CurrentUser = null; FreeRam = null;
        }

        private void LoadPrevioslySavedParameters()   //Select Previous Data from DB and write it into the combobox and Parameters
        {
            string modeApp = "";
            int iCombo = 0;
            int numberOfFio = 0;

            //  numberOfFio = CountRowInTableDB("PeopleGroup");

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
                                }
                                catch (Exception expt) { logger.Info(expt.ToString()); }
                            }
                        }
                    }

                    using (var sqlCommand = new SQLiteCommand("SELECT FIO FROM PeopleGroup;", sqlConnection))
                    {
                        using (var reader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in reader)
                            {
                                if (record["FIO"]?.ToString()?.Length > 0)
                                { numberOfFio++; }
                            }
                        }
                    }
                }


                //loading parameters
                listParameters = new List<ParameterConfig>();

                ParameterOfConfigurationInSQLiteDB parameters = new ParameterOfConfigurationInSQLiteDB();
                parameters.databasePerson = databasePerson;
                listParameters = parameters.GetParameters("%%").FindAll(x => x.isExample == "no"); //load only real data

                    DEFAULT_DAY_OF_SENDING_REPORT = GetValueOfConfigParameter(listParameters, @"DEFAULT_DAY_OF_SENDING_REPORT", "28");
                    DEFAULT_RECEIVING_PORT_MAILSERVER = GetValueOfConfigParameter(listParameters, @"DEFAULT_RECEIVING_PORT_MAILSERVER", "587");
                    clrRealRegistrationRegistry = Color.FromName(GetValueOfConfigParameter(listParameters, @"clrRealRegistration", "PaleGreen"));

                    sServer1DB = GetValueOfConfigParameter(listParameters, @"SKDServer", null);
                    sServer1UserNameDB = GetValueOfConfigParameter(listParameters, @"SKDUser", null);
                    sServer1UserPasswordDB = GetValueOfConfigParameter(listParameters, @"SKDUserPassword", null, true);

                    mysqlServerDB = GetValueOfConfigParameter(listParameters, @"MySQLServer", null);
                    mysqlServerUserNameDB = GetValueOfConfigParameter(listParameters, @"MySQLUser", null);
                    mysqlServerUserPasswordDB = GetValueOfConfigParameter(listParameters, @"MySQLUserPassword", null, true);

                    mailServerDB = GetValueOfConfigParameter(listParameters, @"MailServer", null);
                    mailServerUserNameDB = GetValueOfConfigParameter(listParameters, @"MailUser", null);
                    mailServerUserPasswordDB = GetValueOfConfigParameter(listParameters, @"MailUserPassword", null, true);
                
                listParameters = null;
                parameters = null;
            }
            
            try { comboBoxFio.SelectedIndex = 0; } catch { }

            using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(myRegKey, Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree, System.Security.AccessControl.RegistryRights.ReadKey))
            {
                try { sServer1Registry = EvUserKey?.GetValue("SKDServer")?.ToString(); } catch { logger.Warn("Registry GetValue SKDServer"); }
                try { sServer1UserNameRegistry = EncryptionDecryptionCriticalData.DecryptBase64ToString(EvUserKey?.GetValue("SKDUser")?.ToString(), btsMess1, btsMess2).ToString(); } catch { }
                try { sServer1UserPasswordRegistry = EncryptionDecryptionCriticalData.DecryptBase64ToString(EvUserKey?.GetValue("SKDUserPassword")?.ToString(), btsMess1, btsMess2).ToString(); } catch { }

                try { mailServerRegistry = EvUserKey?.GetValue("MailServer")?.ToString(); } catch { logger.Warn("Registry GetValue Mail"); }
                try { mailServerUserNameRegistry = EvUserKey?.GetValue("MailUser")?.ToString(); } catch { }
                try { mailServerUserPasswordRegistry = EncryptionDecryptionCriticalData.DecryptBase64ToString(EvUserKey?.GetValue("MailUserPassword")?.ToString(), btsMess1, btsMess2).ToString(); } catch { }

                try { mysqlServerRegistry = EvUserKey?.GetValue("MySQLServer")?.ToString() ; } catch { logger.Warn("Registry GetValue MySQL"); }
                try { mysqlServerUserNameRegistry = EvUserKey?.GetValue("MySQLUser")?.ToString(); } catch { }
                try { mysqlServerUserPasswordRegistry = EncryptionDecryptionCriticalData.DecryptBase64ToString(EvUserKey?.GetValue("MySQLUserPassword")?.ToString(), btsMess1, btsMess2).ToString(); } catch { }

                try { modeApp = EvUserKey?.GetValue("ModeApp")?.ToString(); } catch { logger.Warn("Registry GetValue ModeApp"); }
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
        
        private string GetValueOfConfigParameter(List<ParameterConfig> listOfParameters, string nameParameter, string defaultValue, bool pass = false)
        {
                return listOfParameters.FindLast(x => x.parameterName == nameParameter)?.parameterValue != null ?
                       listParameters.FindLast(x => x.parameterName == nameParameter)?.parameterValue :
                       defaultValue; 
        }


        private void AddParameterInConfigItem_Click(object sender, EventArgs e)
        {
            AddParameterInConfig();
        }



        private void AddParameterInConfig()
        {
            _controlVisible(panelView, false);
            btnPropertiesSave.Text = "Сохранить параметр";
            RemoveClickEvent(btnPropertiesSave);
            btnPropertiesSave.Click += new EventHandler(ButtonPropertiesSave_inConfig);

            listParameters = new List<ParameterConfig>();
            ParameterOfConfigurationInSQLiteDB parameter = new ParameterOfConfigurationInSQLiteDB();

            parameter.databasePerson = databasePerson;
            listParameters = parameter.GetParameters("%%");

            foreach (string sParameter in allParametersOfConfig)
            {
                if (!(listParameters.FindLast(x => x.parameterName == sParameter)?.parameterValue?.Length > 0))
                {
                    listParameters.Add(new ParameterConfig()
                    {
                        parameterName = sParameter,
                        parameterDescription = "Example",
                        parameterValue = "",
                        isPassword = false,
                        isExample = "yes"
                    });
                }
            }

            // listParameters = parameter.GetParameters("%%").FindAll(x => x.isExample != "no"); //load only real data

            InitializeParameterFormSettings(listParameters);
        }

        private void InitializeParameterFormSettings(List<ParameterConfig> listParameters)
        {
            panelViewResize(numberPeopleInLoading);

            periodCombo = new ListBox
            {
                Location = new Point(20, 120),
                Size = new Size(590, 100),
                Parent = groupBoxProperties,
                DrawMode = DrawMode.OwnerDrawFixed,
                Sorted=true
            };

            periodCombo.DrawItem += new DrawItemEventHandler(ListBox_DrawItem);
            periodCombo.DataSource = listParameters.Select(x => x.parameterName).ToList();
            if (listParameters.Count > 0) periodCombo.SelectedIndex = 0;
            toolTip1.SetToolTip(periodCombo, "Перечень параметров");

            labelServer1 = new Label
            {
                Text = "",
                BackColor = Color.PaleGreen,
                Location = new Point(20, 59),
                Size = new Size(590, 24),
                BorderStyle = BorderStyle.None,
                TextAlign = ContentAlignment.MiddleLeft,
                Parent = groupBoxProperties
            };
            textBoxSettings16 = new TextBox
            {
                Text = "",
                Location = new Point(150, 61),
                Size = new Size(150, 20),
                BorderStyle = BorderStyle.FixedSingle,
                Parent = groupBoxProperties
            };
            toolTip1.SetToolTip(textBoxSettings16, "");

            labelSettings9 = new Label
            {
                Text = "",
                BackColor = Color.PaleGreen,
                Location = new Point(340, 61),
                Size = new Size(250, 22),
                BorderStyle = BorderStyle.None,
                TextAlign = ContentAlignment.TopLeft,
                Parent = groupBoxProperties
            };

            checkBox1 = new CheckBox
            {
                Text = "Данные хранить зашифрованными",
                BackColor = Color.PaleGreen,
                Location = new Point(20, 91),
                Size = new Size(590, 22),
                TextAlign = ContentAlignment.TopLeft,
                Parent = groupBoxProperties,
                Checked = false
            };
            //  periodCombo.KeyPress += new KeyPressEventHandler(SelectComboBoxParameters_SelectedIndexChanged);
            periodCombo.SelectedIndexChanged += new EventHandler(ListBox_SelectedIndexChanged);
            textBoxSettings16.TextChanged += new EventHandler(textbox_textChanged);
            checkBox1.CheckStateChanged += new EventHandler(checkBox1_CheckStateChanged);
            textBoxSettings16.BringToFront();
            labelSettings9.BringToFront();
            checkBox1.BringToFront();

            groupBoxProperties.Visible = true;
        }

        private void checkBox1_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                textBoxSettings16.PasswordChar = '*';
            }
            else
            {
                textBoxSettings16.PasswordChar = '\0';
            }
        }


        private void ListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string result = _listBoxReturnSelected(sender as ListBox);
            string tooltip = "";

            checkBox1.Checked = false;
            labelServer1.Text = "";
            labelSettings9.Text = "";
            textBoxSettings16.Text = "";
            toolTip1.SetToolTip(textBoxSettings16, tooltip);
            
            checkBox1.Checked = listParameters.FindLast(x => x.parameterName == result).isPassword;
            labelServer1.Text = result;
            labelSettings9.Text = listParameters.FindLast(x => x.parameterName == result)?.parameterDescription;
            textBoxSettings16.Text = listParameters.FindLast(x => x.parameterName == result)?.parameterValue;
            tooltip = listParameters.FindLast(x => x.parameterName == result)?.parameterDescription;
            toolTip1.SetToolTip(textBoxSettings16, tooltip);
        }

        private void textbox_textChanged(object sender, EventArgs e)
        {
            int result;
            bool correct = false;

            //allow numbers from 1 to 28
            if ((sender as TextBox).Text.Length > 0)
            {
                correct = Int32.TryParse((sender as TextBox).Text, out result);
                if (correct)
                {
                    if (result > 28) { (sender as TextBox).Text = "28"; }
                    else if (result < 1) { (sender as TextBox).Text = "1"; }
                    else
                    {
                        (sender as TextBox).Text = result.ToString();
                    }
                }
            }
        }

        private void ButtonPropertiesSave_inConfig(object sender, EventArgs e) //SaveProperties()
        {
            string resultSaving;
            ParameterOfConfigurationInSQLiteDB parameter = new ParameterOfConfigurationInSQLiteDB();
            parameter.databasePerson = databasePerson;
            parameter.ParameterName = labelServer1.Text;
            parameter.ParameterValue = textBoxSettings16.Text;
            parameter.ParameterDescription = labelSettings9.Text;
            parameter.isPassword = checkBox1.Checked;
            parameter.isExample = "no";
            resultSaving = parameter.SaveParameter();
            MessageBox.Show(parameter.ParameterName+"="+ parameter.ParameterValue + "\n"+resultSaving);

            DisposeTemporaryControls();
            _controlVisible(panelView, true);

            ShowDataTableDbQuery(databasePerson, "ConfigDB", "SELECT ParameterName AS 'Имя параметра', " +
            "Value AS 'Данные', Description AS 'Описание', DateCreated AS 'Дата создания/модификации'",
            " ORDER BY ParameterName asc, DateCreated desc; ");
        }




        private void ExecuteSql(string SqlQuery, System.IO.FileInfo FileDB) //Prepare DB and execute of SQL Query
        {
            if (!System.IO.File.Exists(databasePerson.FullName))
            { SQLiteConnection.CreateFile(databasePerson.FullName); }
            using (var connection = new SQLiteConnection($"Data Source={databasePerson.FullName};Version=3;"))
            {
                connection.Open();

                using (var command = new SQLiteCommand(SqlQuery, connection))
                {
                    try
                    { command.ExecuteNonQuery(); } catch (Exception expt) { logger.Info("ExecuteSql: " + SqlQuery + "\n" + expt.ToString()); }
                }
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

        //void ShowDataTableDbQuery(
        private void ShowDataTableDbQuery(System.IO.FileInfo databasePerson, string myTable, string mySqlQuery, string mySqlWhere) //Query data from the Table of the DB
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
            logger.Trace("ShowDataTableDbQuery: " + iCounterLine);

            nameOfLastTableFromDB = myTable;
            sLastSelectedElement = "dataGridView";
        }

        private void ShowDatatableOnDatagridview(DataTable dt, string[] nameHidenColumnsArray1, string nameLastTable) //Query data from the Table of the DB
        {
            DataTable dataTable = dt.Copy();
            for (int i = 0; i < nameHidenColumnsArray1.Length; i++)
            {
                if (nameHidenColumnsArray1[i]?.Length > 0)
                    try { dataTable.Columns[nameHidenColumnsArray1[i]].ColumnMapping = MappingType.Hidden; } catch { }
            }

            _dataGridViewSource(dataTable);
            _toolStripStatusLabelSetText(StatusLabel2, "Всего записей: " + _dataGridView1RowsCount());

            nameOfLastTableFromDB = nameLastTable;
            sLastSelectedElement = "dataGridView";
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
                            logger.Info("Удалена таблица: " + myTable);
                        } catch { logger.Info("DeleteTable: далить таблицу не удалось: " + myTable); }
                    }
                    using (var sqlCommand = new SQLiteCommand("vacuum;", sqlConnection))   //vacuum DB
                    {
                        sqlCommand.ExecuteNonQuery();
                    }
                    sqlConnection.Close();
                }
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

        private void DeleteDataTableQueryParameters(System.IO.FileInfo databasePerson, string myTable, string mySqlParameter1, string mySqlData1,
            string mySqlParameter2 = "", string mySqlData2 = "", string mySqlParameter3 = "", string mySqlData3 = "",
            string mySqlParameter4 = "", string mySqlData4 = "", string mySqlParameter5 = "", string mySqlData5 = "", string mySqlParameter6 = "", string mySqlData6 = "") //Delete data from the Table of the DB by NAV (both parameters are string)
        {
            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    var sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();

                    if (mySqlParameter1.Length > 0 && mySqlParameter2.Length > 0 && mySqlParameter3.Length > 0 && mySqlParameter4.Length > 0
                        && mySqlParameter5.Length > 0 && mySqlParameter6.Length > 0)
                    {
                        using (var sqlCommand = new SQLiteCommand("DELETE FROM '" + myTable + "' Where " + mySqlParameter1 + "= @" + mySqlParameter1 +
                            " AND " + mySqlParameter2 + "= @" + mySqlParameter2 + " AND " + mySqlParameter3 + "= @" + mySqlParameter3 +
                            " AND " + mySqlParameter4 + "= @" + mySqlParameter4 + " AND " + mySqlParameter5 + "= @" + mySqlParameter5 +
                            " AND " + mySqlParameter6 + "= @" + mySqlParameter6 + ";", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@" + mySqlParameter1, DbType.String).Value = mySqlData1;
                            sqlCommand.Parameters.Add("@" + mySqlParameter2, DbType.String).Value = mySqlData2;
                            sqlCommand.Parameters.Add("@" + mySqlParameter3, DbType.String).Value = mySqlData3;
                            sqlCommand.Parameters.Add("@" + mySqlParameter4, DbType.String).Value = mySqlData4;
                            sqlCommand.Parameters.Add("@" + mySqlParameter5, DbType.String).Value = mySqlData5;
                            sqlCommand.Parameters.Add("@" + mySqlParameter6, DbType.String).Value = mySqlData6;

                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }
                    }
                    else if (mySqlParameter1.Length > 0 && mySqlParameter2.Length > 0 && mySqlParameter3.Length > 0 && mySqlParameter4.Length > 0
                        && mySqlParameter5.Length > 0)
                    {
                        using (var sqlCommand = new SQLiteCommand("DELETE FROM '" + myTable + "' Where " + mySqlParameter1 + "= @" + mySqlParameter1 +
                            " AND " + mySqlParameter2 + "= @" + mySqlParameter2 + " AND " + mySqlParameter3 + "= @" + mySqlParameter3 +
                            " AND " + mySqlParameter4 + "= @" + mySqlParameter4 + " AND " + mySqlParameter5 + "= @" + mySqlParameter5 + ";", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@" + mySqlParameter1, DbType.String).Value = mySqlData1;
                            sqlCommand.Parameters.Add("@" + mySqlParameter2, DbType.String).Value = mySqlData2;
                            sqlCommand.Parameters.Add("@" + mySqlParameter3, DbType.String).Value = mySqlData3;
                            sqlCommand.Parameters.Add("@" + mySqlParameter4, DbType.String).Value = mySqlData4;
                            sqlCommand.Parameters.Add("@" + mySqlParameter5, DbType.String).Value = mySqlData5;

                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }
                    }
                    else if (mySqlParameter1.Length > 0 && mySqlParameter2.Length > 0 && mySqlParameter3.Length > 0 && mySqlParameter4.Length > 0)
                    {
                        using (var sqlCommand = new SQLiteCommand("DELETE FROM '" + myTable + "' Where " + mySqlParameter1 + "= @" + mySqlParameter1 +
                            " AND " + mySqlParameter2 + "= @" + mySqlParameter2 + " AND " + mySqlParameter3 + "= @" + mySqlParameter3 +
                            " AND " + mySqlParameter4 + "= @" + mySqlParameter4 + ";", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@" + mySqlParameter1, DbType.String).Value = mySqlData1;
                            sqlCommand.Parameters.Add("@" + mySqlParameter2, DbType.String).Value = mySqlData2;
                            sqlCommand.Parameters.Add("@" + mySqlParameter3, DbType.String).Value = mySqlData3;
                            sqlCommand.Parameters.Add("@" + mySqlParameter4, DbType.String).Value = mySqlData4;

                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }
                    }
                    else if (mySqlParameter1.Length > 0 && mySqlParameter2.Length > 0 && mySqlParameter3.Length > 0)
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
                    sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                }
            }
            myTable = null; mySqlParameter1 = null; mySqlData1 = null;
        }


        private void CheckAliveIntellectServer(string serverName, string userName, string userPasswords) //Check alive the SKD Intellect-server and its DB's 'intellect'
        {
            bServer1Exist = false;
            string stringConnection;
             _toolStripStatusLabelSetText(StatusLabel2, "Проверка доступности " + serverName + ". Ждите окончания процесса...");

            stringConnection = "Data Source=" + serverName + "\\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID=" + userName + ";Password=" + userPasswords + "; Connect Timeout=5";

            try
            {
                logger.Trace("CheckAliveIntellectServer: " + stringConnection);
                using (var sqlConnection = new System.Data.SqlClient.SqlConnection(stringConnection))
                {
                    sqlConnection.Open();
                    string query = "SELECT database_id FROM sys.databases WHERE Name ='intellect' ";
                    logger.Trace("CheckAliveIntellectServer: " + query);

                    using (var sqlCommand = new System.Data.SqlClient.SqlCommand(query, sqlConnection))
                    {
                        using (var sqlReader = sqlCommand.ExecuteReader())
                        {
                            bServer1Exist = true;
                        }
                    }
                }
            }
            catch (Exception expt)
            {
                bServer1Exist = false;
                logger.Info(expt.ToString());
            }

            if (!bServer1Exist)
            {
                _toolStripStatusLabelSetText(StatusLabel2, "Ошибка доступа к " + serverName + " SQL БД СКД-сервера!");
                _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
            }

            stringConnection = null;
        }

        private async void GetFio_Click(object sender, EventArgs e)  //DoListsFioGroupsMailings()
        {
            _ProgressBar1Start();
            currentAction = "GetFIO";
            CheckBoxesFiltersAll_CheckedState(false);
            CheckBoxesFiltersAll_Enable(false);
            _MenuItemEnabled(LoadDataItem, false);
            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(SettingsMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            _MenuItemEnabled(GetFioItem, false);
            _controlEnable(dataGridView1, false);

            await Task.Run(() => CheckAliveIntellectServer(sServer1, sServer1UserName, sServer1UserPassword));

            if (bServer1Exist)
            {
                await Task.Run(() => DoListsFioGroupsMailings());

                _MenuItemVisible(listFioItem, true);
                _MenuItemEnabled(SettingsMenuItem, true);
                _MenuItemEnabled(GetFioItem, true);
                _MenuItemEnabled(FunctionMenuItem, true);
                _MenuItemEnabled(LoadDataItem, true);
                _MenuItemEnabled(GroupsMenuItem, true);

                _controlEnable(dataGridView1, true);
                _controlVisible(dataGridView1, true);
                _controlVisible(pictureBox1, false);
                _controlEnable(comboBoxFio, true);

                _toolStripStatusLabelForeColor(StatusLabel1, Color.Black);
                _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);
            }
            else
            {
                _MenuItemEnabled(SettingsMenuItem, true);
            }
            _ProgressBar1Stop();
        }

        static DataTable dtTempIntermediate;
        static List<PeopleShift> peopleShifts = new List<PeopleShift>();
        private void DoListsFioGroupsMailings()  //  GetDataFromRemoteServers()  ImportTablePeopleToTableGroupsInLocalDB()
        {
            _toolStripStatusLabelSetText(StatusLabel2, "Получаю данные с серверов...");
            GetUsersFromAD();

            dtTempIntermediate = dtPeople.Clone();
            GetDataFromRemoteServers(dtTempIntermediate, peopleShifts);

            _toolStripStatusLabelSetText(StatusLabel2, "Формирую и записываю группы в локальную базу...");
            WriteGroupsMailingsInLocalDb(dtTempIntermediate, peopleShifts);

            _toolStripStatusLabelSetText(StatusLabel2, "Записываю ФИО в локальную базу...");

            WritePeopleInLocalDB(databasePerson.ToString(), dtTempIntermediate);
            // ImportListGroupsDescriptionInLocalDB

            if (currentAction != @"sendEmail")
            {
                var namesDistinctColumnsArray = arrayAllColumnsDataTablePeople.Except(arrayHiddenColumnsFIO).ToArray(); //take distinct data
                dtPersonTemp = GetDistinctRecords(dtTempIntermediate, namesDistinctColumnsArray);
                ShowDatatableOnDatagridview(dtPersonTemp, arrayHiddenColumnsFIO, "ListFIO");
                _toolStripStatusLabelSetText(StatusLabel2, "Списки ФИО и департаментов получены.");
                namesDistinctColumnsArray = null;
            }
        }


        //Get the list of registered users
        private void GetDataFromRemoteServers(DataTable dataTablePeople, List<PeopleShift> peopleShifts)
        {
            PersonFull personFromServer = new PersonFull();
            DataRow row;
            string stringConnection;
            string query;
            string fio = "";
            string nav = "";
            string groupName = "";
            string depName = "";
            string depBoss = "";

            string timeStart = ConvertDecimalTimeToStringHHMM(_numUpDownReturn(numUpDownHourStart), _numUpDownReturn(numUpDownMinuteStart));
            string timeEnd = ConvertDecimalTimeToStringHHMM(_numUpDownReturn(numUpDownHourEnd), _numUpDownReturn(numUpDownMinuteEnd));
            string dayStartShift = "";
            string dayStartShift_ = "";

            listFIO = new List<Person>();
            List<Department> departments = new List<Department>();
            //  List<string> listCodesWithIdCard = new List<string>(); //NAV-codes staff who have idCards

            _comboBoxClr(comboBoxFio);
            _toolStripStatusLabelSetText(StatusLabel2, "Запрашиваю данные с " + sServer1 + ". Ждите окончания процесса...");
            stringConnection = "Data Source=" + sServer1 + "\\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID=" + sServer1UserName + ";Password=" + sServer1UserPassword + "; Connect Timeout=60";
            logger.Trace(stringConnection);

            string confitionToLoad = "";
            using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
            {
                sqlConnection.Open();

                _toolStripStatusLabelSetText(StatusLabel2, "Формирую запрос для получения списка ФИО из MySQL базы...");
                using (var sqlCommand = new SQLiteCommand("SELECT City FROM SelectedCityToLoadFromWeb;", sqlConnection))
                {
                    using (var reader = sqlCommand.ExecuteReader())
                    {
                        foreach (DbDataRecord record in reader)
                        {
                            if (record["City"]?.ToString()?.Length > 0 && !record["City"].ToString().ToLower().Contains("city"))
                            {
                                if (confitionToLoad.Length == 0)
                                {
                                    confitionToLoad += " WHERE city LIKE '" + record["City"].ToString() + "'";
                                }
                                else
                                { confitionToLoad += " OR city LIKE '" + record["City"].ToString() + "'"; }
                            }
                        }
                    }
                }
            }

            try
            {
                // import users and group from SCA server
                using (var sqlConnection = new System.Data.SqlClient.SqlConnection(stringConnection))
                {
                    sqlConnection.Open();

                    //import group from SCA server
                    query = "SELECT id,level_id,name,owner_id,parent_id,region_id,schedule_id FROM OBJ_DEPARTMENT";
                    logger.Trace(query);
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
                                        departments.Add(new Department()
                                        {
                                            _departmentId = record["id"].ToString(),
                                            _departmentDescription = record["name"].ToString(),
                                            _departmentBossCode = sServer1
                                        });
                                    }
                                    _ProgressWork1Step(1);
                                } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                            }
                        }
                    }

                    //import users from с SCA server
                    query = "SELECT id, name, surname, patronymic, post, tabnum, parent_id, facility_code, card FROM OBJ_PERSON WHERE is_locked = '0' AND facility_code NOT LIKE '' AND card NOT LIKE '' ";
                    logger.Trace(query);
                    using (var sqlCommand = new System.Data.SqlClient.SqlCommand(query, sqlConnection))
                    {
                        using (var sqlReader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in sqlReader)
                            {
                                try
                                {
                                    if (record["name"]?.ToString()?.Trim()?.Length > 0)
                                    {
                                        row = dataTablePeople.NewRow();
                                        fio = (record["name"]?.ToString()?.Trim() + " " + record["surname"]?.ToString()?.Trim() + " " + record["patronymic"]?.ToString()?.Trim()).Replace(@"  ", @" ");
                                        groupName = record["parent_id"]?.ToString()?.Trim();
                                        nav = record["tabnum"]?.ToString()?.Trim()?.ToUpper();
                                        try
                                        {
                                            depName = departments.First((x) => x._departmentId == groupName)._departmentDescription;
                                        } catch (Exception expt) { logger.Warn(expt.Message); }

                                        row[N_ID] = Convert.ToInt32(record["id"].ToString().Trim());
                                        row[FIO] = fio;
                                        row[CODE] = nav;

                                        row[GROUP] = groupName;
                                        row[DEPARTMENT] = depName;
                                        row[DEPARTMENT_ID] = sServer1.IndexOf('.') > -1 ? sServer1.Remove(sServer1.IndexOf('.')) : sServer1;

                                        row[EMPLOYEE_POSITION] = record["post"].ToString().Trim();

                                        row[DESIRED_TIME_IN] = timeStart;
                                        row[DESIRED_TIME_OUT] = timeEnd;

                                        dataTablePeople.Rows.Add(row);

                                        listFIO.Add(new Person { FIO = fio, NAV = nav });
                                        //    listCodesWithIdCard.Add(nav);

                                        _ProgressWork1Step(1);
                                    }
                                } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                            }
                        }
                    }
                }

                // import users, shifts and group from web DB
                int tmpSeconds = 0;
                _toolStripStatusLabelSetText(StatusLabel2, "Запрашиваю данные с " + mysqlServer + ". Ждите окончания процесса...");

                groupName = mysqlServer;
                /* groups.Add(new DepartmentFull()
                 {
                     _departmentId = groupName, //group's name for staff who will have been imported from web DB
                     _departmentDescription = "Stuff from the web server",
                     _departmentBossCode = "0",
                     _departmentBossEmail = "noemail"
                 });
                 _ProgressWork1Step(1);*/

                stringConnection = @"server=" + mysqlServer + @";User=" + mysqlServerUserName + @";Password=" + mysqlServerUserPassword + @";database=wwwais;convert zero datetime=True;Connect Timeout=60"; //Allow Zero Datetime=true;
                logger.Trace(stringConnection);
                using (var sqlConnection = new MySql.Data.MySqlClient.MySqlConnection(stringConnection))
                {
                    sqlConnection.Open();

                    // import departments from web DB
                    query = "SELECT id, parent_id, name, boss_code FROM dep_struct ORDER by id";
                    logger.Trace(query);
                    using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(query, sqlConnection))
                    {
                        using (MySql.Data.MySqlClient.MySqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                if (reader?.GetString(@"name")?.Length > 0)
                                {
                                    departments.Add(new Department()
                                    {
                                        _departmentId = reader?.GetString(@"id"),
                                        _departmentDescription = reader?.GetString(@"name"),
                                        _departmentBossCode = reader?.GetString(@"boss_code")
                                    });
                                }
                                _ProgressWork1Step(1);
                            }
                        }
                    }

                    // import individual shifts of people from web DB
                    query = "Select code,start_date,mo_start,mo_end,tu_start,tu_end,we_start,we_end,th_start,th_end,fr_start,fr_end, " +
                                    "sa_start,sa_end,su_start,su_end,comment FROM work_time ORDER by start_date";
                    logger.Trace(query);
                    using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(query, sqlConnection))
                    {
                        using (MySql.Data.MySqlClient.MySqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                if (reader.GetString(@"code")?.Length > 0)
                                {

                                    try { dayStartShift = DateTimeToYYYYMMDD(reader.GetMySqlDateTime(@"start_date").ToString()); } catch
                                    { dayStartShift = DateTimeToYYYYMMDD("1980-01-01"); }

                                    peopleShifts.Add(new PeopleShift()
                                    {
                                        _nav = reader.GetString(@"code").Replace('C', 'S'),
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
                        dayStartShift = peopleShifts.FindLast((x) => x._nav == "0")._dayStartShift;
                        dayStartShift_ = "Общий график с " + dayStartShift;

                        tmpSeconds = peopleShifts.FindLast((x) => x._nav == "0" && x._dayStartShift == dayStartShift)._MoStart;
                        timeStart = ConvertSecondsTimeToStringHHMMArray(tmpSeconds)[2];

                        tmpSeconds = peopleShifts.FindLast((x) => x._nav == "0" && x._dayStartShift == dayStartShift)._MoEnd;
                        timeEnd = ConvertSecondsTimeToStringHHMMArray(tmpSeconds)[2];

                        logger.Trace("Общий график с " + dayStartShift);
                    } catch { }

                    // import people from web DB
                    query = "Select code, family_name, first_name, last_name, vacancy, department, boss_id, city FROM personal " + confitionToLoad;//where hidden=0
                    logger.Trace(query);
                    using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(query, sqlConnection))
                    {
                        using (MySql.Data.MySqlClient.MySqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                if (reader.GetString(@"family_name")?.Trim()?.Length > 0)
                                {
                                    row = dataTablePeople.NewRow();
                                    personFromServer = new PersonFull();

                                    fio = (reader.GetString(@"family_name")?.Trim() + " " + reader.GetString(@"first_name")?.Trim() + " " + reader.GetString(@"last_name")?.Trim())?.Replace(@"  ", @" ");

                                    personFromServer.FIO = fio.Replace("&acute;", "'");
                                    personFromServer.NAV = reader.GetString(@"code")?.Trim()?.ToUpper()?.Replace('C', 'S');
                                    personFromServer.DepartmentId = reader.GetString(@"department")?.Trim();

                                    depName = departments.FindLast((x) => x._departmentId == personFromServer?.DepartmentId)?._departmentDescription;
                                    personFromServer.Department = depName ?? personFromServer?.DepartmentId;

                                    depBoss = departments.Find((x) => x._departmentId == personFromServer?.DepartmentId)?._departmentBossCode;
                                    personFromServer.DepartmentBossCode = depBoss?.Length > 0 ? depBoss : reader.GetString(@"boss_id")?.Trim();

                                    personFromServer.City = reader.GetString(@"city")?.Trim();

                                    personFromServer.PositionInDepartment = reader.GetString(@"vacancy")?.Trim();
                                    personFromServer.GroupPerson = groupName;

                                    personFromServer.Shift = dayStartShift_;

                                    personFromServer.ControlInHHMM = timeStart;
                                    personFromServer.ControlOutHHMM = timeEnd;

                                    dayStartShift = peopleShifts.FindLast((x) => x._nav == personFromServer.NAV)._dayStartShift;
                                    if (dayStartShift?.Length > 0)
                                    {
                                        personFromServer.Shift = "Индивидуальный график с " + dayStartShift;

                                        tmpSeconds = peopleShifts.FindLast((x) => x._nav == personFromServer.NAV)._MoStart;
                                        personFromServer.ControlInHHMM = ConvertSecondsTimeToStringHHMMArray(tmpSeconds)[2];

                                        tmpSeconds = peopleShifts.FindLast((x) => x._nav == personFromServer.NAV)._MoEnd;
                                        personFromServer.ControlOutHHMM = ConvertSecondsTimeToStringHHMMArray(tmpSeconds)[2];

                                        personFromServer.Comment = peopleShifts.FindLast((x) => x._nav == personFromServer.NAV)._Comment;

                                        logger.Trace("Индивидуальный график с " + dayStartShift + " " + personFromServer.NAV + " " + personFromServer.Comment);
                                    }
                                    row[FIO] = personFromServer.FIO;
                                    row[CODE] = personFromServer.NAV;

                                    row[GROUP] = personFromServer.GroupPerson;
                                    row[PLACE_EMPLOYEE] = personFromServer.City;

                                    row[DEPARTMENT] = personFromServer.Department;
                                    row[DEPARTMENT_ID] = personFromServer.DepartmentId;
                                    row[EMPLOYEE_POSITION] = personFromServer.PositionInDepartment;
                                    row[CHIEF_ID] = personFromServer.DepartmentBossCode;

                                    row[EMPLOYEE_SHIFT] = personFromServer.Shift;

                                    row[DESIRED_TIME_IN] = personFromServer.ControlInHHMM;
                                    row[DESIRED_TIME_OUT] = personFromServer.ControlOutHHMM;

                                    /////////////////////
                                    // If need only people with idCard - 
                                    // should uncomment next string with "if (listCodesWithIdCard....."
                                    /////////////////////

                                    // if (listCodesWithIdCard.IndexOf(personFromServer.NAV) != -1)
                                    {
                                        dataTablePeople.Rows.Add(row);
                                    }

                                    listFIO.Add(new Person { FIO = personFromServer.FIO, NAV = personFromServer.NAV });

                                    _ProgressWork1Step(1);
                                }
                            }
                        }
                    }
                }
                dataTablePeople.AcceptChanges();
                logger.Trace("departments.count: " + departments.Count);
                _toolStripStatusLabelSetText(StatusLabel2, "ФИО и наименования департаментов получены.");
            } catch (Exception expt)
            {
                logger.Info("Возникла ошибка во время получения и обработки данных с серверов: " + expt.ToString());
                _toolStripStatusLabelSetText(StatusLabel2, "Возникла ошибка во время получения данных с серверов.");
                _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
            }

            stringConnection = null; query = null;
            //  listCodesWithIdCard = null;
            row = null;
        }

        private void WriteGroupsMailingsInLocalDb(DataTable dataTablePeople, List<PeopleShift> peopleShifts)
        {
            _toolStripStatusLabelSetText(StatusLabel2, "Формирую обновленные списоки ФИО, департаментов и рассылок...");

            logger.Trace("Приступаю к формированию списков ФИО и департаментов...");
            string query;
            string depId = "";
            string depName = "";
            string depBoss = "";
            string depDescr = "";
            string depBossEmail = "";
            List<DepartmentFull> groups = new List<DepartmentFull>();
            HashSet<Department> departmentsUniq = new HashSet<Department>();
            HashSet<DepartmentFull> departmentsEmailUniq = new HashSet<DepartmentFull>();
            _ProgressWork1Step(1);

            foreach (var dr in dataTablePeople.AsEnumerable())
            {
                //todo
                //get data 'Default_Recepient_code_From_Db' from LocalDB for webServer
                depId = dr[DEPARTMENT_ID]?.ToString();
                //                 = staffAD.Find((x) => x.code == reader.GetString(@"code")).mail

                depBossEmail = staffAD.Find((x) => x.code == dr[CHIEF_ID]?.ToString())?.mail;
                if (depId?.Length > 0)
                {
                    groups.Add(new DepartmentFull()
                    {
                        _departmentId = "@" + depId,
                        _departmentDescription = dr[DEPARTMENT]?.ToString(),
                        _departmentBossCode = dr[CHIEF_ID]?.ToString(),
                        _departmentBossEmail = depBossEmail
                    });

                    logger.Trace(groups.Count + " @" + depId + " " + dr[DEPARTMENT]?.ToString() + " " + dr[CHIEF_ID]?.ToString() + " " + depBossEmail);
                }

                depId = dr[GROUP]?.ToString();
                if (depId?.Length > 0)
                {
                    if (depId == mysqlServer)
                    {
                        groups.Add(new DepartmentFull()
                        {
                            _departmentId = depId,
                            _departmentDescription = "web",
                            _departmentBossCode = "GetCodeFromDB",
                            _departmentBossEmail = "GetEmailFromDB"
                        });
                        logger.Trace(groups.Count + " _ " + depId + " " + "web" + " " + "GetCodeFromDB" + " " + "GetEmailFromDB");
                    }
                    else
                    {
                        groups.Add(new DepartmentFull()
                        {
                            _departmentId = depId,
                            _departmentDescription = dr[DEPARTMENT]?.ToString(),
                            _departmentBossCode = "GetCodeFromDB",
                            _departmentBossEmail = "GetEmailFromDB"
                        });
                        logger.Trace(groups.Count + " _ " + depId + " " + dr[DEPARTMENT]?.ToString() + " " + "GetCodeFromDB" + " " + "GetEmailFromDB");
                    }
                }
                _ProgressWork1Step(1);
            }
            logger.Trace("groups.count: " + groups.Distinct().Count());

            foreach (var strDepartment in groups.Distinct())
            {
                if (strDepartment._departmentId?.Length > 0)
                {
                    departmentsUniq.Add(new Department
                    {
                        _departmentId = strDepartment._departmentId,
                        _departmentDescription = strDepartment._departmentDescription,
                        _departmentBossCode = strDepartment._departmentBossCode
                    });

                    if (strDepartment._departmentBossEmail?.Length > 0)
                    {
                        departmentsEmailUniq.Add(new DepartmentFull
                        {
                            _departmentId = strDepartment._departmentId,
                            _departmentDescription = strDepartment._departmentDescription,
                            _departmentBossEmail = strDepartment._departmentBossEmail
                        });
                    }
                }
            }
            _ProgressWork1Step(1);

            if (databasePerson.Exists)
            {
                logger.Trace("Чищу базу от старых списков с ФИО...");
                DeleteAllDataInTableQuery(databasePerson, "LastTakenPeopleComboList");

                foreach (var department in departmentsUniq.ToList().Distinct())
                {
                    DeleteDataTableQueryParameters(databasePerson, "PeopleGroup", "GroupPerson", department._departmentId);
                    _ProgressWork1Step(1);
                }
                _ProgressWork1Step(1);

                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();

                    var sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();

                    logger.Trace("Готовлю списки исключений из рассылок...");
                    query = "SELECT RecipientEmail FROM MailingException;";
                    logger.Trace(query);
                    string dbRecordTemp;
                    using (var sqlCommand = new SQLiteCommand(query, sqlConnection))
                    {
                        using (var reader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in reader)
                            {
                                dbRecordTemp = record["RecipientEmail"]?.ToString();

                                if (dbRecordTemp?.Length > 0)
                                {
                                    logger.Trace(dbRecordTemp);
                                    departmentsEmailUniq.RemoveWhere(x => x._departmentBossEmail == dbRecordTemp);
                                }
                            }
                        }
                    }
                    sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();

                    logger.Trace("Записываю новые группы ...");
                    sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                    foreach (var deprtment in departmentsUniq.ToList().Distinct())
                    {
                        depName = deprtment._departmentId;
                        depDescr = deprtment._departmentDescription;

                        depBoss = deprtment._departmentBossCode?.Length > 0 ? deprtment._departmentBossCode : "Default_Recepient_code_From_Db";
                        using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PeopleGroupDesciption' (GroupPerson, GroupPersonDescription, Recipient) " +
                                               "VALUES (@GroupPerson, @GroupPersonDescription, @Recipient)", sqlConnection))
                        {
                            command.Parameters.Add("@GroupPerson", DbType.String).Value = depName;
                            command.Parameters.Add("@GroupPersonDescription", DbType.String).Value = depDescr;
                            command.Parameters.Add("@Recipient", DbType.String).Value = depBoss;
                            try { command.ExecuteNonQuery(); } catch { }
                        }

                        logger.Trace("CreatedGroup: " + depName + "(" + depDescr + ")");
                        _ProgressWork1Step(1);
                    }
                    sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();

                    logger.Trace("Записываю новые рассылки по группам с учетом исключений...");
                    string recipientEmail = "";
                    sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                    foreach (var deprtment in departmentsEmailUniq.ToList().Distinct())
                    {
                        depName = deprtment._departmentId;
                        depDescr = deprtment._departmentDescription;
                        depBoss = deprtment._departmentBossCode;
                        recipientEmail = deprtment._departmentBossEmail;

                        if (depName.StartsWith("@") && recipientEmail.Contains('@'))
                        {
                            using (SQLiteCommand sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'Mailing' (SenderEmail, RecipientEmail, GroupsReport, NameReport, Description, Period, Status, DateCreated, SendingLastDate, TypeReport, DayReport)" +
                               " VALUES (@SenderEmail, @RecipientEmail, @GroupsReport, @NameReport, @Description, @Period, @Status, @DateCreated, @SendingLastDate, @TypeReport, @DayReport)", sqlConnection))
                            {
                                sqlCommand.Parameters.Add("@SenderEmail", DbType.String).Value = mailServerUserName;
                                sqlCommand.Parameters.Add("@RecipientEmail", DbType.String).Value = recipientEmail;
                                sqlCommand.Parameters.Add("@GroupsReport", DbType.String).Value = depName;
                                sqlCommand.Parameters.Add("@NameReport", DbType.String).Value = depName;
                                sqlCommand.Parameters.Add("@Description", DbType.String).Value = depDescr;
                                sqlCommand.Parameters.Add("@Period", DbType.String).Value = "Текущий месяц";
                                sqlCommand.Parameters.Add("@Status", DbType.String).Value = "Активная";
                                sqlCommand.Parameters.Add("@DateCreated", DbType.String).Value = DateTime.Now.ToYYYYMMDDHHMM();
                                sqlCommand.Parameters.Add("@SendingLastDate", DbType.String).Value = "";
                                sqlCommand.Parameters.Add("@TypeReport", DbType.String).Value = "Полный";
                                sqlCommand.Parameters.Add("@DayReport", DbType.String).Value = DEFAULT_DAY_OF_SENDING_REPORT;

                                try { sqlCommand.ExecuteNonQuery(); } catch { }
                            }

                            logger.Trace("SaveMailing: " + recipientEmail + " " + depName + " " + depDescr);
                        }
                        _ProgressWork1Step(1);
                    }
                    sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();

                    logger.Trace("Записываю новые индивидуальные графики...");
                    sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
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
                                sqlCommand.Parameters.Add("@DayInputed", DbType.String).Value = DateTime.Now.ToYYYYMMDDHHMM();
                                try { sqlCommand.ExecuteNonQuery(); } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                                _ProgressWork1Step(1);
                            }
                        }
                    }
                    sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();

                    logger.Info("Записано групп: " + departmentsUniq.ToArray().Distinct().Count());
                    logger.Info("Записано рассылок: " + departmentsEmailUniq.ToArray().Distinct().Count());
                }
            }

            departmentsUniq = null;
            departmentsEmailUniq = null;

            _ProgressWork1Step(1);
            _toolStripStatusLabelSetText(StatusLabel2, "Списки ФИО и департаментов получены.");
        }

        private void listFioItem_Click(object sender, EventArgs e) //ListFioReturn()
        {
            nameOfLastTableFromDB = "ListFIO";
            _controlEnable(comboBoxFio, true);
            SeekAndShowMembersOfGroup("");
        }

        private async void Export_Click(object sender, EventArgs e)
        {
            _ProgressBar1Start();
            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(SettingsMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            _controlEnable(dataGridView1, false);

            filePathExcelReport = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(filePathApplication), "InOut_" + DateTimeToYYYYMMDD() + @".xlsx");
            await Task.Run(() => ExportDatatableSelectedColumnsToExcel(ref dtPersonTemp, "InOutStaff", ref filePathExcelReport));
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("explorer.exe", " /select, " + filePathExcelReport)); // //System.Reflection.Assembly.GetExecutingAssembly().Location)

            _MenuItemEnabled(FunctionMenuItem, true);
            _MenuItemEnabled(SettingsMenuItem, true);
            _MenuItemEnabled(GroupsMenuItem, true);
            _controlEnable(dataGridView1, true);
            _ProgressBar1Stop();
        }

        private void ExportDatatableSelectedColumnsToExcel(ref DataTable dataTable, string nameReport, ref string filePath)  //Export DataTable to Excel 
        {
            reportExcelReady = false;
            dataTable.SetColumnsOrder(orderColumnsFinacialReport);
            DataView viewExport = new DataView(dataTable);
            viewExport.Sort = "[Фамилия Имя Отчество], [Дата регистрации] ASC";
            DataTable dtExport = viewExport.ToTable();

            logger.Trace("В таблице " + dataTable.TableName + " столбцов всего - " + dtExport.Columns.Count + ", строк - " + dtExport.Rows.Count);
            _toolStripStatusLabelSetText(StatusLabel2, "Генерирую Excel-файл по отчету: '" + nameReport + "'");
            _ProgressWork1Step(1);

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
                _ProgressWork1Step(1);

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
                Microsoft.Office.Interop.Excel.Worksheet sheet = workbook.Worksheets.get_Item(1);

                sheet.Name = nameReport;
                //sheet.Names.Add("next", "=" + Path.GetFileNameWithoutExtension(filePathExcelReport) + "!$A$1", true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _ProgressWork1Step(1);

                //colourize background of column
                //the last column
                sheet.Columns[columnsInTable].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSilver;  // последняя колонка

                //other ways to colourize background:
                //sheet.Columns[columnsInTable].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Silver);
                // or 
                //sheet.Columns[columnsInTable].Interior.Color = System.Drawing.Color.Silver;

                sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(FIO)) + 1)]
               .Interior.Color = Color.DarkSeaGreen;

                try
                {
                    sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(EMPLOYEE_BEING_LATE)) + 1)]
                    .Interior.Color = Color.SandyBrown;
                    sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(EMPLOYEE_BEING_LATE)) + 1)]
                   .Font.Color = Color.Red;
                    //.Font.Color = ColorTranslator.ToOle(Color.Red);
                    //arranges text in the center of the column
                    Microsoft.Office.Interop.Excel.Range rangeColumnA = sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(EMPLOYEE_BEING_LATE)) + 1)];
                    rangeColumnA.Cells.EntireColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(EMPLOYEE_EARLY_DEPARTURE)) + 1)]
                    .Interior.Color = Color.SandyBrown;
                    sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(EMPLOYEE_EARLY_DEPARTURE)) + 1)]
                   .Font.Color = Color.Red;
                    //arranges text in the center of the column
                    Microsoft.Office.Interop.Excel.Range rangeColumnB = sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(EMPLOYEE_EARLY_DEPARTURE)) + 1)];
                    rangeColumnB.Cells.EntireColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                } catch (Exception expt) { logger.Warn("нарушения: " + expt.ToString()); }
                _ProgressWork1Step(1);

                try
                {
                    Microsoft.Office.Interop.Excel.Range rangeColumnC = sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(@"Отпуск")) + 1)];
                    rangeColumnC.Cells.EntireColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                } catch (Exception expt) { logger.Warn("Отпуск: " + expt.ToString()); }

                try
                {
                    Microsoft.Office.Interop.Excel.Range rangeColumnD = sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(EMPLOYEE_HOOKY)) + 1)];
                    rangeColumnD.Cells.EntireColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                } catch (Exception expt) { logger.Warn("Отгул: " + expt.ToString()); }

                try
                {
                    Microsoft.Office.Interop.Excel.Range rangeColumnE = sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(EMPLOYEE_SICK_LEAVE)) + 1)];
                    rangeColumnE.Cells.EntireColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                } catch (Exception expt) { logger.Warn("Больничный: " + expt.ToString()); }

                try
                {
                    Microsoft.Office.Interop.Excel.Range rangeColumnF = sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(EMPLOYEE_ABSENCE)) + 1)];
                    rangeColumnF.Cells.EntireColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                } catch (Exception expt) { logger.Warn("Отсутствовал: " + expt.ToString()); }
                _ProgressWork1Step(1);

                //first row
                Microsoft.Office.Interop.Excel.Range rangeColumnName = sheet.Range["A1", GetExcelColumnName(columnsInTable) + 1];
                rangeColumnName.Cells.WrapText = true;
                rangeColumnName.Cells.Interior.Color = Color.Silver;
                rangeColumnName.Cells.Font.Size = 7;
                // rangeColumnName.Cells.Font.Bold = true;

                for (int column = 0; column < columnsInTable; column++)
                {
                    sheet.Cells[1, column + 1].Value = nameColumns[column];
                    sheet.Columns[column + 1].NumberFormat = "@"; // set format data of cells - text
                }
                _ProgressWork1Step(1);

                foreach (DataRow row in dtExport.Rows)
                {
                    rows++;
                    for (int column = 0; column < columnsInTable; column++)
                    {
                        sheet.Cells[rows, column + 1].Value = row[indexColumns[column]];
                    }
                    _ProgressWork1Step(1);
                }

                //colourize parts of text in the selected cell by different colors
                /*
                Microsoft.Office.Interop.Excel.Range rng = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[1, (i + 1)];
                rng.Value = "Green Red";
                rng.Characters[1, 5].Font.Color = Color.Green;
                rng.Characters[7, 9].Font.Color = Color.Red;
                */

                //get the using range           
                Microsoft.Office.Interop.Excel.Range range = sheet.UsedRange;
                //sheet.Cells.Range["A1", GetExcelColumnName(columnsInTable) + rowsInTable];
                // Microsoft.Office.Interop.Excel.Range range = sheet.Range["A2", GetExcelColumnName(columnsInTable) + (rows - 1)];

                //Шрифт для диапазона
                range.Cells.Font.Name = "Tahoma";
                //Размер шрифта
                range.Cells.Font.Size = 8;
                //ширина колонок - авто
                range.Cells.EntireColumn.AutoFit();
                _ProgressWork1Step(1);

                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                _ProgressWork1Step(1);

                //Autofilter
                range.Select();
                range.AutoFilter(1, Type.Missing, Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);

                //save document
                workbook.SaveAs(filePath,
                    Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                    true, false, //save without asking
                    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
                    Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value
                    );

                //close document
                workbook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                workbooks.Close();
                _ProgressWork1Step(1);

                //clear temporary objects
                releaseObject(range);
                releaseObject(rangeColumnName);
                releaseObject(sheet);
                releaseObject(workbook);
                releaseObject(workbooks);
                excel.Quit();
                releaseObject(excel);
                indexColumns = null;
                nameColumns = null;

                _ProgressWork1Step(1);

                _toolStripStatusLabelSetText(StatusLabel2, "Отчет сохранен в файл: " + filePath);
                _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);
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
                logger.Trace("Exception releasing object of Excel: " + ex.Message);
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
                StatusLabel2.Text = @"Выбран: " + ShortFIO(textBoxFIO.Text);
            } catch { }
            if (comboBoxFio.SelectedIndex > -1)
            {
                LoadDataItem.BackColor = Color.PaleGreen;
                groupBoxPeriod.BackColor = Color.PaleGreen;
                groupBoxTimeStart.BackColor = Color.PaleGreen;
                groupBoxTimeEnd.BackColor = Color.PaleGreen;
                groupBoxFilterReport.BackColor = SystemColors.Control;
            }
            sComboboxFIO = null;
        }

        private void CreateGroupItem_Click(object sender, EventArgs e)
        {
            if (textBoxGroup.Text.Trim().Length > 0)
            {
                CreateGroupInDB(databasePerson, textBoxGroup.Text.Trim(), textBoxGroupDescription.Text.Trim());
            }

            PersonOrGroupItem.Text = "Работать с одной персоной";
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
                }
                _toolStripStatusLabelSetText(StatusLabel2, "Группа - \"" + nameGroup + "\" создана");
            }
        }

        private void ListGroupsItem_Click(object sender, EventArgs e)
        {
            ListGroups();
        }

        private void ListGroups()
        {
            groupBoxProperties.Visible = false;
            dataGridView1.Visible = false;

            UpdateAmountAndRecepientOfPeopleGroupDesciption();

            ShowDataTableDbQuery(databasePerson, "PeopleGroupDesciption",
                "SELECT GroupPerson AS 'Группа', GroupPersonDescription AS 'Описание группы', AmountStaffInDepartment AS 'Колличество сотрудников в группе', Recipient AS 'Получатель рассылки' ", " group by GroupPerson ORDER BY GroupPerson asc; ");

            LoadDataItem.BackColor = Color.PaleGreen;
            groupBoxPeriod.BackColor = Color.PaleGreen;
            groupBoxTimeStart.BackColor = Color.PaleGreen;
            groupBoxTimeEnd.BackColor = Color.PaleGreen;
            groupBoxFilterReport.BackColor = SystemColors.Control;

            DeleteGroupItem.Visible = true;
            dataGridView1.Visible = true;
            MembersGroupItem.Enabled = true;
            _controlEnable(comboBoxFio, false);

            dataGridView1.Select();

            PersonOrGroupItem.Text = "Работать с одной персоной";
        }

        private void UpdateAmountAndRecepientOfPeopleGroupDesciption()
        {
            logger.Trace("UpdateAmountAndRecepientOfPeopleGroupDesciption");
            List<string> groupsUncount = new List<string>();
            List<AmountMembersOfGroup> amounts = new List<AmountMembersOfGroup>();
            HashSet<string> groupsUniq = new HashSet<string>();
            List<DepartmentFull> emails = new List<DepartmentFull>();
            string tmpRec = "";
            string query = "";

            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();

                    SQLiteCommand sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();

                    //set to empty for amounts and recepients in the PeopleGroupDesciption
                    query = "UPDATE 'PeopleGroupDesciption' SET AmountStaffInDepartment='0';";
                    using (var command = new SQLiteCommand(query, sqlConnection))
                    { command.ExecuteNonQuery(); }
                    query = "UPDATE 'PeopleGroupDesciption' SET Recipient='';";
                    using (var command = new SQLiteCommand(query, sqlConnection))
                    { command.ExecuteNonQuery(); }
                    logger.Trace(query);

                    sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                    query = "SELECT GroupPerson, DepartmentId FROM PeopleGroup;";
                    logger.Trace(query);
                    using (var sqlCommand = new SQLiteCommand(query, sqlConnection))
                    {
                        string idGroup = "";
                        string group = "";
                        using (var reader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in reader)
                            {
                                group = record["GroupPerson"]?.ToString();
                                idGroup = "@" + record["DepartmentId"]?.ToString();

                                if (group.Length > 0)
                                {
                                    groupsUncount.Add(group);
                                    groupsUncount.Add(idGroup);
                                }
                            }
                        }
                    }

                    ////////////////////////////////////
                    /////  ("Mailing", "SenderEmail TEXT, RecipientEmail TEXT, GroupsReport TEXT, NameReport TEXT, Description TEXT, Period TEXT, Status TEXT, SendingLastDate TEXT, TypeReport TEXT, DayReport TEXT, DateCreated TEXT", databasePerson);
                    // ("PeopleGroupDesciption", "GroupPerson TEXT, GroupPersonDescription TEXT, AmountStaffInDepartment TEXT, Recipient TEXT", databasePerson);
                    query = "SELECT GroupPerson FROM PeopleGroupDesciption;";
                    logger.Trace(query);
                    using (var sqlCommand = new SQLiteCommand(query, sqlConnection))
                    {
                        using (var reader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in reader)
                            {
                                tmpRec = record["GroupPerson"]?.ToString();

                                if (tmpRec?.Length > 0)
                                { groupsUniq.Add(tmpRec); }
                            }
                        }
                    }

                    foreach (string grp in groupsUniq)
                    {
                        tmpRec = "";
                        query = "SELECT GroupsReport, RecipientEmail FROM Mailing  WHERE GroupsReport like '" + grp + "' ;";
                        logger.Trace(query);
                        using (var sqlCommand = new SQLiteCommand(query, sqlConnection))
                        {
                            using (var reader = sqlCommand.ExecuteReader())
                            {
                                foreach (DbDataRecord record in reader)
                                {
                                    if (tmpRec.Length == 0)
                                    { tmpRec += record["RecipientEmail"]?.ToString(); }
                                    else { tmpRec += ", " + record["RecipientEmail"]?.ToString(); }
                                }
                            }
                        }
                        emails.Add(new DepartmentFull
                        {
                            _departmentId = grp,
                            _departmentBossEmail = tmpRec
                        });
                    }
                    ///////////////////
                    ///

                    logger.Trace("groupsUncount: " + (new HashSet<string>(groupsUncount)).Count());
                    foreach (string group in new HashSet<string>(groupsUncount))
                    {
                        amounts.Add(new AmountMembersOfGroup()
                        {
                            _groupName = group,
                            _amountMembers = groupsUncount.Where(x => x == group).Count(),
                            _emails = emails.Find(x => x._departmentId == group)._departmentBossEmail
                        });
                    }

                    sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                    if (amounts.Count > 0)
                    {
                        foreach (var group in amounts.ToArray())
                        {
                            query = "UPDATE 'PeopleGroupDesciption' SET AmountStaffInDepartment='" + group._amountMembers + "' WHERE GroupPerson like '" + group._groupName + "';";
                            logger.Trace(query);
                            using (var command = new SQLiteCommand(query, sqlConnection))
                            { command.ExecuteNonQuery(); }

                            query = "UPDATE 'PeopleGroupDesciption' SET Recipient='" + group._emails + "' WHERE GroupPerson like '" + group._groupName + "';";
                            logger.Trace(query);
                            using (var command = new SQLiteCommand(query, sqlConnection))
                            { command.ExecuteNonQuery(); }
                        }
                    }

                    sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                    sqlCommand1.Dispose();
                }
            }
            amounts = null; groupsUncount = null;
        }

        private void MembersGroupItem_Click(object sender, EventArgs e)//SearchMembersSelectedGroup()
        { SearchMembersSelectedGroup(); }

        private void SearchMembersSelectedGroup()
        {
            DataGridViewSeekValuesInSelectedRow dgSeek = new DataGridViewSeekValuesInSelectedRow();
            if (nameOfLastTableFromDB == "PeopleGroup" || nameOfLastTableFromDB == "PeopleGroupDesciption")
            {
                dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] { GROUP });
                SeekAndShowMembersOfGroup(dgSeek.values[0]);
            }
            else if (nameOfLastTableFromDB == "Mailing")
            {
                dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] { @"Отчет по группам" });
                SeekAndShowMembersOfGroup(dgSeek.values[0]);
            }
            dgSeek = null;
        }

        private void SeekAndShowMembersOfGroup(string nameGroup)
        {
            //  dtPeopleListLoaded?.Dispose();
            //  dtPeopleListLoaded = dtPeople.Clone();
            var dtTemp = dtPeople.Clone();

            numberPeopleInLoading = 0;
            DataRow dataRow;
            string dprtmnt = "", query = ""; ;

            query = "Select FIO, NAV, GroupPerson, ControllingHHMM, ControllingOUTHHMM, Shift, Comment, Department, PositionInDepartment, DepartmentId, City, Boss FROM PeopleGroup ";
            if (nameGroup.Contains("@"))
            { query += " where DepartmentId like '" + nameGroup.Remove(0, 1) + "'"; }
            else if (nameGroup.Length > 0)
            { query += " where GroupPerson like '" + nameGroup + "'"; }

            query += ";";
            logger.Trace("SeekAndShowMembersOfGroup: " + query);
            using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
            {
                sqlConnection.Open();
                using (var sqlCommand = new SQLiteCommand(query, sqlConnection))
                {
                    using (var sqlReader = sqlCommand.ExecuteReader())
                    {
                        foreach (DbDataRecord record in sqlReader)
                        {
                            if (record["FIO"]?.ToString()?.Length > 0 && record["NAV"]?.ToString()?.Length > 0)
                            {
                                try { dprtmnt = record[@"Department"].ToString(); } catch { dprtmnt = record[@"GroupPerson"]?.ToString(); }

                                dataRow = dtTemp.NewRow();
                                dataRow[FIO] = record[@"FIO"].ToString();
                                dataRow[CODE] = record[@"NAV"].ToString();

                                dataRow[GROUP] = record[@"GroupPerson"]?.ToString();
                                dataRow[DEPARTMENT] = dprtmnt;
                                dataRow[DEPARTMENT_ID] = record[@"DepartmentId"]?.ToString();
                                dataRow[EMPLOYEE_POSITION] = record[@"PositionInDepartment"]?.ToString();
                                dataRow[PLACE_EMPLOYEE] = record[@"City"]?.ToString();
                                dataRow[CHIEF_ID] = record[@"Boss"]?.ToString();

                                dataRow[DESIRED_TIME_IN] = record[@"ControllingHHMM"]?.ToString();
                                dataRow[DESIRED_TIME_OUT] = record[@"ControllingOUTHHMM"]?.ToString();

                                dataRow[EMPLOYEE_SHIFT_COMMENT] = record["Comment"]?.ToString();
                                dataRow[EMPLOYEE_SHIFT] = record[@"Shift"]?.ToString();

                                dtTemp.Rows.Add(dataRow);
                                numberPeopleInLoading++;
                            }
                        }
                    }
                }
            }

            if (numberPeopleInLoading > 0)
            {
                var namesDistinctCollumnsArray = arrayAllColumnsDataTablePeople.Except(arrayHiddenColumnsFIO).ToArray(); //take distinct data
                dtPersonTemp = GetDistinctRecords(dtTemp, namesDistinctCollumnsArray);
                ShowDatatableOnDatagridview(dtPersonTemp, arrayHiddenColumnsFIO, "PeopleGroup");
                _MenuItemVisible(DeletePersonFromGroupItem, true);
            }

            query = null;
            dataRow = null;
            dtTemp.Dispose();
        }


        private void importPeopleInLocalDBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dtPeopleListLoaded?.Clear();
            dtPeopleListLoaded = dtPeople.Copy();
            HashSet<Department> departmentsUniq = new HashSet<Department>();

            ImportTextToTable(dtPeopleListLoaded, ref departmentsUniq);
            WritePeopleInLocalDB(databasePerson.ToString(), dtPeopleListLoaded);
            ImportListGroupsDescriptionInLocalDB(databasePerson.ToString(), departmentsUniq);
            departmentsUniq = null;
        }

        private void ImportTextToTable(DataTable dt, ref HashSet<Department> departmentsUniq) //Fill dtPeople
        {
            List<string> listRows = LoadDataIntoList();

            string checkHourS;
            string checkHourE;

            string getThreeRows = "";
            getThreeRows = "Маска:\nФИО\tNAV-код\tГруппа\tОтдел\tДолжность\tВремя прихода,часы\tВремя прихода,минуты\tВремя ухода,часы\tВремя ухода,минуты\n\nДанные:\n";
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
                            row[FIO] = cell[0];
                            row[CODE] = cell[1];
                            row[GROUP] = cell[2];
                            row[DEPARTMENT] = cell[3];
                            row[DEPARTMENT_ID] = "";
                            row[EMPLOYEE_POSITION] = cell[4];

                            departmentsUniq.Add(new Department
                            {
                                _departmentId = cell[2],
                                _departmentDescription = cell[3],
                            });

                            checkHourS = cell[5];
                            if (TryParseStringToDecimal(checkHourS) == 0)
                            { checkHourS = numUpHourStart.ToString(); }
                            row[DESIRED_TIME_IN] = ConvertStringsTimeToStringHHMM(checkHourS, cell[6]);

                            checkHourE = cell[7];
                            if (TryParseStringToDecimal(checkHourE) == 0)
                            { checkHourE = numUpDownHourEnd.ToString(); }
                            row[DESIRED_TIME_OUT] = ConvertStringsTimeToStringHHMM(checkHourE, cell[8]);

                            dt.Rows.Add(row);
                            row = dt.NewRow();
                        }
                    }
                }
            }
            else
            { MessageBox.Show("выбранный файл пустой, или \nне подходит для импорта."); }
        }

        private void WritePeopleInLocalDB(string pathToPersonDB, DataTable dtSource) //use listGroups /add reserv1 reserv2
        {
            using (var sqlConnection = new SQLiteConnection($"Data Source={pathToPersonDB};Version=3;"))
            {
                sqlConnection.Open();
                logger.Trace("WritePeopleInLocalDB: " + dtSource + " " + dtSource.Rows.Count);
                //Write people in local DB
                SQLiteCommand sqlCommandTransaction = new SQLiteCommand("begin", sqlConnection);
                sqlCommandTransaction.ExecuteNonQuery();
                foreach (var dr in dtSource.AsEnumerable())
                {
                    if (dr[FIO]?.ToString()?.Length > 0 && dr[CODE]?.ToString()?.Length > 0)
                    {
                        using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'PeopleGroup' (FIO, NAV, GroupPerson, ControllingHHMM, ControllingOUTHHMM, Shift, Comment, Department, PositionInDepartment, DepartmentId, City, Boss) " +
                                " VALUES (@FIO, @NAV, @GroupPerson, @ControllingHHMM, @ControllingOUTHHMM, @Shift, @Comment, @Department, @PositionInDepartment, @DepartmentId, @City, @Boss)", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@FIO", DbType.String).Value = dr[FIO]?.ToString();
                            sqlCommand.Parameters.Add("@NAV", DbType.String).Value = dr[CODE]?.ToString();

                            sqlCommand.Parameters.Add("@GroupPerson", DbType.String).Value = dr[GROUP]?.ToString();
                            sqlCommand.Parameters.Add("@Department", DbType.String).Value = dr[DEPARTMENT]?.ToString();
                            sqlCommand.Parameters.Add("@DepartmentId", DbType.String).Value = dr[DEPARTMENT_ID].ToString();
                            sqlCommand.Parameters.Add("@PositionInDepartment", DbType.String).Value = dr[EMPLOYEE_POSITION]?.ToString();
                            sqlCommand.Parameters.Add("@City", DbType.String).Value = dr[PLACE_EMPLOYEE]?.ToString();
                            sqlCommand.Parameters.Add("@Boss", DbType.String).Value = dr[CHIEF_ID]?.ToString();

                            sqlCommand.Parameters.Add("@ControllingHHMM", DbType.String).Value = dr[DESIRED_TIME_IN]?.ToString();
                            sqlCommand.Parameters.Add("@ControllingOUTHHMM", DbType.String).Value = dr[DESIRED_TIME_OUT]?.ToString();

                            sqlCommand.Parameters.Add("@Shift", DbType.String).Value = dr[EMPLOYEE_SHIFT]?.ToString();
                            sqlCommand.Parameters.Add("@Comment", DbType.String).Value = dr[EMPLOYEE_SHIFT_COMMENT]?.ToString();

                            try
                            { sqlCommand.ExecuteNonQuery(); } catch (Exception expt)
                            {
                                logger.Info("ImportTablePeopleToTableGroupsInLocalDB: ошибка записи в базу: " + dr[FIO] + "\n" + dr[CODE] + "\n" + expt.ToString());
                            }
                        }
                    }
                }

                sqlCommandTransaction = new SQLiteCommand("end", sqlConnection);
                sqlCommandTransaction.ExecuteNonQuery();


                logger.Trace("Записываю списки в базу и контролы...");
                sqlCommandTransaction = new SQLiteCommand("begin", sqlConnection);
                sqlCommandTransaction.ExecuteNonQuery();
                foreach (var str in listFIO)
                {
                    if (str.FIO?.Length > 0 && str.NAV?.Length > 0)
                    {
                        using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'LastTakenPeopleComboList' (ComboList) VALUES (@ComboList)", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@ComboList", DbType.String).Value = str.FIO + "|" + str.NAV;
                            try { sqlCommand.ExecuteNonQuery(); } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                            _ProgressWork1Step(1);
                        }
                    }
                }
                sqlCommandTransaction = new SQLiteCommand("end", sqlConnection);
                sqlCommandTransaction.ExecuteNonQuery();
                sqlCommandTransaction?.Dispose();
                sqlConnection.Close();

                foreach (var str in listFIO)
                { _comboBoxAdd(comboBoxFio, str.FIO + "|" + str.NAV); }
                _ProgressWork1Step(1);
                if (_comboBoxCountItems(comboBoxFio) > 0)
                { _comboBoxSelectIndex(comboBoxFio, 0); }
                _ProgressWork1Step(1);
                logger.Info("Записано ФИО: " + listFIO.Count);
            }
        }

        private void ImportListGroupsDescriptionInLocalDB(string pathToPersonDB, HashSet<Department> departmentsUniq) //use listGroups
        {
            using (var connection = new SQLiteConnection($"Data Source={pathToPersonDB};Version=3;"))
            {
                connection.Open();
                SQLiteCommand commandTransaction = new SQLiteCommand("begin", connection);
                commandTransaction.ExecuteNonQuery();
                using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PeopleGroupDesciption' (GroupPerson, GroupPersonDescription, Recipient) " +
                                        "VALUES (@GroupPerson, @GroupPersonDescription ,@Recipient)", connection))
                {
                    foreach (var group in departmentsUniq)
                    {
                        command.Parameters.Add("@GroupPerson", DbType.String).Value = group._departmentId;
                        command.Parameters.Add("@GroupPersonDescription", DbType.String).Value = group._departmentDescription;
                        command.Parameters.Add("@Recipient", DbType.String).Value = group._departmentBossCode;
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
            DataGridViewSeekValuesInSelectedRow dgSeek;
            if (_dataGridView1CurrentRowIndex() > -1)
            {
                dgSeek = new DataGridViewSeekValuesInSelectedRow();
                dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] {
                 FIO, CODE, DEPARTMENT, EMPLOYEE_POSITION,
                 DESIRED_TIME_IN, DESIRED_TIME_OUT,
                 CHIEF_ID, EMPLOYEE_SHIFT, DEPARTMENT_ID, PLACE_EMPLOYEE
                            });

                using (var connection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    connection.Open();
                    if (group?.Length > 0)
                    {
                        using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PeopleGroupDesciption' (GroupPerson) " +
                                                "VALUES (@GroupPerson )", connection))
                        {
                            command.Parameters.Add("@GroupPerson", DbType.String).Value = group;
                            try { command.ExecuteNonQuery(); } catch { }
                        }
                    }

                    if (group?.Length > 0 && dgSeek.values[1]?.Length > 0)
                    {
                        using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'PeopleGroup' (FIO, NAV, GroupPerson, ControllingHHMM, ControllingOUTHHMM, Department, PositionInDepartment, Comment, Shift, DepartmentId, City, Boss) " +
                                                    "VALUES (@FIO, @NAV, @GroupPerson, @ControllingHHMM, @ControllingOUTHHMM, @Department, @PositionInDepartment, @Comment, @Shift, @DepartmentId, @City, @Boss)", connection))
                        {
                            sqlCommand.Parameters.Add("@FIO", DbType.String).Value = dgSeek.values[0];
                            sqlCommand.Parameters.Add("@NAV", DbType.String).Value = dgSeek.values[1];
                            sqlCommand.Parameters.Add("@GroupPerson", DbType.String).Value = group;
                            sqlCommand.Parameters.Add("@Department", DbType.String).Value = dgSeek.values[2];
                            sqlCommand.Parameters.Add("@DepartmentId", DbType.String).Value = dgSeek.values[8];
                            sqlCommand.Parameters.Add("@PositionInDepartment", DbType.String).Value = dgSeek.values[3];
                            sqlCommand.Parameters.Add("@City", DbType.String).Value = dgSeek.values[9];
                            sqlCommand.Parameters.Add("@Boss", DbType.String).Value = dgSeek.values[6];
                            sqlCommand.Parameters.Add("@Shift", DbType.String).Value = dgSeek.values[7];
                            sqlCommand.Parameters.Add("@ControllingHHMM", DbType.String).Value = dgSeek.values[4];
                            sqlCommand.Parameters.Add("@ControllingOUTHHMM", DbType.String).Value = dgSeek.values[5];
                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }
                        StatusLabel2.Text = "\"" + ShortFIO(dgSeek.values[0]) + "\"" + " добавлен в группу \"" + group + "\"";
                        _toolStripStatusLabelBackColor(StatusLabel2, SystemColors.Control);
                    }
                    else if (group?.Length > 0 && dgSeek?.values[1]?.Length == 0)
                    {
                        StatusLabel2.Text = "Отсутствует NAV-код у: " + ShortFIO(textBoxFIO.Text);
                        _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                    }
                    else if (group?.Length == 0 && dgSeek?.values[1]?.Length > 0)
                    {
                        StatusLabel2.Text = "Не указана группа, в которую нужно добавить!";
                        _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                    }
                }

            }
            SeekAndShowMembersOfGroup(group);

            labelGroup.BackColor = SystemColors.Control;
            PersonOrGroupItem.Text = "Работать с одной персоной";
            nameOfLastTableFromDB = "PeopleGroup";
            group = groupDescription = null; dgSeek = null;
        }

        private void GetNamePoints() //Get names of the pass by points
        {
            passByPoints = new List<PassByPoint>();
            if (databasePerson.Exists)
            {
                string stringConnection = @"Data Source=" + sServer1 + @"\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID=" + sServer1UserName + @";Password=" + sServer1UserPassword + @"; Connect Timeout=60";
                string query;
                string name;
                string direction;
                using (var sqlConnection = new System.Data.SqlClient.SqlConnection(stringConnection))
                {
                    sqlConnection.Open();
                    query = "Select id, name FROM OBJ_ABC_ARC_READER;";
                    using (var sqlCommand = new System.Data.SqlClient.SqlCommand(query, sqlConnection))
                    {
                        using (var sqlReader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in sqlReader)
                            {
                                if (record["id"]?.ToString()?.Trim()?.Length > 0 && record["name"]?.ToString()?.Trim()?.Length > 0)
                                {
                                    name = record["name"].ToString().Trim();
                                    if (name.Contains("выход"))
                                    { direction = "Выход"; }
                                    else
                                    { direction = "Вход"; }

                                    passByPoints.Add(new PassByPoint
                                    {
                                        _id = record["id"].ToString().Trim(),
                                        _server = sServer1,
                                        _name = name,
                                        _direction = direction
                                    });
                                }
                            }
                        }
                    }
                }
                stringConnection = query = name = direction = null;
            }
        }

        private async void GetDataItem_Click(object sender, EventArgs e) //GetData()
        {
            _ProgressBar1Start();

            await Task.Run(() => CheckAliveIntellectServer(sServer1, sServer1UserName, sServer1UserPassword));

            _changeControlBackColor(groupBoxPeriod, SystemColors.Control);
            _changeControlBackColor(groupBoxTimeStart, SystemColors.Control);
            _changeControlBackColor(groupBoxTimeEnd, SystemColors.Control);
            _MenuItemBackColorChange(LoadDataItem, SystemColors.Control);

            _MenuItemEnabled(LoadDataItem, false);
            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(SettingsMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            CheckBoxesFiltersAll_Enable(false);

            _controlVisible(dataGridView1, false);

            if (bServer1Exist)
            {
                await Task.Run(() => GetData());

                _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);
                _MenuItemBackColorChange(LoadDataItem, SystemColors.Control);
                _MenuItemBackColorChange(TableExportToExcelItem, Color.PaleGreen);

                _MenuItemEnabled(LoadDataItem, true);
                _MenuItemEnabled(FunctionMenuItem, true);
                _MenuItemEnabled(SettingsMenuItem, true);
                _MenuItemEnabled(GroupsMenuItem, true);
                _MenuItemEnabled(VisualModeItem, true);
                _MenuItemVisible(VisualModeItem, true);
                _MenuItemEnabled(TableModeItem, true);
                _MenuItemEnabled(TableExportToExcelItem, true);

                _controlVisible(dataGridView1, true);

                _controlEnable(checkBoxReEnter, true);
                _controlEnable(checkBoxTimeViolations, false);
                _controlEnable(checkBoxWeekend, false);
                _controlEnable(checkBoxCelebrate, false);
                CheckBoxesFiltersAll_CheckedState(false);

                panelViewResize(numberPeopleInLoading);
                _changeControlBackColor(groupBoxFilterReport, Color.PaleGreen);
            }
            else
            {
                GetInfoSetup();
                _MenuItemEnabled(SettingsMenuItem, true);
            }
            _ProgressBar1Stop();

            if (dtPersonTemp?.Rows.Count > 0)
            {
                _MenuItemVisible(TableExportToExcelItem, true);
                _toolStripStatusLabelSetText(StatusLabel2, "Данные регистраций пропусков персонала загружены");
            }
        }

        private void GetData()
        {
            string group = _textBoxReturnText(textBoxGroup);

            //Clear work tables
            dtPersonRegistrationsFullList.Clear();
            reportStartDay = _dateTimePickerStart().Split(' ')[0];
            reportLastDay = _dateTimePickerEnd().Split(' ')[0];
            logger.Trace("GetData: " + group + " " + reportStartDay + " " + reportLastDay);

            GetNamePoints();  //Get names of the points
            GetRegistrations(group, _dateTimePickerStart(), _dateTimePickerEnd(), "");

            dtPersonTemp = dtPersonRegistrationsFullList.Clone();
            dtPersonTempAllColumns = dtPersonRegistrationsFullList.Copy(); //store all columns

            var namesDistinctColumnsArray = arrayAllColumnsDataTablePeople.Except(nameHidenColumnsArray).ToArray(); //take distinct data
            dtPersonTemp = GetDistinctRecords(dtPersonTempAllColumns, namesDistinctColumnsArray);
            ShowDatatableOnDatagridview(dtPersonTemp, nameHidenColumnsArray, "PeopleGroup");

            namesDistinctColumnsArray = null;
        }

        private void GetRegistrations(string nameGroup, string startDate, string endDate, string doPostAction)
        {
            PersonFull person = new PersonFull();

            outPerson = new List<OutPerson>();

            outResons = new List<OutReasons>
            {
                new OutReasons()
                { _id = "0", _hourly = 1, _name = @"Unknow", _visibleName = @"Неизвестная" }
            };

            string query = "";
            string stringConnection = @"server=" + mysqlServer + @";User=" + mysqlServerUserName + @";Password=" + mysqlServerUserPassword + @";database=wwwais;pooling = false; convert zero datetime=True;Connect Timeout=60";
            logger.Trace(stringConnection);
            using (var sqlConnection = new MySql.Data.MySqlClient.MySqlConnection(stringConnection))
            {
                sqlConnection.Open();
                query = "Select id, name,hourly, visibled_name FROM out_reasons";
                logger.Trace(query);

                using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(query, sqlConnection))
                {
                    using (MySql.Data.MySqlClient.MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (reader?.GetString(@"id")?.Length > 0 && reader?.GetString(@"name")?.Length > 0)
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
                query = "Select * FROM out_users where reason_date >= '" + startDate.Split(' ')[0] + "' AND reason_date <= '" + endDate.Split(' ')[0] + "' ";
                logger.Trace(query);
                using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(query, sqlConnection))
                {
                    using (MySql.Data.MySqlClient.MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (reader?.GetString(@"reason_id")?.Length > 0 && reader?.GetString(@"user_code")?.Length > 0)
                            {
                                resonId = outResons.Find((x) => x._id == reader.GetString(@"reason_id"))._id;
                                try { date = DateTimeToYYYYMMDD(reader.GetString(@"reason_date")); } catch { date = ""; }

                                outPerson.Add(new OutPerson()
                                {
                                    _reason_id = resonId,
                                    _nav = reader.GetString(@"user_code").ToUpper()?.Replace('C', 'S'),
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
                logger.Trace("Всего с " + startDate.Split(' ')[0] + " по " + endDate.Split(' ')[0] + " на сайте есть - " + outPerson.Count + " записей с отсутствиями");
            }
            _ProgressWork1Step(1);

            if ((nameOfLastTableFromDB == "PeopleGroupDesciption" || nameOfLastTableFromDB == "PeopleGroup" || nameOfLastTableFromDB == "Mailing" ||
                nameOfLastTableFromDB == "ListFIO" || doPostAction == "sendEmail") && nameGroup.Length > 0)
            {
                _toolStripStatusLabelSetText(StatusLabel2, "Получаю данные по группе " + nameGroup);
                //  if (doPostAction != "sendEmail")
                {
                    dtPeopleGroup.Clear();
                    LoadGroupMembersFromDbToDataTable(nameGroup, ref dtPeopleGroup);
                }

                logger.Trace("GetRegistrations, DT - " + dtPeopleGroup.TableName + " , всего записей - " + dtPeopleGroup.Rows.Count);
                foreach (DataRow row in dtPeopleGroup.Rows)
                {
                    if (row[FIO]?.ToString()?.Length > 0 && (row[GROUP]?.ToString() == nameGroup || ("@" + row[DEPARTMENT_ID]?.ToString()) == nameGroup))
                    {
                        person = new PersonFull()
                        {
                            FIO = row[FIO].ToString(),
                            NAV = row[CODE]?.ToString(),
                            GroupPerson = nameGroup,
                            Department = row[DEPARTMENT]?.ToString(),
                            City = row[PLACE_EMPLOYEE]?.ToString(),
                            DepartmentBossCode = row[CHIEF_ID]?.ToString(),
                            PositionInDepartment = row[EMPLOYEE_POSITION]?.ToString(),
                            DepartmentId = row[DEPARTMENT_ID]?.ToString(),
                            ControlInHHMM = row[DESIRED_TIME_IN]?.ToString(),
                            ControlOutHHMM = row[DESIRED_TIME_OUT]?.ToString(),
                            Comment = row[EMPLOYEE_SHIFT_COMMENT]?.ToString(),
                            Shift = row[EMPLOYEE_SHIFT]?.ToString()
                        };

                        person.ControlInSeconds = ConvertStringTimeHHMMToSeconds(person.ControlInHHMM);
                        person.ControlOutSeconds = ConvertStringTimeHHMMToSeconds(person.ControlOutHHMM);

                        GetPersonRegistrationFromServer(ref dtPersonRegistrationsFullList, person, startDate, endDate);     //Search Registration at checkpoints of the selected person
                    }
                }
                nameOfLastTableFromDB = "PeopleGroup";
                _toolStripStatusLabelSetText(StatusLabel2, "Данные по группе \"" + nameGroup + "\" получены");
            }
            else
            {
                person = new PersonFull();
                person.NAV = _textBoxReturnText(textBoxNav);
                person.FIO = _textBoxReturnText(textBoxFIO);

                _toolStripStatusLabelSetText(StatusLabel2, "Получаю данные по \"" + ShortFIO(person.FIO) + "\" ");

                person.GroupPerson = "One Person";
                person.Department = "";
                person.DepartmentId = "";
                person.City = "";
                person.DepartmentBossCode = "";
                person.PositionInDepartment = "Сотрудник";

                person.Shift = "";
                person.Comment = "";

                person.ControlInHHMM = ConvertStringsTimeToStringHHMM(_numUpDownReturn(numUpDownHourStart).ToString(), _numUpDownReturn(numUpDownMinuteStart).ToString());
                person.ControlOutHHMM = ConvertStringsTimeToStringHHMM(_numUpDownReturn(numUpDownHourEnd).ToString(), _numUpDownReturn(numUpDownMinuteEnd).ToString());

                logger.Trace("GetRegistrations,One Person: " + person.FIO);

                GetPersonRegistrationFromServer(ref dtPersonRegistrationsFullList, person, startDate, endDate);

                _toolStripStatusLabelSetText(StatusLabel2, "Данные с СКД по \"" + ShortFIO(_textBoxReturnText(textBoxFIO)) + "\" получены!");
            }

            query = null; stringConnection = null;
        }

        private void GetPersonRegistrationFromServer(ref DataTable dtTarget, PersonFull person, string startDay, string endDay)
        {
            logger.Trace("GetPersonRegistrationFromServer, person - " + person.NAV);

            SeekAnualDays(ref dtTarget, ref person, false,
                ConvertStringDateToIntArray(startDay), ConvertStringDateToIntArray(endDay),
                ref myBoldedDates, ref workSelectedDays);
            DataRow rowPerson;
            string stringConnection = "";
            string query = "";
            HashSet<string> personWorkedDays = new HashSet<string>();
            string stringIdCardIntellect = "";

            string[] cellData;
            string namePoint = "";
            string direction = "";

            string stringDataNew = "";
            string fullPointName = "";
            int seconds = 0;

            //is looking for a NAV and an idCard
            try
            {
                stringConnection = @"Data Source=" + sServer1 + @"\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID=" + sServer1UserName + @";Password=" + sServer1UserPassword + @";Connect Timeout=240";
                using (var sqlConnection = new System.Data.SqlClient.SqlConnection(stringConnection))
                {
                    sqlConnection.Open();

                    query = "Select id, tabnum FROM OBJ_PERSON where tabnum like '" + person.NAV + "';";
                    using (var cmd = new System.Data.SqlClient.SqlCommand(query, sqlConnection))
                    {
                        using (var reader = cmd.ExecuteReader())
                        {
                            foreach (DbDataRecord record in reader)
                            {
                                if (record?["tabnum"]?.ToString()?.Trim() == person.NAV)
                                {
                                    stringIdCardIntellect = record["id"].ToString().Trim();
                                    person.idCard = Convert.ToInt32(record["id"].ToString().Trim());
                                    break;
                                }
                                _ProgressWork1Step(1);
                            }
                        }
                    }

                    // PassByPoint
                    query = "Select id, tabnum FROM OBJ_PERSON where tabnum like '" + person.NAV + "';";
                    using (var cmd = new System.Data.SqlClient.SqlCommand(query, sqlConnection))
                    {
                        using (var reader = cmd.ExecuteReader())
                        {
                            foreach (DbDataRecord record in reader)
                            {
                                if (record?["tabnum"]?.ToString()?.Trim() == person.NAV)
                                {
                                    stringIdCardIntellect = record["id"].ToString().Trim();
                                    person.idCard = Convert.ToInt32(record["id"].ToString().Trim());
                                    break;
                                }
                                _ProgressWork1Step(1);
                            }
                        }
                    }

                    query = "SELECT param0, param1, objid, objtype, CONVERT(varchar, date, 120) AS date, CONVERT(varchar, PROTOCOL.time, 114) AS time FROM protocol " +
                       " where objtype like 'ABC_ARC_READER' AND param1 like '" + stringIdCardIntellect + "' AND date >= '" + startDay + "' AND date <= '" + endDay + "' " +
                       " ORDER BY date ASC";
                    using (var cmd = new System.Data.SqlClient.SqlCommand(query, sqlConnection))
                    {
                        using (var reader = cmd.ExecuteReader())
                        {
                            foreach (DbDataRecord record in reader)
                            {
                                try
                                {
                                    if (record["param0"]?.ToString()?.Trim()?.Length > 0 && record["param1"]?.ToString()?.Length > 0)
                                    {
                                        logger.Trace(person.NAV);
                                        stringDataNew = record["date"]?.ToString()?.Trim()?.Split(' ')[0];
                                        person.idCard = Convert.ToInt32(record["param1"].ToString().Trim());
                                        seconds = ConvertStringTimeHHMMToSeconds(record["time"]?.ToString()?.Trim());
                                        fullPointName = record["objid"]?.ToString()?.Trim();

                                        namePoint = passByPoints.Find((x) => x._id == fullPointName)._name;
                                        direction = passByPoints.Find((x) => x._id == fullPointName)._direction;

                                        personWorkedDays.Add(stringDataNew);

                                        rowPerson = dtTarget.NewRow();
                                        rowPerson[FIO] = record["param0"].ToString().Trim();
                                        rowPerson[CODE] = person.NAV;
                                        rowPerson[N_ID] = record["param1"].ToString().Trim();
                                        rowPerson[GROUP] = person.GroupPerson;
                                        rowPerson[DEPARTMENT] = person.Department;
                                        rowPerson[EMPLOYEE_POSITION] = person.PositionInDepartment;
                                        rowPerson[PLACE_EMPLOYEE] = person.City;
                                        rowPerson[EMPLOYEE_SHIFT] = person.Shift;

                                        //day of registration. real data
                                        rowPerson[DATE_REGISTRATION] = stringDataNew;
                                        rowPerson[TIME_REGISTRATION] = seconds;

                                        rowPerson[SERVER_SKD] = sServer1;
                                        rowPerson[NAME_CHECKPOINT] = namePoint;
                                        rowPerson[DIRECTION_WAY] = direction;

                                        rowPerson[DESIRED_TIME_IN] = person.ControlInHHMM;
                                        rowPerson[DESIRED_TIME_OUT] = person.ControlOutHHMM;
                                        rowPerson[REAL_TIME_IN] = ConvertSecondsToStringHHMM(seconds);

                                        dtTarget.Rows.Add(rowPerson);

                                        _ProgressWork1Step(1);
                                    }
                                } catch (Exception expt) { logger.Warn("GetPersonRegistrationFromServer " + expt.ToString()); }
                            }
                        }
                    }
                }

                stringDataNew = null; query = null; stringConnection = null;
                _ProgressWork1Step(1);
            } catch (Exception Expt)
            { MessageBox.Show(Expt.ToString(), @"Сервер не доступен, или неправильная авторизация", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            // рабочие дни в которые отсутствовал данная персона
            foreach (string day in workSelectedDays.Except(personWorkedDays))
            {
                rowPerson = dtTarget.NewRow();
                rowPerson[FIO] = person.FIO;
                rowPerson[CODE] = person.NAV;
                rowPerson[GROUP] = person.GroupPerson;
                rowPerson[N_ID] = person.idCard;

                rowPerson[DEPARTMENT] = person.Department;
                rowPerson[EMPLOYEE_POSITION] = person.PositionInDepartment;
                rowPerson[PLACE_EMPLOYEE] = person.City;

                rowPerson[EMPLOYEE_SHIFT] = person.Shift;

                rowPerson[TIME_REGISTRATION] = "0"; //must be "0"!!!!
                rowPerson[DATE_REGISTRATION] = day;
                rowPerson[DAY_OF_WEEK] = DayOfWeekRussian((DateTime.Parse(day)).DayOfWeek.ToString());
                rowPerson[DESIRED_TIME_IN] = person.ControlInHHMM;
                rowPerson[DESIRED_TIME_OUT] = person.ControlOutHHMM;
                rowPerson[EMPLOYEE_ABSENCE] = "1";

                dtTarget.Rows.Add(rowPerson);//добавляем рабочий день в который  сотрудник не выходил на работу
                _ProgressWork1Step(1);
            }
            dtTarget.AcceptChanges();

            foreach (DataRow row in dtTarget.Rows)
            {
                if (row[CODE].ToString() == person.NAV)
                {
                    row[FIO] = person.FIO;
                    row[GROUP] = person.GroupPerson;
                    row[N_ID] = person.idCard;
                    row[DEPARTMENT] = person.Department;
                    row[PLACE_EMPLOYEE] = person.City;
                    row[EMPLOYEE_POSITION] = person.PositionInDepartment;
                    row[EMPLOYEE_SHIFT] = person.Shift;
                }
            }
            dtTarget.AcceptChanges();

            foreach (DataRow row in dtTarget.Rows) // search whole table
            {
                foreach (string day in workSelectedDays)
                {
                    if (row[DATE_REGISTRATION]?.ToString() == day && row[CODE]?.ToString() == person.NAV)
                    {
                        try
                        {
                            row[EMPLOYEE_SHIFT_COMMENT] = outPerson.Find((x) => x._date == day && x._nav == person.NAV)._reason_id;
                            logger.Trace("GetPersonRegistrationFromServer, outPerson " + person.NAV + ", outReason - " + row[EMPLOYEE_SHIFT_COMMENT].ToString());
                        } catch { }
                        break;
                    }
                }
            }
            dtTarget.AcceptChanges();

            _ProgressWork1Step(1);

            rowPerson = null;
            namePoint = null; direction = null;
            stringIdCardIntellect = null; cellData = new string[1];
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
        private void LoadGroupMembersFromDbToDataTable(string namePointedGroup, ref DataTable peopleGroup) // dtPeopleGroup //"Select * FROM PeopleGroup where GroupPerson like '" + _textBoxReturnText(textBoxGroup) + "';"
        {
            DataRow dataRow;

            string query = "Select FIO, NAV, GroupPerson, ControllingHHMM, ControllingOUTHHMM, Shift, Comment, Department, PositionInDepartment, DepartmentId, City, Boss  FROM PeopleGroup ";
            if (namePointedGroup.StartsWith(@"@"))
            { query += "where DepartmentId like '" + namePointedGroup.Remove(0, 1) + "'; "; }
            else if (namePointedGroup.Length > 0)
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
                                if (record[@"FIO"]?.ToString()?.Length > 0 && record[@"NAV"]?.ToString()?.Length > 0)
                                {
                                    dataRow = peopleGroup.NewRow();

                                    dataRow[FIO] = record[@"FIO"].ToString();
                                    dataRow[CODE] = record[@"NAV"].ToString();

                                    dataRow[GROUP] = record[@"GroupPerson"].ToString();
                                    dataRow[DEPARTMENT] = record[@"Department"].ToString();
                                    dataRow[DEPARTMENT_ID] = record[@"DepartmentId"].ToString();
                                    dataRow[EMPLOYEE_POSITION] = record[@"PositionInDepartment"].ToString();
                                    dataRow[PLACE_EMPLOYEE] = record[@"City"].ToString();
                                    dataRow[EMPLOYEE_SHIFT_COMMENT] = record[@"Comment"].ToString();
                                    dataRow[EMPLOYEE_SHIFT] = record[@"Shift"].ToString();

                                    dataRow[DESIRED_TIME_IN] = record[@"ControllingHHMM"].ToString();
                                    dataRow[DESIRED_TIME_OUT] = record[@"ControllingOUTHHMM"].ToString();

                                    peopleGroup.Rows.Add(dataRow);
                                }
                            } catch { }
                        }
                    }
                }
            }
            query = null; dataRow = null;

            logger.Trace("LoadGroupMembersFromDbToDataTable, всего записей - " + peopleGroup.Rows.Count + ", группа - " + namePointedGroup);
        }


        private void infoItem_Click(object sender, EventArgs e)
        {
            ShowDataTableDbQuery(databasePerson, "TechnicalInfo",
                "SELECT PCName AS 'Версия Windows', POName AS 'Путь к ПО', POVersion AS 'Версия ПО', " +
                "LastDateStarted AS 'Дата использования', CurrentUser, FreeRam, GuidAppication ",
                "ORDER BY LastDateStarted DESC");
        }

        private void EnterEditAnualItem_Click(object sender, EventArgs e) //Select - EnterEditAnual() or ExitEditAnual()
        {
            if (EditAnualDaysItem.Text.Contains(@"Выходные(рабочие) дни"))
            {
                AddAnualDateItem.Font = new Font(this.Font, FontStyle.Bold);
                EditAnualDaysItem.Font = new Font(this.Font, FontStyle.Bold);
                EnterEditAnual();
            }
            else if (EditAnualDaysItem.Text.Contains(@"Завершить редактирование"))
            {
                ExitEditAnual();
                EditAnualDaysItem.Font = new Font(this.Font, FontStyle.Regular);
                AddAnualDateItem.Font = new Font(this.Font, FontStyle.Regular);
            }
        }

        private void EnterEditAnual()
        {
            ShowDataTableDbQuery(databasePerson, "BoldedDates",
                "SELECT DayBolded AS 'Праздничный (выходной) день', DayType AS 'Тип выходного дня', " +
                "NAV AS 'Персонально(NAV) или для всех(0)', DayDesciption AS 'Описание', DateCreated AS 'Дата добавления'",
                " ORDER BY DayBolded desc, NAV asc; ");

            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            _MenuItemEnabled(AddAnualDateItem, true);
            CheckBoxesFiltersAll_Visible(false);

            comboBoxFio.Items.Add("");
            comboBoxFio.SelectedIndex = 0;

            toolTip1.SetToolTip(textBoxGroup, "Тип дня - 'Выходной' или 'Рабочий'");
            toolTip1.SetToolTip(textBoxGroupDescription, "Описание дня");
            labelGroup.Text = "Тип";
            textBoxGroup.Text = "";

            StatusLabel2.ForeColor = Color.Crimson;
            EditAnualDaysItem.Text = @"Завершить редактирование";
            EditAnualDaysItem.ToolTipText = @"Выйти из режима редактирования рабочих и выходных дней";
            _toolStripStatusLabelSetText(StatusLabel2, @"Режим редактирования рабочих и выходных дней");
        }

        private void ExitEditAnual()
        {
            comboBoxFio.Items?.RemoveAt(comboBoxFio.FindStringExact(""));
            if (comboBoxFio.Items.Count > 0)
            { comboBoxFio.SelectedIndex = 0; }

            _MenuItemEnabled(FunctionMenuItem, true);
            _MenuItemEnabled(GroupsMenuItem, true);
            _MenuItemEnabled(AddAnualDateItem, false);

            CheckBoxesFiltersAll_Visible(true);

            EditAnualDaysItem.Text = @"Выходные(рабочие) дни";
            EditAnualDaysItem.ToolTipText = @"Войти в режим добавления/удаления праздничных дней";

            toolTip1.SetToolTip(textBoxGroup, "Создать группу");
            toolTip1.SetToolTip(textBoxGroupDescription, "Изменить описание группы");
            labelGroup.Text = "Группа";
            textBoxGroup.Text = "";

            StatusLabel2.ForeColor = Color.Black;
            _toolStripStatusLabelSetText(StatusLabel2, @"Завершен 'Режим работы с праздниками и выходными'");

            nameOfLastTableFromDB = "ListFIO";
            SeekAndShowMembersOfGroup("");
        }

        private void AddAnualDateItem_Click(object sender, EventArgs e) //AddAnualDate()
        {
            AddAnualDate();
            ShowDataTableDbQuery(databasePerson, "BoldedDates", "SELECT DayBolded AS 'Праздничный (выходной) день', DayType AS 'Тип выходного дня', " +
            "NAV AS 'Персонально(NAV) или для всех(0)', DayDesciption AS 'Описание', DateCreated AS 'Дата добавления'",
            " ORDER BY DayBolded desc, NAV asc; ");
        }

        private void AddAnualDate()
        {
            monthCalendar.AddAnnuallyBoldedDate(monthCalendar.SelectionStart);
            monthCalendar.UpdateBoldedDates();

            string dayType = "";
            if (textBoxGroup?.Text?.Trim()?.Length == 0 || textBoxGroup?.Text?.ToLower()?.Trim() == "выходной")
            { dayType = "Выходной"; }
            else { dayType = "Рабочий"; }

            string nav = "";
            if (textBoxNav?.Text?.Trim()?.Length != 6)
            { nav = "0"; }
            else { nav = textBoxNav.Text.Trim(); }

            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();

                    using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'BoldedDates' (DayBolded, NAV, DayType, DayDesciption, DateCreated) " +
                        " VALUES (@BoldedDate, @NAV, @DayType, @DayDesciption, @DateCreated)", sqlConnection))
                    {
                        sqlCommand.Parameters.Add("@BoldedDate", DbType.String).Value = monthCalendar.SelectionStart.ToString("yyyy-MM-dd");
                        sqlCommand.Parameters.Add("@NAV", DbType.String).Value = nav;
                        sqlCommand.Parameters.Add("@DayType", DbType.String).Value = dayType;
                        sqlCommand.Parameters.Add("@DayDesciption", DbType.String).Value = textBoxGroupDescription.Text.Trim();
                        sqlCommand.Parameters.Add("@DateCreated", DbType.String).Value = DateTimeToYYYYMMDD();
                        try { sqlCommand.ExecuteNonQuery(); } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                    }
                }
            }
        }

        private void DeleteAnualDateItem_Click(object sender, EventArgs e) //DeleteAnualDay()
        { DeleteAnualDay(); }

        private void DeleteAnualDay()
        {
            DataGridViewSeekValuesInSelectedRow dgSeek = new DataGridViewSeekValuesInSelectedRow();
            dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] {
                @"Праздничный (выходной) день", @"Тип выходного дня", @"Персонально(NAV) или для всех(0)", @"Дата добавления" });

            DeleteDataTableQueryParameters(databasePerson, "BoldedDates",
                "DayBolded", dgSeek.values[0], "DayType", dgSeek.values[1],
                "NAV", dgSeek.values[2], "DateCreated", dgSeek.values[3]);

            ShowDataTableDbQuery(databasePerson, "BoldedDates", "SELECT DayBolded AS 'Праздничный (выходной) день', DayType AS 'Тип выходного дня', " +
            "NAV AS 'Персонально(NAV) или для всех(0)', DayDesciption AS 'Описание', DateCreated AS 'Дата добавления'",
            " ORDER BY DayBolded desc, NAV asc; ");
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

        public void CheckBoxesFiltersAll_Visible(bool state)
        {
            _controlVisible(checkBoxTimeViolations, state);
            _controlVisible(checkBoxReEnter, state);
            _controlVisible(checkBoxCelebrate, state);
            _controlVisible(checkBoxWeekend, state);
        }

        private async void checkBox_CheckStateChanged(object sender, EventArgs e)
        { await Task.Run(() => checkBoxCheckStateChanged()); }

        private void checkBoxCheckStateChanged()
        {
            DataTable dtEmpty = new DataTable();
            PersonFull emptyPerson = new PersonFull();
            SeekAnualDays(ref dtEmpty, ref emptyPerson, false,
                _dateTimePickerReturnArray(dateTimePickerStart),
                _dateTimePickerReturnArray(dateTimePickerEnd),
                ref myBoldedDates, ref workSelectedDays);

            dtEmpty.Dispose();

            CheckBoxesFiltersAll_Enable(false);
            _controlVisible(dataGridView1, false);

            string nameGroup = _textBoxReturnText(textBoxGroup);

            DataTable dtTempIntermediate = dtPeople.Clone();
            dtPersonTempAllColumns = dtPeople.Clone();
            PersonFull person = new PersonFull()
            {
                FIO = _textBoxReturnText(textBoxFIO),
                NAV = _textBoxReturnText(textBoxNav),
                GroupPerson = nameGroup,
                Department = nameGroup,
                ControlInSeconds = (int)(60 * 60 * numUpHourStart + 60 * numUpMinuteStart),
                ControlOutSeconds = (int)(60 * 60 * numUpHourEnd + 60 * numUpMinuteEnd),
                ControlInHHMM = ConvertDecimalTimeToStringHHMM(numUpHourStart, numUpMinuteStart),
                ControlOutHHMM = ConvertDecimalTimeToStringHHMM(numUpHourEnd, numUpMinuteEnd)
            };

            dtPersonTemp?.Clear();

            if ((nameOfLastTableFromDB == "PeopleGroupDesciption" || nameOfLastTableFromDB == "PeopleGroup") && nameGroup.Length > 0)
            {
                dtPeopleGroup.Clear();
                LoadGroupMembersFromDbToDataTable(nameGroup, ref dtPeopleGroup);

                if (_CheckboxCheckedStateReturn(checkBoxReEnter))
                {
                    foreach (DataRow row in dtPeopleGroup.Rows)
                    {
                        if (row[FIO]?.ToString()?.Length > 0 && (row[GROUP]?.ToString() == nameGroup || (@"@" + row[DEPARTMENT_ID]?.ToString()) == nameGroup))
                        {
                            person = new PersonFull
                            {
                                FIO = row[FIO].ToString(),
                                NAV = row[CODE].ToString(),
                                GroupPerson = row[GROUP].ToString(),
                                Department = row[DEPARTMENT].ToString(),
                                PositionInDepartment = row[EMPLOYEE_POSITION].ToString(),
                                City = row[PLACE_EMPLOYEE]?.ToString(),
                                DepartmentId = row[DEPARTMENT_ID].ToString(),
                                ControlInSeconds = ConvertStringTimeHHMMToSeconds(row[DESIRED_TIME_IN].ToString()),
                                ControlOutSeconds = ConvertStringTimeHHMMToSeconds(row[DESIRED_TIME_OUT].ToString()),
                                ControlInHHMM = row[DESIRED_TIME_IN].ToString(),
                                ControlOutHHMM = row[DESIRED_TIME_OUT].ToString(),
                                Comment = row[EMPLOYEE_SHIFT_COMMENT].ToString(),
                                Shift = row[EMPLOYEE_SHIFT].ToString()
                            };

                            FilterDataByNav(ref person, ref dtPersonRegistrationsFullList, ref dtTempIntermediate);
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
                { FilterDataByNav(ref person, ref dtPersonRegistrationsFullList, ref dtTempIntermediate); }
            }

            //store all columns
            dtPersonTempAllColumns = dtTempIntermediate.Copy();

            //show selected data     
            //distinct Records        
            var namesDistinctColumnsArray = arrayAllColumnsDataTablePeople.Except(nameHidenColumnsArray).ToArray(); //take distinct data
            dtPersonTemp = GetDistinctRecords(dtTempIntermediate, namesDistinctColumnsArray);

            //change enabling of checkboxes
            if (_CheckboxCheckedStateReturn(checkBoxReEnter))// if (checkBoxReEnter.Checked)
            {
                _controlEnable(checkBoxTimeViolations, true);
                _controlEnable(checkBoxWeekend, true);
                _controlEnable(checkBoxCelebrate, true);

                if (_CheckboxCheckedStateReturn(checkBoxTimeViolations))  // if (checkBoxStartWorkInTime.Checked)
                { _MenuItemBackColorChange(LoadDataItem, SystemColors.Control); }
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

            ShowDatatableOnDatagridview(dtPersonTemp, nameHidenColumnsArray, "PeopleGroup");
            panelViewResize(numberPeopleInLoading);
            _controlVisible(dataGridView1, true);
            _controlEnable(checkBoxReEnter, true);
            dtTempIntermediate = null;
        }

        public static DataTable GetDistinctRecords(DataTable dt, string[] Columns)
        {
            DataTable dtUniqRecords = new DataTable();
            dtUniqRecords = dt.DefaultView.ToTable(true, Columns);
            return dtUniqRecords;
        }

        private void FilterDataByNav(ref PersonFull person, ref DataTable dataTableSource, ref DataTable dataTableForStoring, string typeReport = "Полный")
        {
            logger.Trace("FilterDataByNav: " + person.NAV + "| dataTableSource: " + dataTableSource.Rows.Count, "| typeReport: " + typeReport);
            DataRow rowDtStoring;
            DataTable dtTemp = dataTableSource.Clone();

            HashSet<string> hsDays = new HashSet<string>();
            DataTable dtAllRegistrationsInSelectedDay = dataTableSource.Clone(); //All registrations in the selected day

            int firstRegistrationInDay;
            int lastRegistrationInDay;
            int workedSeconds = 0;
            string exceptReason = "";

            try
            {
                var allWorkedDaysPerson = dataTableSource.Select("[NAV-код] = '" + person.NAV + "'");

                if (_CheckboxCheckedStateReturn(checkBoxReEnter) || currentAction == "sendEmail") //checkBoxReEnter.Checked
                {
                    foreach (DataRow dataRowDate in allWorkedDaysPerson) //make the list of worked days
                    { hsDays.Add(dataRowDate[DATE_REGISTRATION]?.ToString()); }

                    foreach (var workedDay in hsDays.ToArray())
                    {
                        //Select only one row with selected NAV for the selected workedDay
                        dtAllRegistrationsInSelectedDay = allWorkedDaysPerson.Distinct().CopyToDataTable().Select("[Дата регистрации] = '" + workedDay + "'").CopyToDataTable();

                        rowDtStoring = dtAllRegistrationsInSelectedDay.Select("[Дата регистрации] = '" + workedDay + "'").First();
                        rowDtStoring[DAY_OF_WEEK] = DayOfWeekRussian(DateTime.Parse(workedDay).DayOfWeek.ToString());

                        //find first registration within the during selected workedDay
                        firstRegistrationInDay = Convert.ToInt32(dtAllRegistrationsInSelectedDay.Compute("MIN([Время регистрации])", string.Empty));

                        //find last registration within the during selected workedDay
                        lastRegistrationInDay = Convert.ToInt32(dtAllRegistrationsInSelectedDay.Compute("MAX([Время регистрации])", string.Empty));

                        //take and convert a real time coming into a string timearray
                        rowDtStoring[TIME_REGISTRATION] = firstRegistrationInDay;              //("Время регистрации", typeof(decimal)), //15
                        rowDtStoring[REAL_TIME_IN] = ConvertSecondsToStringHHMM(firstRegistrationInDay);  //("Фактич. время прихода ЧЧ:ММ", typeof(string)),//24

                        // rowDtStoring[@"Реальное время ухода"] = lastRegistrationInDay;                 //("Реальное время ухода", typeof(decimal)), //18
                        rowDtStoring[REAL_TIME_OUT] = ConvertSecondsToStringHHMM(lastRegistrationInDay);     //("Фактич. время ухода ЧЧ:ММ", typeof(string)), //25

                        //worked out times
                        workedSeconds = lastRegistrationInDay - firstRegistrationInDay;
                        rowDtStoring[EMPLOYEE_TIME_SPENT] = workedSeconds;                                  // ("Реальное отработанное время", typeof(decimal)), //26
                        //  rowDtStoring[@"Отработанное время ЧЧ:ММ"] = ConvertSecondsToStringHHMM(workedSeconds);  //("Отработанное время ЧЧ:ММ", typeof(string)), //27
                        logger.Trace("FilterDataByNav: " + person.NAV + "| " + rowDtStoring[DATE_REGISTRATION]?.ToString() + " " + firstRegistrationInDay + " - " + lastRegistrationInDay);

                        //todo 
                        //will calculate if day of week different
                        if (firstRegistrationInDay > (person.ControlInSeconds + offsetTimeIn) && firstRegistrationInDay != 0) // "Опоздание ЧЧ:ММ", typeof(bool)),           //28
                        {
                            if (typeReport == "Полный")
                            { rowDtStoring[EMPLOYEE_BEING_LATE] = ConvertSecondsToStringHHMM(firstRegistrationInDay - person.ControlInSeconds); }
                            else if (typeReport == "Упрощенный")
                            { rowDtStoring[EMPLOYEE_BEING_LATE] = "1"; }
                        }

                        if (lastRegistrationInDay < (person.ControlOutSeconds - offsetTimeOut) && lastRegistrationInDay != 0)  // "Ранний уход ЧЧ:ММ", typeof(bool)),                 //29
                        {
                            if (typeReport == "Полный")
                            { rowDtStoring[EMPLOYEE_EARLY_DEPARTURE] = ConvertSecondsToStringHHMM(person.ControlOutSeconds - lastRegistrationInDay); }
                            else if (typeReport == "Упрощенный")
                            { rowDtStoring[EMPLOYEE_EARLY_DEPARTURE] = "1"; }
                        }

                        if (rowDtStoring[EMPLOYEE_ABSENCE]?.ToString() == "1" && typeReport == "Полный")  // "Ранний уход ЧЧ:ММ", typeof(bool)),                 //29
                        {
                            rowDtStoring[EMPLOYEE_ABSENCE] = "Да";
                        }

                        exceptReason = rowDtStoring[EMPLOYEE_SHIFT_COMMENT]?.ToString();

                        rowDtStoring[EMPLOYEE_SHIFT_COMMENT] = outResons.Find((x) => x._id == exceptReason)?._visibleName;

                        switch (exceptReason)
                        {
                            case "2":
                            case "10":
                            case "11":
                            case "18":
                                if (typeReport == "Полный")
                                { rowDtStoring[@"Отпуск"] = outResons.Find((x) => x._id == exceptReason)?._visibleName; }
                                else if (typeReport == "Упрощенный")
                                { rowDtStoring[@"Отпуск"] = "1"; }
                                break;
                            case "3":
                            case "21":
                                if (typeReport == "Полный")
                                { rowDtStoring[EMPLOYEE_SICK_LEAVE] = outResons.Find((x) => x._id == exceptReason)?._visibleName; }
                                else if (typeReport == "Упрощенный")
                                { rowDtStoring[EMPLOYEE_SICK_LEAVE] = "1"; }
                                break;
                            case "1":
                            case "9":
                            case "12":
                            case "20":
                                if (typeReport == "Полный")
                                { rowDtStoring[EMPLOYEE_HOOKY] = outResons.Find((x) => x._id == exceptReason)?._visibleName; }
                                else if (typeReport == "Упрощенный")
                                { rowDtStoring[EMPLOYEE_HOOKY] = "1"; }
                                break;
                            case "4":
                            case "5":
                            case "6":
                            case "7":
                                rowDtStoring[EMPLOYEE_SHIFT_COMMENT] = outResons.Find((x) => x._id == exceptReason)?._visibleName;
                                break;
                            default:
                                break;
                        }

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
                    SeekAnualDays(ref dtTemp, ref person, true,
                        ConvertStringDateToIntArray(reportStartDay),
                        ConvertStringDateToIntArray(reportLastDay),
                        ref myBoldedDates, ref workSelectedDays);
                }

                if (_CheckboxCheckedStateReturn(checkBoxTimeViolations)) //checkBoxStartWorkInTime Checking
                { QueryDeleteDataFromDataTable(ref dtTemp, "[Опоздание ЧЧ:ММ]='' AND [Ранний уход ЧЧ:ММ]=''", person.NAV); }

                foreach (DataRow dr in dtTemp.AsEnumerable())
                { dataTableForStoring.ImportRow(dr); }

                allWorkedDaysPerson = null;
            } catch (Exception expt) { MessageBox.Show(expt.ToString()); }

            hsDays = null;
            rowDtStoring = null;
            dtTemp = null;
            exceptReason = null;
            dtAllRegistrationsInSelectedDay = null;
        }



        /// <summary>
        /// ///////////////////////////////////////////////////////////////
        /// </summary>
        //check dubled function!!!!!!!!

        private List<string> ReturnBoldedDaysFromDB(string nav, string dayType)
        {
            List<string> boldedDays = new List<string>();

            using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
            {
                sqlConnection.Open();
                using (var sqlCommand = new SQLiteCommand(
                    "SELECT DayBolded FROM BoldedDates WHERE (NAV LIKE '" + nav + "' OR  NAV LIKE '0') AND DayType LIKE '" + dayType + "';", sqlConnection))
                {
                    using (var sqlReader = sqlCommand.ExecuteReader())
                    {
                        foreach (DbDataRecord record in sqlReader)
                        {
                            if (record["DayBolded"]?.ToString()?.Length > 0)
                            {
                                boldedDays.Add(record["DayBolded"].ToString());
                            }
                        }
                    }
                }
                sqlConnection.Close();
            }
            return boldedDays;
        }

        private void SeekAnualDays(ref DataTable dt, ref PersonFull person, bool delRow, int[] startOfPeriod, int[] endOfPeriod, ref string[] boldedDates, ref string[] workDates)//   //Exclude Anual Days from the table "PersonTemp" DB
        {
            //test
            logger.Trace("SeekAnualDays: " + person.NAV + " " + startOfPeriod[0]);
            logger.Trace("SeekAnualDays,start: " + startOfPeriod[0] + " " + startOfPeriod[1] + " " + startOfPeriod[2]);
            logger.Trace("SeekAnualDays,end: " + endOfPeriod[0] + " " + endOfPeriod[1] + " " + endOfPeriod[2]);

            List<string> daysBolded = new List<string>();

            var oneDay = TimeSpan.FromDays(1);
            var twoDays = TimeSpan.FromDays(2);
            var mySelectedStartDay = new DateTime(startOfPeriod[0], startOfPeriod[1], startOfPeriod[2]);
            var mySelectedEndDay = new DateTime(endOfPeriod[0], endOfPeriod[1], endOfPeriod[2]);
            var myMonthCalendar = new MonthCalendar();

            myMonthCalendar.MaxSelectionCount = 60;
            myMonthCalendar.SelectionRange = new SelectionRange(mySelectedStartDay, mySelectedEndDay);
            myMonthCalendar.FirstDayOfWeek = Day.Monday;

            for (int year = -1; year < 3; year++)
            {
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startOfPeriod[0] + year, 1, 1));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startOfPeriod[0] + year, 1, 2));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startOfPeriod[0] + year, 1, 7));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startOfPeriod[0] + year, 3, 8));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startOfPeriod[0] + year, 5, 1));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startOfPeriod[0] + year, 5, 2));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startOfPeriod[0] + year, 5, 9));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startOfPeriod[0] + year, 6, 28));
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startOfPeriod[0] + year, 8, 24));    // (plavayuschaya data)
                myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startOfPeriod[0] + year, 10, 16));   // (plavayuschaya data)
            }

            // Алгоритм для вычисления католической Пасхи http://snippets.dzone.com/posts/show/765
            int Y = startOfPeriod[0];
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
            myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startOfPeriod[0], monthEaster, dayEaster) + oneDay);

            foreach (string dayAdditional in ReturnBoldedDaysFromDB(person.NAV, @"Выходной")) // or - Рабочий
            { myMonthCalendar.AddAnnuallyBoldedDate(DateTime.Parse(dayAdditional)); }

            //Independence day
            DateTime dayBolded = new DateTime(startOfPeriod[0], 8, 24);
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Sunday:
                    myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startOfPeriod[0], 8, 24) + oneDay);    // (plavayuschaya data)
                    break;
                case (int)Day.Saturday:
                    myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startOfPeriod[0], 8, 24) + twoDays);    // (plavayuschaya data)
                    break;
                default:
                    break;
            }

            //day of Ukraine Force
            dayBolded = new DateTime(startOfPeriod[0], 10, 16);
            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Sunday:
                    myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startOfPeriod[0], 10, 16) + oneDay);    // (plavayuschaya data)
                    break;
                case (int)Day.Saturday:
                    myMonthCalendar.AddAnnuallyBoldedDate(new DateTime(startOfPeriod[0], 10, 16) + twoDays);    // (plavayuschaya data)
                    break;
                default:
                    break;
            }

            string singleDate = null;

            //Remove additional works days from the bolded days 
            foreach (string dayAdditional in ReturnBoldedDaysFromDB(person.NAV, @"Рабочий"))
            { myMonthCalendar.RemoveAnnuallyBoldedDate(DateTime.Parse(dayAdditional)); }

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
                        QueryDeleteDataFromDataTable(ref dt, "[Дата регистрации]='" + singleDate + "'", person.NAV); // ("Дата регистрации",typeof(string)),//12
                    }
                }
                daysSelected.Add(singleDate);
            }
            dt.AcceptChanges();

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
                            QueryDeleteDataFromDataTable(ref dt, "[Дата регистрации]='" + singleDate + "'", person.NAV); // ("Дата регистрации",typeof(string)),//12
                        }
                    }
                }
            }
            dt.AcceptChanges();

            boldedDates = daysBolded.ToArray();
            workDates = daysSelected.Except(daysBolded).ToArray();

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

            foreach (string dayAdditional in ReturnBoldedDaysFromDB("0", @"Выходной")) // or - Рабочий
            { monthCalendar.AddAnnuallyBoldedDate(DateTime.Parse(dayAdditional)); }

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

            foreach (string dayAdditional in ReturnBoldedDaysFromDB("0", @"Рабочий"))
            { monthCalendar.RemoveAnnuallyBoldedDate(DateTime.Parse(dayAdditional)); }

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





        private void QueryDeleteDataFromDataTable(ref DataTable dt, string queryFull, string NAVcode) //Delete data from the Table of the DB by NAV (both parameters are string)
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
                DeleteTable(databasePerson, "ConfigDB");
                DeleteTable(databasePerson, "EquipmentSettings");
                DeleteTable(databasePerson, "LastTakenPeopleComboList");
                GC.Collect();

                ClearFilesInApplicationFolders(@"*.xlsx", "Excel-файлов");
                ClearFilesInApplicationFolders(@"*.log", "логов");

                _textBoxSetText(textBoxFIO, "");
                _textBoxSetText(textBoxGroup, "");
                _textBoxSetText(textBoxGroupDescription, "");
                _textBoxSetText(textBoxNav, "");

                _comboBoxClr(comboBoxFio);

                TryMakeDB();
            }
            else
            { TryMakeDB(); }

            DataTable dt = new DataTable();
            _dataGridViewSource(dt);

            _toolStripStatusLabelSetText(StatusLabel2, @"Все таблицы очищены");
        }


        //gathering a person's features from textboxes and other controls
        private void SelectPersonFromControls(ref PersonFull personSelected)
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
            PersonFull personVisual = new PersonFull();

            decimal[] timeIn = new decimal[5];
            decimal[] timeOut = new decimal[5];

            try
            {
                DataGridViewSeekValuesInSelectedRow dgSeek = new DataGridViewSeekValuesInSelectedRow();
                dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] {
                    FIO, CODE, GROUP,
                    DESIRED_TIME_IN, DESIRED_TIME_OUT,
                    DEPARTMENT, EMPLOYEE_POSITION, EMPLOYEE_SHIFT, DEPARTMENT_ID,
                    EMPLOYEE_SHIFT_COMMENT
                });

                if (_dataGridView1CurrentRowIndex() > -1)
                {
                    personVisual.FIO = dgSeek.values[0];
                    personVisual.NAV = dgSeek.values[1]; //Take the name of selected group
                    personVisual.ControlInHHMM = dgSeek.values[3]; //Take the name of selected group
                    timeIn = ConvertStringTimeHHMMToDecimalArray(personVisual.ControlInHHMM);
                    personVisual.ControlInSeconds = (int)timeIn[4];
                    personVisual.ControlOutHHMM = dgSeek.values[4]; //Take the name of selected group
                    timeOut = ConvertStringTimeHHMMToDecimalArray(personVisual.ControlOutHHMM);
                    personVisual.ControlOutSeconds = (int)timeOut[4];
                    _numUpDownSet(numUpDownHourStart, timeIn[0]);
                    _numUpDownSet(numUpDownMinuteStart, timeIn[1]);

                    _numUpDownSet(numUpDownHourEnd, timeOut[0]);
                    _numUpDownSet(numUpDownMinuteEnd, timeOut[1]);

                    personVisual.Department = dgSeek.values[5];
                    personVisual.PositionInDepartment = dgSeek.values[6];
                    personVisual.Shift = dgSeek.values[7];
                    personVisual.DepartmentId = dgSeek.values[8];
                    personVisual.Comment = dgSeek.values[9];

                    if (nameOfLastTableFromDB == "PeopleGroup" || nameOfLastTableFromDB == "ListFIO")
                    {
                        personVisual.GroupPerson = dgSeek.values[2]; //Take the name of selected group
                        StatusLabel2.Text = @"Выбрана группа: " + personVisual.GroupPerson + @" | Курсор на: " + personVisual.FIO;
                        groupBoxPeriod.BackColor = Color.PaleGreen;
                    }
                    else if (nameOfLastTableFromDB == "PersonRegistrationsList")
                    {
                        personVisual.GroupPerson = _textBoxReturnText(textBoxGroup);
                        StatusLabel2.Text = @"Выбран: " + personVisual.FIO;
                    }
                }
            } catch (Exception expt) { logger.Info("VisualItem_Click: " + expt.ToString()); }

            if (personVisual.FIO.Length == 0)
            {
                SelectPersonFromControls(ref personVisual);
            }

            _controlVisible(dataGridView1, false);

            CheckBoxesFiltersAll_Enable(false);

            if (_CheckboxCheckedStateReturn(checkBoxReEnter))
            {
                logger.Trace("DrawFullWorkedPeriodRegistration: ");
                DrawFullWorkedPeriodRegistration(ref personVisual);
            }
            else
            {
                logger.Trace("DrawRegistration: ");
                DrawRegistration(ref personVisual);
            }

            _MenuItemVisible(TableModeItem, true);
            _MenuItemVisible(VisualModeItem, false);
            _MenuItemVisible(ChangeColorMenuItem, true);
            _MenuItemVisible(TableExportToExcelItem, false);

            _controlVisible(panelView, true);
            _controlVisible(pictureBox1, true);
        }

        private void DrawRegistration(ref PersonFull personDraw)  // Visualisation of registration
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
                string dayRegistration = "";
                string directionPass = ""; //string pointName = "";
                int minutesIn = 0;     // время входа в минутах планируемое
                int minutesInFact = 0;     // время выхода в минутах фактическое
                int minutesOut = 0;    // время входа в минутах планируемое
                int minutesOutFact = 0;    // время выхода в минутах фактическое

                //variable for a person
                string dayPrevious = "";      //дата в предыдущей выборке
                string directionPrevious = "";
                int timePrevious = 0;
                logger.Trace("DrawRegistration: " + group + " " + fio + " " + nav);
                logger.Trace("DrawRegistration,dtPersonTempAllColumns: " + dtPersonTempAllColumns.Rows.Count);
                //select and distinct dataRow
                var rowsPersonRegistrationsForDraw = dtPersonTempAllColumns.Clone();
                if (group?.Length > 0)
                { rowsPersonRegistrationsForDraw = dtPersonTempAllColumns.Select("[Группа] = '" + group + "'").CopyToDataTable(); }
                else if (nav?.Length == 6)
                { rowsPersonRegistrationsForDraw = dtPersonTempAllColumns.Select("[NAV-код] = '" + nav + "'").CopyToDataTable(); }
                else if (nav?.Length != 6 && fio?.Length > 1)
                { rowsPersonRegistrationsForDraw = dtPersonTempAllColumns.Select("[Фамилия Имя Отчество] = '" + fio + "'").CopyToDataTable(); }
                logger.Trace("DrawRegistration,rowsPersonRegistrationsForDraw.Rows.Count: " + rowsPersonRegistrationsForDraw.Rows.Count);

                //select and count unique NAV-codes - the number of selected people
                HashSet<string> hsNAV = new HashSet<string>(); //unique NAV-codes
                foreach (DataRow row in rowsPersonRegistrationsForDraw.Rows)
                { hsNAV.Add(row[CODE].ToString()); }
                string[] arrayNAVs = hsNAV.ToArray(); //unique NAV-codes
                int countNAVs = arrayNAVs.Count(); //the number of selected people
                numberPeopleInLoading = countNAVs;
                logger.Trace("DrawRegistration,countNAVs: " + countNAVs);

                //count max number of events in-out all of selected people (the group or a single person)
                //It needs to prevent the error "index of scope"
                DataTable dtEmpty = new DataTable();
                SeekAnualDays(ref dtEmpty, ref personDraw, false,
                    _dateTimePickerReturnArray(dateTimePickerStart),
                    _dateTimePickerReturnArray(dateTimePickerEnd),
                    ref myBoldedDates, ref workSelectedDays);
                dtEmpty.Dispose();

                foreach (DataRow row in rowsPersonRegistrationsForDraw.Rows)
                {
                    for (int k = 0; k < workSelectedDays.Length; k++)
                    {
                        if (workSelectedDays[k].Length == 10 && row[DATE_REGISTRATION].ToString().Contains(workSelectedDays[k]))  //если приход пришел в рабочий день - считаем
                        { iLenghtRect++; }
                    }
                }
                logger.Trace("DrawRegistration,iLenghtRect: " + iLenghtRect);

                panelView.Height = iShiftHeightAll + iStringHeight * workSelectedDays.Count() * countNAVs;
                // panelView.AutoScroll = false;
                // panelView.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                // panelView.Anchor = AnchorStyles.Bottom;
                // panelView.Anchor = AnchorStyles.Left;
                // panelView.Dock = DockStyle.None;
                _panelResume(panelView);

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
                        logger.Trace("DrawRegistration,draw: " + singleNav);

                        foreach (string workDay in workSelectedDays)
                        {
                            foreach (DataRow row in rowsPersonRegistrationsForDraw.Rows)
                            {
                                nav = row[CODE].ToString();
                                dayRegistration = row[DATE_REGISTRATION].ToString();


                                if (singleNav.Contains(nav) && dayRegistration.Contains(workDay))
                                {
                                    fio = row[FIO].ToString();
                                    minutesIn = (int)ConvertStringTimeHHMMToDecimalArray(row[DESIRED_TIME_IN].ToString())[3];
                                    minutesInFact = (int)ConvertStringTimeHHMMToDecimalArray(row[REAL_TIME_IN].ToString())[3];
                                    minutesOut = (int)ConvertStringTimeHHMMToDecimalArray(row[DESIRED_TIME_OUT].ToString())[3];
                                    minutesOutFact = (int)ConvertStringTimeHHMMToDecimalArray(row[REAL_TIME_OUT].ToString())[3];
                                    directionPass = row[DIRECTION_WAY].ToString().ToLower();

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
                    sLastSelectedElement = "DrawRegistration";
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
            } catch (Exception expt) { logger.Info(expt.ToString()); }
        }

        private void DrawFullWorkedPeriodRegistration(ref PersonFull personDraw)  // Draw the whole period registration
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
            { hsNAV.Add(row[CODE].ToString()); }
            string[] arrayNAVs = hsNAV.ToArray(); //unique NAV-codes
            int countNAVs = arrayNAVs.Count(); //the number of selected people
            numberPeopleInLoading = countNAVs;

            //count max number of events in-out all of selected people (the group or a single person)
            //It needs to prevent the error "index of scope"
            DataTable dtEmpty = new DataTable();
            SeekAnualDays(ref dtEmpty, ref personDraw, false,
                _dateTimePickerReturnArray(dateTimePickerStart),
                _dateTimePickerReturnArray(dateTimePickerEnd),
                ref myBoldedDates, ref workSelectedDays);
            dtEmpty.Dispose();

            foreach (DataRow row in rowsPersonRegistrationsForDraw)
            {
                for (int k = 0; k < workSelectedDays.Length; k++)
                {
                    if (workSelectedDays[k].Length == 10 && row[DATE_REGISTRATION].ToString().Contains(workSelectedDays[k]))  //если приход пришел в рабочий день - считаем
                    { iLenghtRect++; }
                }
            }

            panelView.Height = iShiftHeightAll + iStringHeight * workSelectedDays.Count() * countNAVs;
            // panelView.AutoScroll = false;
            // panelView.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            // panelView.Anchor = AnchorStyles.Bottom;
            // panelView.Anchor = AnchorStyles.Left;
            // panelView.Dock = DockStyle.None;
            _panelResume(panelView);

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
                            nav = row[CODE].ToString();
                            dayRegistration = row[DATE_REGISTRATION].ToString();

                            if (singleNav.Contains(nav) && dayRegistration.Contains(workDay))
                            {
                                fio = row[FIO].ToString();
                                minutesIn = (int)ConvertStringTimeHHMMToDecimalArray(row[DESIRED_TIME_IN].ToString())[3];
                                minutesInFact = (int)ConvertStringTimeHHMMToDecimalArray(row[REAL_TIME_IN].ToString())[3];
                                minutesOut = (int)ConvertStringTimeHHMMToDecimalArray(row[DESIRED_TIME_OUT].ToString())[3];
                                minutesOutFact = (int)ConvertStringTimeHHMMToDecimalArray(row[REAL_TIME_OUT].ToString())[3];
                                directionPass = row[DIRECTION_WAY].ToString().ToLower();

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
                sLastSelectedElement = "DrawFullWorkedPeriodRegistration";
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
            _controlVisible(pictureBox1, false);

            try
            {
                if (panelView?.Controls?.Count > 1) panelView.Controls.RemoveAt(1);
                bmp?.Dispose();
                pictureBox1?.Dispose();
            } catch { }

            _controlVisible(dataGridView1, true);

            sLastSelectedElement = "dataGridView";
            panelViewResize(numberPeopleInLoading);

            CheckBoxesFiltersAll_Enable(true);
            _MenuItemVisible(TableExportToExcelItem, true);
            _MenuItemVisible(TableModeItem, false);
            _MenuItemVisible(VisualModeItem, true);
            _MenuItemVisible(ChangeColorMenuItem, false);
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

        private void ColorRegistrationMenuItem_Click(object sender, EventArgs e) //ChangeColorGraphics(); 
        { ChangeColorGraphics(); }

        private void ChangeColorGraphics()
        {
            switch (clrRealRegistration.Name)
            {
                case "PaleGreen":
                    ColorizeDraw(Color.MediumAquamarine);
                    break;
                case "MediumAquamarine":
                    ColorizeDraw(Color.DarkOrange);
                    break;
                case "DarkOrange":
                    ColorizeDraw(Color.LimeGreen);
                    break;
                case "LimeGreen":
                    ColorizeDraw(Color.Orange);
                    break;
                default:
                    ColorizeDraw(Color.PaleGreen);
                    break;
            }

            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'ConfigDB' (ParameterName, Value, DateCreated) " +
                        " VALUES (@ParameterName, @Value, @DateCreated)", sqlConnection))
                    {
                        sqlCommand.Parameters.Add("@ParameterName", DbType.String).Value = "clrRealRegistration";
                        sqlCommand.Parameters.Add("@Value", DbType.String).Value = clrRealRegistration.Name;
                        sqlCommand.Parameters.Add("@DateCreated", DbType.String).Value = DateTime.Now.ToYYYYMMDDHHMM();
                        try { sqlCommand.ExecuteNonQuery(); } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                    }
                }
            }

            logger.Trace("ColorizeDraw:clrRealRegistration.Name: " + clrRealRegistration.Name);
        }

        public void ColorizeDraw(Color color)
        {
            PersonFull personSelected = new PersonFull();
            SelectPersonFromControls(ref personSelected);

            clrRealRegistration = color;
            switch (sLastSelectedElement)
            {
                case "DrawFullWorkedPeriodRegistration":
                    DrawFullWorkedPeriodRegistration(ref personSelected);
                    break;
                case "DrawRegistration":
                    DrawRegistration(ref personSelected);
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
                 @"1.1. Добавить имя СКД-сервера Интеллект (SERVER.DOMAIN.SUBDOMAIN),\nимя и пароль пользователя для доступа к SQL-серверу СКД\n" +
                 @"1.2. Добавить имя сервера с базой сотрудников для корпоративного сайта (SERVER.DOMAIN.SUBDOMAIN),\nимя и пароль пользователя для доступа к MySQL-серверу\n" +
                 @"1.3. Добавить имя почтового сервера (SERVER.DOMAIN.SUBDOMAIN),\nemail и пароль пользователя для отправки рассылок с отчетами\n" +
                 @"2. Сохранить введенные параметры\n" +
                 @"2.1. В случае ввода некорректных данных получения данных и отправка рассылок будет заблокирована\n" +
                 @"2.2. Если данные введены корректно, необходимо перезапустить программу в обычном режиме\n" +
                 @"3. После этого можно:\n" +
                 @"3.1. Получать списки сотрудников, хранимый на СКД-сервере и корпоративном сайте\n" +
                 @"3.2. Использовать ранее сохраненные группы пользователей локально\n" +
                 @"3.3. Добавлять или корректировать праздничные дни персонального для каждого или всех, отгулы, отпуски\n" +
                 @"3.4. Загружать данные регистраций по группам или отдельным сотрудников, генерировать отчеты, визуализировать, отправлять отчеты по спискам рассылок, сгенерированным автоматически\n" +
                 @"3.5. Создавать собственные группы генерации отчетов, собственные рассылки отчетов\n" +
                 @"3.6. Проводить анализ данных как в табличном виде так и визуально, экспортировать данные в Excel файл.\n\nДата и время локального ПК: " +
                _dateTimePickerReturn(dateTimePickerEnd),
                //dateTimePickerEnd.Value,
                @"Информация о программе",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1);
        }

        private void PrepareForMakingFormMailing(object sender, EventArgs e) //MailingItem()
        {
            nameOfLastTableFromDB = "Mailing";
            _MenuItemEnabled(SettingsMenuItem, false);
            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            CheckBoxesFiltersAll_Enable(false);
            _controlVisible(panelView, false);

            btnPropertiesSave.Text = "Сохранить рассылку";
            RemoveClickEvent(btnPropertiesSave);
            btnPropertiesSave.Click += new EventHandler(ButtonPropertiesSave_MailingSave);

            MakeFormMailing();
        }

        private void MakeFormMailing()
        {
            List<string> listComboParameters = new List<string>();

            List<string> periodComboParameters = new List<string>();
            periodComboParameters.Add("Предыдущий месяц");
            periodComboParameters.Add("Текущий месяц");

            List<string> listComboParameters9 = new List<string>();
            listComboParameters9.Add("Активная");
            listComboParameters9.Add("Неактивная");

            List<string> listComboParameters15 = new List<string>();
            listComboParameters15.Add("Полный");
            listComboParameters15.Add("Упрощенный");

            //get list groups from DB and add to listComboParameters
            using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
            {
                sqlConnection.Open();
                using (var sqlCommand = new SQLiteCommand(
                    "SELECT GroupPerson FROM PeopleGroupDesciption;", sqlConnection))
                {
                    using (var sqlReader = sqlCommand.ExecuteReader())
                    {
                        foreach (DbDataRecord record in sqlReader)
                        {
                            if (record["GroupPerson"]?.ToString()?.Length > 0)
                            {
                                listComboParameters.Add(record["GroupPerson"].ToString());
                            }
                        }
                    }
                }
                sqlConnection.Close();
            }
            listComboParameters.Add("Все");

            ViewFormSettings(
                    "", "", "",
                    "Получатель рассылки", "", @"Получатель рассылки в виде: Name@Domain.Subdomain",
                    "", "", "",
                    "Имя", "", @"Краткое наименование рассылки",
                    "Описание", "", "Понятное описание рассылки",
                    "", "", "",
                    "Группы", listComboParameters, "Выполнять подготовку и отправку отчетов по каждой выбранной группе сотрудников",
                    "Период", periodComboParameters, "Выбрать, за какой период делать отчет",
                    "Статус", listComboParameters9, "Статус рассылки",
                    "", "", "",
                    "", "", "",
                    "", "", "",
                    "Вариант отчета", listComboParameters15, "Вариант отображения данных в отчете",
                    "День подготовки отчета", "", "День, в который выполнять подготовку и отправку данного отчета./nДни месяца с 1 до 28"
                    );
        }

        private void SaveMailing(string recipientEmail, string senderEmail, string groupsReport, string nameReport, string descriptionReport,
            string periodPreparing, string status, string dateCreatingMailing, string SendingDate, string typeReport, string daySendingReport)
        {
            bool recipientValid = false;
            bool senderValid = false;

            if (recipientEmail.Length > 0 && recipientEmail.Contains('.') && recipientEmail.Contains('@') && recipientEmail.Split('.').Count() > 1)
            { recipientValid = true; }

            if (senderEmail.Length > 0 && senderEmail.Contains('.') && senderEmail.Contains('@') && senderEmail.Split('.').Count() > 1)
            { senderValid = true; }

            if (databasePerson.Exists && nameReport.Length > 0 && senderValid && recipientValid)
            {
                using (SQLiteConnection sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();

                    SQLiteCommand sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();

                    using (SQLiteCommand sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'Mailing' (SenderEmail, RecipientEmail, GroupsReport, NameReport, Description, Period, Status, DateCreated, SendingLastDate, TypeReport, DayReport)" +
                               " VALUES (@SenderEmail, @RecipientEmail, @GroupsReport, @NameReport, @Description, @Period, @Status, @DateCreated, @SendingLastDate, @TypeReport, @DayReport)", sqlConnection))
                    {
                        sqlCommand.Parameters.Add("@SenderEmail", DbType.String).Value = senderEmail;
                        sqlCommand.Parameters.Add("@RecipientEmail", DbType.String).Value = recipientEmail;
                        sqlCommand.Parameters.Add("@GroupsReport", DbType.String).Value = groupsReport;
                        sqlCommand.Parameters.Add("@NameReport", DbType.String).Value = nameReport;
                        sqlCommand.Parameters.Add("@Description", DbType.String).Value = descriptionReport;
                        sqlCommand.Parameters.Add("@Period", DbType.String).Value = periodPreparing;
                        sqlCommand.Parameters.Add("@Status", DbType.String).Value = status;
                        sqlCommand.Parameters.Add("@DateCreated", DbType.String).Value = dateCreatingMailing;
                        sqlCommand.Parameters.Add("@SendingLastDate", DbType.String).Value = SendingDate;
                        sqlCommand.Parameters.Add("@TypeReport", DbType.String).Value = typeReport;
                        sqlCommand.Parameters.Add("@DayReport", DbType.String).Value = daySendingReport;

                        try { sqlCommand.ExecuteNonQuery(); } catch { }
                    }

                    sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                }
                _toolStripStatusLabelSetText(StatusLabel2, "Добавлена рассылка: " + nameReport + "| Всего рассылок: " + _dataGridView1RowsCount());
            }
        }

        private void ConfigurationItem_Click(object sender, EventArgs e)
        {
            ShowDataTableDbQuery(databasePerson, "ConfigDB", "SELECT ParameterName AS 'Имя параметра', " +
            "Value AS 'Данные', Description AS 'Описание', DateCreated AS 'Дата создания/модификации'",
            " ORDER BY ParameterName asc, DateCreated desc; ");
        }


        private void SettingsProgrammItem_Click(object sender, EventArgs e)
        { MakeFormSettings(); }

        private void MakeFormSettings()
        {
            EnableMainMenuItems(false);
            _controlVisible(panelView, false);

            btnPropertiesSave.Text = "Сохранить настройки";
            RemoveClickEvent(btnPropertiesSave);
            btnPropertiesSave.Click += new EventHandler(buttonPropertiesSave_Click);
            ViewFormSettings(
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
                "Пароль", mysqlServerUserPassword, "Пароль пользователя MySQL-сервера \"MySQLServer\"",
                "", new List<string>(), "",
                "", "", ""
                );
        }

        private void ViewFormSettings(
            string label1, string txtbox1, string tooltip1, string label2, string txtbox2, string tooltip2, string label3, string txtboxPassword3, string tooltip3,
            string label4, string txtbox4, string tooltip4, string label5, string txtbox5, string tooltip5, string label6, string txtboxPassword6, string tooltip6,
            string nameLabel7, List<string> listStrings7, string tooltip7, string periodLabel8, List<string> periodStrings8, string tooltip8, string label9, List<string> listStrings9, string tooltip9,
            string label10, string txtbox10, string tooltip10, string label11, string txtbox11, string tooltip11, string label12, string txtboxPassword12, string tooltip12,
            string label15, List<string> listStrings15, string tooltip15,
            string label16, string txtbox16, string tooltip16
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

                periodCombo = new ListBox
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

            if (listStrings15.Count > 1 && label15.Length > 0)
            {
                labelSettings15 = new Label
                {
                    Text = label15,
                    BackColor = Color.PaleGreen,
                    Location = new Point(20, 180),
                    Size = new Size(590, 22),
                    BorderStyle = BorderStyle.None,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Parent = groupBoxProperties
                };
                comboSettings15 = new ComboBox
                {
                    Location = new Point(300, 181),
                    Size = new Size(120, 20),
                    Parent = groupBoxProperties
                };
                comboSettings15.DataSource = listStrings15;
                toolTip1.SetToolTip(comboSettings15, tooltip15);
            }

            if (label16.Length > 0)
            {
                labelSettings16 = new Label
                {
                    Text = label16,
                    BackColor = Color.PaleGreen,
                    Location = new Point(20, 210),
                    Size = new Size(590, 22),
                    BorderStyle = BorderStyle.None,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Parent = groupBoxProperties
                };
                textBoxSettings16 = new TextBox
                {
                    Text = txtbox16,
                    //PasswordChar = '*',
                    Location = new Point(300, 211),
                    Size = new Size(90, 20),
                    BorderStyle = BorderStyle.FixedSingle,
                    Parent = groupBoxProperties,
                    MaxLength = 2,

                };
                toolTip1.SetToolTip(textBoxSettings16, tooltip16);
                textBoxSettings16.KeyPress += new KeyPressEventHandler(textBoxSettings16_KeyPress);
                textBoxSettings16.TextChanged += new EventHandler(textboxDate_textChanged);
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

            labelSettings15?.BringToFront();
            comboSettings15?.BringToFront();

            labelSettings16?.BringToFront();
            textBoxSettings16?.BringToFront();

            groupBoxProperties.Visible = true;
        }

        private void buttonPropertiesCancel_Click(object sender, EventArgs e)
        {
            string btnName = btnPropertiesSave.Text;

            DisposeTemporaryControls();

            if (btnName == @"Сохранить рассылку")
            {
                ShowDataTableDbQuery(databasePerson, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
                "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
                "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
                " ORDER BY RecipientEmail asc, DateCreated desc; ");
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

            _controlDispose(labelSettings15);
            _controlDispose(comboSettings15);

            _controlDispose(labelSettings16);
            _controlDispose(textBoxSettings16);
        }

        private void buttonPropertiesSave_Click(object sender, EventArgs e) //SaveProperties()
        {
            SaveProperties(); //btnPropertiesSave 

            DisposeTemporaryControls();
            EnableMainMenuItems(true);
            _controlVisible(panelView, true);
        }



        private void ButtonPropertiesSave_MailingSave(object sender, EventArgs e) //SaveProperties()
        {
            string recipientEmail = _textBoxReturnText(textBoxServer1UserName);
            string senderEmail = mailServerUserName;
            if (mailServerUserName.Length == 0)
            { senderEmail = _textBoxReturnText(textBoxServer1); }
            string nameReport = _textBoxReturnText(textBoxMailServerName);
            string description = _textBoxReturnText(textBoxMailServerUserName);
            string report = _comboBoxReturnSelected(listCombo);
            string period = _listBoxReturnSelected(periodCombo);
            string status = _comboBoxReturnSelected(comboSettings9);
            string typeReport = _comboBoxReturnSelected(comboSettings15);
            string dayReport = _textBoxReturnText(textBoxSettings16);

            if (recipientEmail.Length > 5 && nameReport.Length > 0)
            {
                SaveMailing(recipientEmail, senderEmail,
                    report, nameReport, description, period, status,
                    DateTime.Now.ToYYYYMMDDHHMM(), "", typeReport, dayReport);
            }

            ShowDataTableDbQuery(databasePerson, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
            "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
            "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
            " ORDER BY RecipientEmail asc, DateCreated desc; ");

            DisposeTemporaryControls();
            EnableMainMenuItems(true);
            _controlVisible(panelView, true);
        }

        private void SaveProperties() //Save Parameters into Registry and variables
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

            CheckAliveIntellectServer(server, user, password);

            if (bServer1Exist)
            {
                _controlVisible(groupBoxProperties, false);
                _MenuItemEnabled(GetFioItem, true);

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
                    using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(myRegKey))
                    {
                        try { EvUserKey.SetValue("SKDServer", server, Microsoft.Win32.RegistryValueKind.String); } catch { }
                        try { EvUserKey.SetValue("SKDUser", EncryptionDecryptionCriticalData.EncryptStringToBase64Text(user, btsMess1, btsMess2), Microsoft.Win32.RegistryValueKind.String); } catch { }
                        try { EvUserKey.SetValue("SKDUserPassword", EncryptionDecryptionCriticalData.EncryptStringToBase64Text(password, btsMess1, btsMess2), Microsoft.Win32.RegistryValueKind.String); } catch { }

                        try { EvUserKey.SetValue("MailServer", mailServer, Microsoft.Win32.RegistryValueKind.String); } catch { }
                        try { EvUserKey.SetValue("MailUser", mailServerUserName, Microsoft.Win32.RegistryValueKind.String); } catch { }
                        try { EvUserKey.SetValue("MailUserPassword", EncryptionDecryptionCriticalData.EncryptStringToBase64Text(mailServerUserPassword, btsMess1, btsMess2), Microsoft.Win32.RegistryValueKind.String); } catch { }

                        try { EvUserKey.SetValue("MySQLServer", mysqlServer, Microsoft.Win32.RegistryValueKind.String); } catch { }
                        try { EvUserKey.SetValue("MySQLUser", mysqlServerUserName, Microsoft.Win32.RegistryValueKind.String); } catch { }
                        try { EvUserKey.SetValue("MySQLUserPassword", EncryptionDecryptionCriticalData.EncryptStringToBase64Text(mysqlServerUserPassword, btsMess1, btsMess2), Microsoft.Win32.RegistryValueKind.String); } catch { }

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
                            try { sqlCommand.Parameters.Add("@Reserv1", DbType.String).Value = EncryptionDecryptionCriticalData.EncryptStringToBase64Text(sServer1UserName, btsMess1, btsMess2); } catch { }
                            try { sqlCommand.Parameters.Add("@Reserv2", DbType.String).Value = EncryptionDecryptionCriticalData.EncryptStringToBase64Text(sServer1UserPassword, btsMess1, btsMess2); } catch { }
                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }

                        using (SQLiteCommand sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'EquipmentSettings' (EquipmentParameterName, EquipmentParameterValue, EquipmentParameterServer, Reserv1, Reserv2)" +
                                " VALUES (@EquipmentParameterName, @EquipmentParameterValue, @EquipmentParameterServer, @Reserv1, @Reserv2)", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@EquipmentParameterName", DbType.String).Value = "MailUser";
                            sqlCommand.Parameters.Add("@EquipmentParameterValue", DbType.String).Value = "MailServer";
                            sqlCommand.Parameters.Add("@EquipmentParameterServer", DbType.String).Value = mailServer;
                            sqlCommand.Parameters.Add("@Reserv1", DbType.String).Value = mailServerUserName;
                            try { sqlCommand.Parameters.Add("@Reserv2", DbType.String).Value = EncryptionDecryptionCriticalData.EncryptStringToBase64Text(mailServerUserPassword, btsMess1, btsMess2); } catch { }
                            try { sqlCommand.ExecuteNonQuery(); } catch { }
                        }

                        using (SQLiteCommand sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'EquipmentSettings' (EquipmentParameterName, EquipmentParameterValue, EquipmentParameterServer, Reserv1, Reserv2)" +
                                " VALUES (@EquipmentParameterName, @EquipmentParameterValue, @EquipmentParameterServer, @Reserv1, @Reserv2)", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@EquipmentParameterName", DbType.String).Value = "MySQLUser";
                            sqlCommand.Parameters.Add("@EquipmentParameterValue", DbType.String).Value = "MySQLServer";
                            sqlCommand.Parameters.Add("@EquipmentParameterServer", DbType.String).Value = mysqlServer;
                            sqlCommand.Parameters.Add("@Reserv1", DbType.String).Value = mysqlServerUserName;
                            try { sqlCommand.Parameters.Add("@Reserv2", DbType.String).Value = EncryptionDecryptionCriticalData.EncryptStringToBase64Text(mysqlServerUserPassword, btsMess1, btsMess2); } catch { }
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

        /*  private void buttonPropertiesSave_Click(object sender, EventArgs e) //SaveProperties()
         {
             string btnName = btnPropertiesSave.Text.ToString();

             if (btnName == @"Сохранить настройки")
             {
                 SaveProperties(); //btnPropertiesSave 
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
                 string period = _listBoxReturnSelected(periodCombo);
                 string status = _comboBoxReturnSelected(comboSettings9);
                 string typeReport = _comboBoxReturnSelected(comboSettings15);
                 string dayReport = _textBoxReturnText(textBoxSettings16);

                 if (recipientEmail.Length > 5 && nameReport.Length > 0)
                 {
                     SaveMailing(recipientEmail, senderEmail,
                         report, nameReport, description, period, status,
                         DateTime.Now.ToYYYYMMDDHHMM(), "", typeReport, dayReport);
                 }

                 ShowDataTableDbQuery(databasePerson, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
                 "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
                 "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
                 " ORDER BY RecipientEmail asc, DateCreated desc; ");
             }

             DisposeTemporaryControls();
             EnableMainMenuItems(true);
             _controlVisible(panelView, true);
         }*/

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

        private void textBoxSettings16_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if (char.IsDigit(e.KeyChar) && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void textboxDate_textChanged(object sender, EventArgs e)
        {
            int result;
            bool correct = false;

            //allow numbers from 1 to 28
            if ((sender as TextBox).Text.Length > 0)
            {
                correct = Int32.TryParse((sender as TextBox).Text, out result);
                if (correct)
                {
                    if (result > 28) { (sender as TextBox).Text = "28"; }
                    else if (result < 1) { (sender as TextBox).Text = "1"; }
                    else
                    {
                        (sender as TextBox).Text = result.ToString();
                    }
                }
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
            if (PersonOrGroupItem.Text == "Работать с одной персоной")
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
            if (PersonOrGroupItem.Text == "Работать с одной персоной")
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
                StatusLabel2.Text = "";
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
            LoadDataItem.Enabled = true;
            LoadDataItem.BackColor = Color.PaleGreen;
            dateTimePickerEnd.MinDate = DateTime.Parse(dateTimePickerStart.Value.Year + "-" + dateTimePickerStart.Value.Month + "-" + dateTimePickerStart.Value.Day);
        }

        private void dateTimePickerEnd_CloseUp(object sender, EventArgs e)
        { dateTimePickerStart.MaxDate = DateTime.Parse(dateTimePickerEnd.Value.Year + "-" + dateTimePickerEnd.Value.Month + "-" + dateTimePickerEnd.Value.Day); }

        private void PersonOrGroupItem_Click(object sender, EventArgs e) //PersonOrGroup()
        { PersonOrGroup(); }

        private void PersonOrGroup()
        {
            string menu = _MenuItemReturnText(PersonOrGroupItem);
            switch (menu)
            {
                case ("Работать с группой"):
                    _MenuItemTextSet(PersonOrGroupItem, "Работать с одной персоной");
                    _controlEnable(comboBoxFio, false);
                    nameOfLastTableFromDB = "PersonRegistrationsList";
                    break;
                case ("Работать с одной персоной"):
                    _MenuItemTextSet(PersonOrGroupItem, "Работать с группой");
                    _controlEnable(comboBoxFio, true);
                    nameOfLastTableFromDB = "PeopleGroup";
                    break;
                default:
                    _MenuItemTextSet(PersonOrGroupItem, "Работать с группой");
                    _controlEnable(comboBoxFio, true);
                    nameOfLastTableFromDB = "PeopleGroup";
                    break;
            }
        }
        //--- End. Behaviour Controls ---//




        //---  Start.  DatagridView functions ---//

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e) //dataGridView1CellClick()
        { dataGridView1CellClick(); }

        private void dataGridView1CellClick()
        {
            DataGridViewSeekValuesInSelectedRow dgSeek;

            int IndexCurrentRow = _dataGridView1CurrentRowIndex();

            if (0 < dataGridView1.Rows.Count && IndexCurrentRow < dataGridView1.Rows.Count)
            {
                decimal[] timeIn = new decimal[4];
                decimal[] timeOut = new decimal[4];

                try
                {
                    dgSeek = new DataGridViewSeekValuesInSelectedRow();

                    if (nameOfLastTableFromDB == "BoldedDates")
                    {
                        //    ShowDataTableDbQuery(databasePerson, "BoldedDates", "SELECT DayBolded AS 'Праздничный (выходной) день', DayType AS 'Тип выходного дня', " +
                        //   "NAV AS 'Персонально(NAV) или для всех(0)', DayDesciption AS 'Описание', DateCreated AS 'Дата добавления'",
                        //    " ORDER BY DayBolded desc, NAV asc; ");
                    }
                    else if (nameOfLastTableFromDB == "PeopleGroupDesciption")
                    {
                        dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] {
                             GROUP,   GROUP_DECRIPTION
                            });

                        textBoxGroup.Text = dgSeek.values[0]; //Take the name of selected group
                        textBoxGroupDescription.Text = dgSeek.values[1]; //Take the name of selected group
                        groupBoxPeriod.BackColor = Color.PaleGreen;
                        groupBoxFilterReport.BackColor = SystemColors.Control;
                        StatusLabel2.Text = @"Выбрана группа: " + dgSeek.values[0];
                        if (textBoxFIO.TextLength > 3)
                        {
                            comboBoxFio.SelectedIndex = comboBoxFio.FindString(textBoxFIO.Text);
                        }
                    }
                    else if (nameOfLastTableFromDB == "ListFIO" || nameOfLastTableFromDB == "PeopleGroup" || nameOfLastTableFromDB == "PersonRegistrationsList")
                    {
                        dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] {
                                GROUP, FIO, CODE,
                                DESIRED_TIME_IN, DESIRED_TIME_OUT
                            });

                        textBoxGroup.Text = dgSeek.values[0];
                        textBoxFIO.Text = dgSeek.values[1];
                        textBoxNav.Text = dgSeek.values[2];

                        StatusLabel2.Text = @"Выбрана группа: " + dgSeek.values[0] +
                            @" |Курсор на: " + ShortFIO(dgSeek.values[1]);

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
                } catch (Exception expt)
                {
                    MessageBox.Show(expt.ToString());
                }
            }
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e) //SearchMembersSelectedGroup()
        { SearchMembersSelectedGroup(); }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e) //DataGridView1CellEndEdit()
        { DataGridView1CellEndEdit(); }

        private void DataGridView1CellEndEdit()
        {
            string fio = "";
            string nav = "";
            string group = "";
            int currRow = _dataGridView1CurrentRowIndex();

            if (currRow > -1)
            {
                try
                {
                    DataGridViewSeekValuesInSelectedRow dgSeek = new DataGridViewSeekValuesInSelectedRow();

                    string currColumn = _dataGridView1ColumnHeaderText(_dataGridView1CurrentColumnIndex());
                    string currCellValue = _dataGridView1CellValue();
                    string editedCell = "";

                    if (nameOfLastTableFromDB == @"BoldedDates")
                    {
                        dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] {
                        @"Праздничный (выходной) день", @"Персонально(NAV) или для всех(0)", @"Тип выходного дня" });

                        string dayType = "";
                        if (textBoxGroup?.Text?.Trim()?.Length == 0 || textBoxGroup?.Text?.ToLower()?.Trim() == "выходной")
                        { dayType = "Выходной"; }
                        else { dayType = "Рабочий"; }

                        if (textBoxNav?.Text?.Trim()?.Length != 6)
                        { nav = "для всех"; }
                        else { nav = textBoxNav.Text.Trim(); }

                        string navD = "";
                        if (dgSeek.values[1]?.Length != 6)
                        { navD = "всех"; }
                        else { navD = dgSeek.values[1]; }

                        using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                        {
                            sqlConnection.Open();

                            using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'BoldedDates' (DayBolded, NAV, DayType, DayDesciption, DateCreated) " +
                                " VALUES (@BoldedDate, @NAV, @DayType, @DayDesciption, @DateCreated)", sqlConnection))
                            {
                                sqlCommand.Parameters.Add("@BoldedDate", DbType.String).Value = monthCalendar.SelectionStart.ToString("yyyy-MM-dd");
                                sqlCommand.Parameters.Add("@NAV", DbType.String).Value = nav;
                                sqlCommand.Parameters.Add("@DayType", DbType.String).Value = dayType;
                                sqlCommand.Parameters.Add("@DayDesciption", DbType.String).Value = textBoxGroupDescription.Text.Trim();
                                sqlCommand.Parameters.Add("@DateCreated", DbType.String).Value = DateTimeToYYYYMMDD();
                                try { sqlCommand.ExecuteNonQuery(); } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                            }
                        }
                    }
                    else if (nameOfLastTableFromDB == @"PeopleGroup" || nameOfLastTableFromDB == "ListFIO")
                    {
                        dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] {
                            FIO, CODE, GROUP,
                            DESIRED_TIME_IN, DESIRED_TIME_OUT,
                            DEPARTMENT, EMPLOYEE_POSITION, EMPLOYEE_SHIFT, EMPLOYEE_SHIFT_COMMENT,
                            DEPARTMENT_ID
                            });

                        fio = dgSeek.values[0];
                        textBoxFIO.Text = fio;

                        nav = dgSeek.values[1];
                        textBoxNav.Text = nav;

                        group = dgSeek.values[2];
                        textBoxGroup.Text = group;

                        //todo
                        //Change to UPDATE, but not Replace!!!!
                        /*
                         myCommand1.CommandText = "UPDATE dept SET dname = :dname, loc = :loc WHERE deptno = @deptno";
                        myCommand1.Parameters.Add("@deptno", 20);
                        myCommand1.Parameters.Add("dname", "SALES");
                        myCommand1.Parameters.Add("loc", "NEW YORK");
                        */
                        using (var connection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                        {
                            connection.Open();
                            using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'PeopleGroup' (FIO, NAV, GroupPerson, ControllingHHMM, ControllingOUTHHMM, Shift, Comment, Department, PositionInDepartment, DepartmentId, City) " +
                                                                                        " VALUES (@FIO, @NAV, @GroupPerson, @ControllingHHMM, @ControllingOUTHHMM, @Shift, @Comment, @Department, @PositionInDepartment, @DepartmentId, @City)", connection))
                            {
                                sqlCommand.Parameters.Add("@FIO", DbType.String).Value = fio;
                                sqlCommand.Parameters.Add("@NAV", DbType.String).Value = nav;

                                sqlCommand.Parameters.Add("@GroupPerson", DbType.String).Value = group;
                                sqlCommand.Parameters.Add("@Department", DbType.String).Value = dgSeek.values[5];
                                sqlCommand.Parameters.Add("@PositionInDepartment", DbType.String).Value = dgSeek.values[6];
                                sqlCommand.Parameters.Add("@DepartmentId", DbType.String).Value = dgSeek.values[9];

                                sqlCommand.Parameters.Add("@ControllingHHMM", DbType.String).Value = dgSeek.values[3];
                                sqlCommand.Parameters.Add("@ControllingOUTHHMM", DbType.String).Value = dgSeek.values[4];

                                sqlCommand.Parameters.Add("@Shift", DbType.String).Value = dgSeek.values[7];
                                sqlCommand.Parameters.Add("@Comment", DbType.String).Value = dgSeek.values[8];

                                try { sqlCommand.ExecuteNonQuery(); } catch { }
                            }
                        }
                        SeekAndShowMembersOfGroup(group);
                        nameOfLastTableFromDB = "PeopleGroup";
                        StatusLabel2.Text = @"Обновлено время прихода " + ShortFIO(fio) + " в группе: " + group;
                    }
                    else if (nameOfLastTableFromDB == @"PeopleGroupDesciption")
                    {
                        dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] {
                             GROUP, GROUP_DECRIPTION });

                        textBoxGroup.Text = dgSeek.values[0]; //Take the name of selected group
                        textBoxGroupDescription.Text = dgSeek.values[1]; //Take the name of selected group
                        groupBoxPeriod.BackColor = Color.PaleGreen;
                        groupBoxFilterReport.BackColor = SystemColors.Control;
                        StatusLabel2.Text = @"Выбрана группа: " + dgSeek.values[0];
                    }
                    else if (nameOfLastTableFromDB == @"Mailing")
                    {
                        dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] {
                            @"Получатель", @"Отчет по группам", @"Наименование", @"Описание",
                            @"Период", @"Статус", @"Тип отчета", @"День отправки отчета" });

                        switch (currColumn)
                        {
                            case "День отправки отчета":
                                int day = 1;
                                bool correct = Int32.TryParse(dgSeek.values[7], out day);
                                if (!correct || day < 1 || day > 28)
                                { editedCell = "1"; }
                                else { editedCell = day.ToString(); }
                                ExecuteSql("UPDATE 'Mailing' SET DayReport='" + editedCell + "' WHERE RecipientEmail='" + dgSeek.values[0]
                                  + "' AND NameReport='" + dgSeek.values[2] + "' AND GroupsReport ='" + dgSeek.values[1]
                                  + "' AND Status ='" + dgSeek.values[5] + "' AND Description ='" + dgSeek.values[3] + "';", databasePerson);
                                break;

                            case "Тип отчета":
                                if (currCellValue == "Полный") { editedCell = "Полный"; }
                                else { editedCell = "Упрощенный"; }

                                ExecuteSql("UPDATE 'Mailing' SET TypeReport='" + editedCell + "' WHERE RecipientEmail='" + dgSeek.values[0]
                                  + "' AND NameReport='" + dgSeek.values[2] + "' AND GroupsReport ='" + dgSeek.values[1]
                                  + "' AND Status ='" + dgSeek.values[5] + "' AND Description ='" + dgSeek.values[3] + "';", databasePerson);
                                break;

                            case "Статус":
                                if (currCellValue == "Активная") { editedCell = "Активная"; }
                                else { editedCell = "Неактивная"; }

                                ExecuteSql("UPDATE 'Mailing' SET Status='" + editedCell + "' WHERE RecipientEmail='" + dgSeek.values[0]
                                  + "' AND NameReport='" + dgSeek.values[2] + "' AND GroupsReport ='" + dgSeek.values[1] + "' AND Description ='" + dgSeek.values[3] + "';", databasePerson);
                                break;

                            case "Период":
                                if (currCellValue == "Текущий месяц") { editedCell = "Текущий месяц"; }
                                else { editedCell = "Предыдущий месяц"; }

                                ExecuteSql("UPDATE 'Mailing' SET Period='" + editedCell + "' WHERE RecipientEmail='" + dgSeek.values[0]
                                  + "' AND NameReport='" + dgSeek.values[2] + "' AND GroupsReport ='" + dgSeek.values[1]
                                  + "' AND Status ='" + dgSeek.values[5] + "' AND Description ='" + dgSeek.values[3] + "';", databasePerson);
                                break;

                            case "Описание":
                                editedCell = currCellValue;

                                ExecuteSql("UPDATE 'Mailing' SET Description='" + editedCell + "' WHERE RecipientEmail='" + dgSeek.values[0]
                                  + "' AND NameReport='" + dgSeek.values[2] + "' AND GroupsReport ='" + dgSeek.values[1] + "' AND Status ='" + dgSeek.values[5] + "';", databasePerson);
                                break;

                            case "Отчет по группам":
                                editedCell = currCellValue;

                                ExecuteSql("UPDATE 'Mailing' SET GroupsReport ='" + editedCell + "' WHERE RecipientEmail='" + dgSeek.values[0]
                                  + "' AND NameReport='" + dgSeek.values[2] + "' AND Description ='" + dgSeek.values[3] + "' AND Status ='" + dgSeek.values[5] + "';", databasePerson);
                                break;

                            case "Получатель":
                                if (currCellValue.Contains('@') && currCellValue.Contains('.'))
                                {
                                    editedCell = currCellValue;

                                    ExecuteSql("UPDATE 'Mailing' SET RecipientEmail ='" + editedCell + "' WHERE Description='" + dgSeek.values[3]
                                      + "' AND NameReport='" + dgSeek.values[2] + "' AND GroupsReport ='" + dgSeek.values[1]
                                      + "' AND Status ='" + dgSeek.values[5] + "' AND Description ='" + dgSeek.values[3] + "';", databasePerson);
                                }
                                break;

                            case "Наименование":
                                editedCell = currCellValue;

                                ExecuteSql("UPDATE 'Mailing' SET NameReport ='" + editedCell + "' WHERE RecipientEmail='" + dgSeek.values[0]
                                  + "' AND  Description='" + dgSeek.values[3] + "' AND GroupsReport ='" + dgSeek.values[1]
                                  + "' AND Period ='" + dgSeek.values[4] + "' AND Status ='" + dgSeek.values[5] + "';", databasePerson);
                                break;

                            default:
                                break;
                        }

                        ShowDataTableDbQuery(databasePerson, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
                        "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
                        "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
                        " ORDER BY RecipientEmail asc, DateCreated desc; ");
                    }
                    else if (nameOfLastTableFromDB == @"MailingException")
                    {
                        dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] { @"Получатель", @"Описание" });

                        switch (currColumn)
                        {
                            case "Получатель":
                                ExecuteSql("UPDATE 'MailingException' SET RecipientEmail='" + currCellValue +
                                    "' WHERE Description='" + dgSeek.values[1] + "';", databasePerson);
                                break;

                            case "Описание":
                                ExecuteSql("UPDATE 'MailingException' SET Description='" + currCellValue +
                                    "' WHERE RecipientEmail='" + dgSeek.values[0] + "';", databasePerson);
                                break;
                            default:
                                break;
                        }

                        ShowDataTableDbQuery(databasePerson, "MailingException", "SELECT RecipientEmail AS 'Получатель', " +
                        "NameReport AS 'Наименование', Description AS 'Описание', DateCreated AS 'Дата создания/модификации', " +
                        " DayReport AS 'День отправки отчета'", " ORDER BY RecipientEmail asc, DateCreated desc; ");
                    }
                    else if (nameOfLastTableFromDB == @"SelectedCityToLoadFromWeb")
                    {
                        dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] { PLACE_EMPLOYEE, @"Дата создания" });

                        switch (currColumn)
                        {
                            case "Местонахождение сотрудника":
                                ExecuteSql("UPDATE 'SelectedCityToLoadFromWeb' SET City='" + dgSeek.values[0] +
                                                    "' WHERE DateCreated='" + dgSeek.values[1] + "';", databasePerson);
                                break;
                            default:
                                break;
                        }

                        ShowDataTableDbQuery(databasePerson, "SelectedCityToLoadFromWeb", "SELECT City AS 'Местонахождение сотрудника', DateCreated AS 'Дата создания'",
                        " ORDER BY City asc, DateCreated desc; ");
                    }
                } catch { }
            }
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            DataGridViewCell cell;
            if (_dataGridView1ColumnCount() > 0 && _dataGridView1RowsCount() > 0)
            {
                if (
                    nameOfLastTableFromDB == @"PeopleGroup" ||
                    nameOfLastTableFromDB == @"Mailing" ||
                    nameOfLastTableFromDB == @"MailingException"
                    )
                {
                    // e.ColumnIndex == this.dataGridView1.Columns["NAV-код"].Index
                    cell = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    cell.ToolTipText = "Для установки нового значения нажмите F2,\nвнесите новое значение,\nа затем нажмите Enter";
                }
            }
            cell = null;
        }

        //right click of mouse on the datagridview
        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {
            int currentMouseOverRow = dataGridView1.HitTest(e.X, e.Y).RowIndex;
            if (e.Button == MouseButtons.Right && currentMouseOverRow > -1)
            {
                ContextMenu mRightClick = new ContextMenu();
                DataGridViewSeekValuesInSelectedRow dgSeek = new DataGridViewSeekValuesInSelectedRow();

                if (nameOfLastTableFromDB == @"PeopleGroupDesciption")
                {
                    dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] { GROUP, GROUP_DECRIPTION });

                    mRightClick.MenuItems.Add(new MenuItem(text: "&Загрузить данные регистраций группы: '" + dgSeek.values[1] +
                        "' за " + _dateTimePickerStartReturnMonth(), onClick: GetDataItem_Click));
                    mRightClick.MenuItems.Add(new MenuItem(text: "&Загрузить данные регистраций группы: '" + dgSeek.values[1] +
                        "' за " + _dateTimePickerStartReturnMonth() + " и &подготовить отчет", onClick: DoReportByRightClick));
                    mRightClick.MenuItems.Add(new MenuItem(text: "Загрузить данные регистраций группы: '" + dgSeek.values[1] +
                        "' за " + _dateTimePickerStartReturnMonth() + " и &отправить отчет адресату: " + mailServerUserName, onClick: DoReportAndEmailByRightClick));
                    mRightClick.MenuItems.Add("-");
                    mRightClick.MenuItems.Add(new MenuItem(text: "Удалить группу: '" + dgSeek.values[0] + "'(" + dgSeek.values[1] + ")", onClick: DeleteCurrentRow));
                    mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                }
                else if (nameOfLastTableFromDB == @"Mailing")
                {
                    dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] {
                        @"Наименование", @"Описание" });

                    mRightClick.MenuItems.Add(new MenuItem(@"Выполнить рассылку:   " + dgSeek.values[0] + "(" + dgSeek.values[1] + ")", DoMainAction));
                    mRightClick.MenuItems.Add(new MenuItem(@"Создать новую рассылку", PrepareForMakingFormMailing));
                    mRightClick.MenuItems.Add(new MenuItem(@"Клонировать рассылку:   " + dgSeek.values[0] + "(" + dgSeek.values[1] + ")", MakeCloneMailing));
                    mRightClick.MenuItems.Add(new MenuItem(@"Состав рассылки:   " + dgSeek.values[0] + "(" + dgSeek.values[1] + ")", MembersGroupItem_Click));
                    mRightClick.MenuItems.Add("-");
                    mRightClick.MenuItems.Add(new MenuItem(@"Удалить рассылку:   " + dgSeek.values[0] + "(" + dgSeek.values[1] + ")", DeleteCurrentRow));
                    mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                }
                else if (nameOfLastTableFromDB == @"MailingException")
                {
                    dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] { @"Получатель" });

                    mRightClick.MenuItems.Add(new MenuItem(@"Добавить новый адрес 'для исключения из рассылок'", MakeNewRecepientExcept));
                    mRightClick.MenuItems.Add("-");
                    mRightClick.MenuItems.Add(new MenuItem(@"Удалить адрес, ранее внесенный как 'исключеный из рассылок':   " + dgSeek.values[0], DeleteCurrentRow));
                    mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                }
                else if (nameOfLastTableFromDB == @"PeopleGroup" || nameOfLastTableFromDB == @"ListFIO")
                {
                    dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] {
                 FIO, CODE, DEPARTMENT, EMPLOYEE_POSITION,
                 CHIEF_ID, EMPLOYEE_SHIFT, DEPARTMENT_ID, PLACE_EMPLOYEE
                            });
                    //todo 
                    //move to the middleware method
                    _textBoxSetText(textBoxGroup, "");

                    mRightClick.MenuItems.Add(new MenuItem(text: "&Загрузить данные регистраций сотрудника: '" + dgSeek.values[0] +
                        "' за " + _dateTimePickerStartReturnMonth(), onClick: GetDataItem_Click));
                    mRightClick.MenuItems.Add("-");
                    mRightClick.MenuItems.Add(new MenuItem(@"Удалить сотрудника из данной группы", DeleteCurrentRow));
                    mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                }
                else if (nameOfLastTableFromDB == @"BoldedDates")
                {
                    dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] {
                        @"Праздничный (выходной) день", @"Персонально(NAV) или для всех(0)", @"Тип выходного дня" });

                    string dayType = "";
                    if (textBoxGroup?.Text?.Trim()?.Length == 0 || textBoxGroup?.Text?.ToLower()?.Trim() == "выходной")
                    { dayType = "Выходной"; }
                    else { dayType = "Рабочий"; }

                    string nav = "";
                    if (textBoxNav?.Text?.Trim()?.Length != 6)
                    { nav = "для всех"; }
                    else { nav = textBoxNav.Text.Trim(); }

                    string navD = "";
                    if (dgSeek.values[1]?.Length != 6)
                    { navD = "всех"; }
                    else { navD = dgSeek.values[1]; }

                    mRightClick.MenuItems.Add(new MenuItem(@"Сохранить для " + nav + @" как '" + dayType + @"' " + monthCalendar.SelectionStart.ToString("yyyy-MM-dd"), AddAnualDateItem_Click));
                    mRightClick.MenuItems.Add("-");
                    mRightClick.MenuItems.Add(new MenuItem(@"Удалить из сохранненых '" + dgSeek.values[2] + @"'  '" + dgSeek.values[0] + @"' для " + navD, DeleteAnualDateItem_Click));
                    mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                }
                else if (nameOfLastTableFromDB == @"SelectedCityToLoadFromWeb")
                {
                    mRightClick.MenuItems.Add(new MenuItem(@"Добавить новый город", AddNewCityToLoadByRightClick));
                    mRightClick.MenuItems.Add("-");
                    mRightClick.MenuItems.Add(new MenuItem(@"Удалить выбранный город", DeleteCityToLoadByRightClick));
                    mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                }
                dgSeek = null;
            }
        }

        private void SelectedToLoadCityItem_Click(object sender, EventArgs e) //SelectedToLoadCity()
        { SelectedToLoadCity(); }

        private void SelectedToLoadCity()
        {
            ShowDataTableDbQuery(databasePerson, "SelectedCityToLoadFromWeb", "SELECT City AS 'Местонахождение сотрудника', DateCreated AS 'Дата создания'",
            " ORDER BY DateCreated desc; ");

            if (dataGridView1?.RowCount < 2)
            {
                AddNewCityToLoad();
                ShowDataTableDbQuery(databasePerson, "SelectedCityToLoadFromWeb", "SELECT City AS 'Местонахождение сотрудника', DateCreated AS 'Дата создания'",
                " ORDER BY DateCreated desc; ");
            }
        }

        private void AddNewCityToLoadByRightClick(object sender, EventArgs e)
        {
            AddNewCityToLoad();
            SelectedToLoadCity();
        }

        private void AddNewCityToLoad()
        {
            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'SelectedCityToLoadFromWeb' (City, DateCreated) " +
                        " VALUES (@City, @DateCreated)", sqlConnection))
                    {
                        sqlCommand.Parameters.Add("@City", DbType.String).Value = "City";
                        sqlCommand.Parameters.Add("@DateCreated", DbType.String).Value = DateTime.Now.ToYYYYMMDDHHMM();
                        try { sqlCommand.ExecuteNonQuery(); } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                    }
                }
            }
        }

        private void DeleteCityToLoadByRightClick(object sender, EventArgs e)
        {
            DeleteCityToLoad();
            SelectedToLoadCity();
        }

        private void DeleteCityToLoad()
        {
            DataGridViewSeekValuesInSelectedRow dgSeek = new DataGridViewSeekValuesInSelectedRow();
            dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] { PLACE_EMPLOYEE, @"Дата создания" });

            DeleteDataTableQueryParameters(databasePerson, "SelectedCityToLoadFromWeb",
                            "City", dgSeek.values[0]);
        }





        private void DoReportAndEmailByRightClick(object sender, EventArgs e)
        { DoReportAndEmailByRightClick(); }

        private void DoReportAndEmailByRightClick()
        {
            DataGridViewSeekValuesInSelectedRow dgSeek = new DataGridViewSeekValuesInSelectedRow();
            dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] { GROUP, GROUP_DECRIPTION });

            _toolStripStatusLabelSetText(StatusLabel2, "Готовлю отчет по группе" + dgSeek.values[0]);
            logger.Trace("DoReportAndEmailByRightClick: " + mailServerUserName + "|" + dgSeek.values[0]);

            MailingAction("sendEmail", mailServerUserName, mailServerUserName,
            dgSeek.values[0], dgSeek.values[0], "Test", SelectedDatetimePickersPeriodMonth(), "Активная", "Полный", DateTime.Now.ToYYYYMMDDHHMM());
            _ProgressBar1Stop();
        }

        private void DoReportByRightClick(object sender, EventArgs e)
        { DoReportByRightClick(); }

        private void DoReportByRightClick()
        {
            DataGridViewSeekValuesInSelectedRow dgSeek = new DataGridViewSeekValuesInSelectedRow();
            dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] { GROUP, GROUP_DECRIPTION });

            _toolStripStatusLabelSetText(StatusLabel2, "Готовлю отчет по группе" + dgSeek.values[0]);
            logger.Trace("DoReportByRightClick: " + dgSeek.values[0]);

            GetRegistrationAndSendReport(dgSeek.values[0], dgSeek.values[0], dgSeek.values[1], SelectedDatetimePickersPeriodMonth(), "Активная", "Полный", DateTime.Now.ToYYYYMMDDHHMM(), false, "", "");

            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("explorer.exe", " /select, " + filePathExcelReport)); // //System.Reflection.Assembly.GetExecutingAssembly().Location)

            _ProgressBar1Stop();
            nameOfLastTableFromDB = "PeopleGroupDesciption";
        }

        private void MakeCloneMailing(object sender, EventArgs e) //MakeCloneMailing()
        { MakeCloneMailing(); }

        private void MakeCloneMailing()
        {
            DataGridViewSeekValuesInSelectedRow dgSeek = new DataGridViewSeekValuesInSelectedRow();
            dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] {
                @"Получатель", @"Отчет по группам", @"Наименование", @"Описание", @"Период"});

            SaveMailing(
               dgSeek.values[0], mailServerUserName, dgSeek.values[1], dgSeek.values[2] + "_1",
               dgSeek.values[3] + "_1", dgSeek.values[4], "Неактивная", DateTime.Now.ToYYYYMMDDHHMM(), "", "Копия", "1");

            ShowDataTableDbQuery(databasePerson, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
            "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
            "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
            " ORDER BY RecipientEmail asc, DateCreated desc; ");
        }

        private void MakeNewRecepientExcept(object sender, EventArgs e) //MakeNewRecepientExcept(), ShowDataTableDbQuery()
        {
            MakeNewRecepientExcept();
            ShowDataTableDbQuery(databasePerson, "MailingException", "SELECT RecipientEmail AS 'Получатель', " +
            "NameReport AS 'Наименование', Description AS 'Описание', DateCreated AS 'Дата создания/модификации', " +
            " DayReport AS 'День отправки отчета'", " ORDER BY RecipientEmail asc, DateCreated desc; ");
        }

        private void MakeNewRecepientExcept()
        {
            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'MailingException' (RecipientEmail, GroupsReport, Description, DateCreated) " +
                        " VALUES (@RecipientEmail, @GroupsReport, @Description, @DateCreated)", sqlConnection))
                    {
                        sqlCommand.Parameters.Add("@RecipientEmail", DbType.String).Value = "example@mail.com";
                        sqlCommand.Parameters.Add("@GroupsReport", DbType.String).Value = "";
                        sqlCommand.Parameters.Add("@Description", DbType.String).Value = "example";
                        sqlCommand.Parameters.Add("@DateCreated", DbType.String).Value = DateTime.Now.ToYYYYMMDDHHMM();
                        try { sqlCommand.ExecuteNonQuery(); } catch (Exception expt) { MessageBox.Show(expt.ToString()); }
                    }
                }
            }
        }

        private void MailingsExceptItem_Click(object sender, EventArgs e)
        {
            _controlEnable(comboBoxFio, false);
            dataGridView1.Select();

            ShowDataTableDbQuery(databasePerson, "MailingException", "SELECT RecipientEmail AS 'Получатель', " +
            "NameReport AS 'Наименование', Description AS 'Описание', DateCreated AS 'Дата создания/модификации', " +
            " DayReport AS 'День отправки отчета'", " ORDER BY RecipientEmail asc, DateCreated desc; ");
        }

        private void MailingsShowItem_Click(object sender, EventArgs e)
        {
            _controlEnable(comboBoxFio, false);
            dataGridView1.Select();

            ShowDataTableDbQuery(databasePerson, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
            "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
            "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
            " ORDER BY RecipientEmail asc, DateCreated desc; ");
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
                        dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] {
                            @"Получатель", @"Отчет по группам", @"Наименование", @"Описание",
                            @"Период", @"Статус", @"Тип отчета", @"День отправки отчета" });
                        _toolStripStatusLabelSetText(StatusLabel2, "Готовлю отчет " + dgSeek.values[2]);

                        ExecuteSql("UPDATE 'Mailing' SET SendingLastDate='" + DateTime.Now.ToYYYYMMDDHHMM() +
                            "' WHERE RecipientEmail='" + dgSeek.values[0] + "' AND NameReport='" + dgSeek.values[2] +
                            "' AND GroupsReport ='" + dgSeek.values[1] + "';", databasePerson);

                        MailingAction("sendEmail", dgSeek.values[0], mailServerUserName,
                            dgSeek.values[1], dgSeek.values[2], dgSeek.values[3], dgSeek.values[4],
                            dgSeek.values[5], dgSeek.values[6], dgSeek.values[7]);

                        logger.Info("DoMainAction, sendEmail: " +
                            dgSeek.values[0] + "|" + dgSeek.values[1] + "|" + dgSeek.values[2] + "|" +
                            dgSeek.values[3] + "|" + dgSeek.values[4] + "|" + dgSeek.values[5] + "|" +
                            dgSeek.values[6] + "|" + dgSeek.values[7]);

                        ShowDataTableDbQuery(databasePerson, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
                        "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
                        "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
                        " ORDER BY RecipientEmail asc, DateCreated desc; ");
                        break;
                    }
                default:
                    break;
            }

            _ProgressBar1Stop();
        }

        private void DeleteCurrentRow(object sender, EventArgs e) //DeleteCurrentRow()
        {
            if (_dataGridView1CurrentRowIndex() > -1)
            { DeleteCurrentRow(); }
        }

        private void DeleteCurrentRow()
        {
            string group = _textBoxReturnText(textBoxGroup);
            DataGridViewSeekValuesInSelectedRow dgSeek = new DataGridViewSeekValuesInSelectedRow();

            switch (nameOfLastTableFromDB)
            {
                case "PeopleGroupDesciption":
                    {
                        dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] { GROUP });

                        DeleteDataTableQueryParameters(databasePerson, "PeopleGroup", "GroupPerson", dgSeek.values[0], "", "", "", "");
                        DeleteDataTableQueryParameters(databasePerson, "PeopleGroupDesciption", "GroupPerson", dgSeek.values[0], "", "", "", "");

                        UpdateAmountAndRecepientOfPeopleGroupDesciption();
                        ShowDataTableDbQuery(databasePerson, "PeopleGroupDesciption", "SELECT GroupPerson AS 'Группа', GroupPersonDescription AS 'Описание группы', AmountStaffInDepartment AS 'Колличество сотрудников в группе' ", " group by GroupPerson ORDER BY GroupPerson asc; ");
                        _toolStripStatusLabelSetText(StatusLabel2, "Удалена группа: " + dgSeek.values[0] + "| Всего групп: " + _dataGridView1RowsCount());
                        MembersGroupItem.Enabled = true;
                        break;
                    }
                case "PeopleGroup" when group.Length > 0:
                    {
                        int indexCurrentRow = _dataGridView1CurrentRowIndex();

                        dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] { CODE, GROUP });
                        DeleteDataTableQueryParameters(databasePerson, "PeopleGroup", "GroupPerson", dgSeek.values[1], "NAV", dgSeek.values[0], "", "");

                        if (indexCurrentRow > 2)
                        { SeekAndShowMembersOfGroup(group); }
                        else
                        {
                            DeleteDataTableQueryParameters(databasePerson, "PeopleGroupDesciption", "GroupPerson", dgSeek.values[1], "", "", "", "");

                            UpdateAmountAndRecepientOfPeopleGroupDesciption();
                            ShowDataTableDbQuery(databasePerson, "PeopleGroupDesciption", "SELECT GroupPerson AS 'Группа', GroupPersonDescription AS 'Описание группы', AmountStaffInDepartment AS 'Колличество сотрудников в группе' ", " group by GroupPerson ORDER BY GroupPerson asc; ");
                        }

                        textBoxGroup.BackColor = Color.White;
                        MembersGroupItem.Enabled = true;
                        break;
                    }
                case "Mailing":
                    {
                        dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] {
                            @"Получатель", @"Наименование", @"Дата создания/модификации",
                            @"Отчет по группам", @"Период", @"Тип отчета" });
                        DeleteDataTableQueryParameters(databasePerson, "Mailing",
                            "RecipientEmail", dgSeek.values[0],
                            "NameReport", dgSeek.values[1],
                            "DateCreated", dgSeek.values[2],
                            "GroupsReport", dgSeek.values[3],
                            "TypeReport", dgSeek.values[5],
                            "Period", dgSeek.values[4]);

                        ShowDataTableDbQuery(databasePerson, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
                        "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
                        "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
                        " ORDER BY RecipientEmail asc, DateCreated desc; ");
                        _toolStripStatusLabelSetText(StatusLabel2, "Удалена рассылка отчета " + dgSeek.values[1] + "| Всего рассылок: " + _dataGridView1RowsCount());
                        break;
                    }
                case "MailingException":
                    {
                        dgSeek.FindValuesInCurrentRow(dataGridView1, new string[] {
                            @"Получатель"});
                        DeleteDataTableQueryParameters(databasePerson, "MailingException",
                            "RecipientEmail", dgSeek.values[0]);

                        ShowDataTableDbQuery(databasePerson, "MailingException", "SELECT RecipientEmail AS 'Получатель', " +
                        "NameReport AS 'Наименование', Description AS 'Описание', DateCreated AS 'Дата создания/модификации', " +
                        "DayReport AS 'День отправки отчета'", " ORDER BY RecipientEmail asc, DateCreated desc; ");
                        _toolStripStatusLabelSetText(StatusLabel2, "Удален из исключений " + dgSeek.values[0] + "| Всего исключений: " + _dataGridView1RowsCount());
                        break;
                    }
                default:
                    break;
            }
            group = null;
        }

        //---  End.  DatagridView functions ---//



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

        private void ModeAppItem_Click(object sender, EventArgs e)  //ModeApp()
        { SwitchAppMode(); }

        private void SwitchAppMode()       // ExecuteAutoMode()
        {
            if (currentModeAppManual)
            {
                _MenuItemTextSet(ModeItem, "Выключить режим e-mail рассылок");
                _menuItemTooltipSet(ModeItem, "Включен автоматический режим. Выполняются Активные рассылки из БД.");
                _MenuItemBackColorChange(ModeItem, Color.DarkOrange);

                _toolStripStatusLabelSetText(StatusLabel2, "Включен режим рассылки отчетов по почте");
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
                _MenuItemTextSet(ModeItem, "Включить режим автоматических e-mail рассылок");
                _menuItemTooltipSet(ModeItem, "Включен интерактивный режим. Все рассылки остановлены.");
                _MenuItemBackColorChange(ModeItem, SystemColors.Control);

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
            long interval = 20 * 1000; //20 sek
            if (manualMode)
            {
                _MenuItemTextSet(ModeItem, "Выключить режим e-mail рассылок");
                _menuItemTooltipSet(ModeItem, "Включен автоматический режим. Выполняются Активные рассылки из БД.");
                _MenuItemBackColorChange(ModeItem, Color.DarkOrange);
                _toolStripStatusLabelSetText(StatusLabel2, "Включен режим авторассылки отчетов");
                _toolStripStatusLabelBackColor(StatusLabel2, Color.PaleGreen);

                timer?.Dispose();
                currentAction = "sendEmail";

                timer = new System.Threading.Timer(new System.Threading.TimerCallback(ScheduleTask), null, 0, interval);
            }
            else
            {
                _MenuItemTextSet(ModeItem, "Включить режим автоматических e-mail рассылок");
                _menuItemTooltipSet(ModeItem, "Включен интерактивный режим. Все рассылки остановлены.");
                _MenuItemBackColorChange(ModeItem, SystemColors.Control);

                _toolStripStatusLabelSetText(StatusLabel2, "Интерактивный режим");
                _toolStripStatusLabelBackColor(StatusLabel2, SystemColors.Control);

                timer?.Dispose();
            }
        }

        private void ScheduleTask(object obj) //SelectMailingDoAction()
        {
            lock (synclock)
            {
                DateTime dd = DateTime.Now;
                if (dd.Hour == 4 && dd.Minute == 10 && sent == false) //do something at Hour 2 and 5 minute //dd.Day == 1 && 
                {
                    _ProgressBar1Start();
                    _toolStripStatusLabelSetText(StatusLabel2, "Ведется работа по подготовке отчетов " + DateTime.Now.ToYYYYMMDDHHMM());
                    _toolStripStatusLabelBackColor(StatusLabel2, Color.LightPink);
                    CheckAliveIntellectServer(sServer1, sServer1UserName, sServer1UserPassword);
                    SelectMailingDoAction();
                    sent = true;
                    _ProgressBar1Stop();
                }
                else
                {
                    sent = false;
                }

                if (dd.Hour == 7 && dd.Minute == 1)
                {
                    _toolStripStatusLabelSetText(StatusLabel2, "Режим почтовых рассылок. " + DateTime.Now.ToYYYYMMDDHHMM());
                    _toolStripStatusLabelBackColor(StatusLabel2, Color.LightCyan);
                    ClearFilesInApplicationFolders(@"*.xlsx", "Excel-файлов");
                    _ProgressBar1Stop();
                }
            }
        }

        private async void TestToSendAllMailingsItem_Click(object sender, EventArgs e) //SelectMailingDoAction()
        {
            _ProgressBar1Start();
            await Task.Run(() => CheckAliveIntellectServer(sServer1, sServer1UserName, sServer1UserPassword));
            await Task.Run(() => SelectMailingDoAction());
            _ProgressBar1Stop();
        }

        private void SelectMailingDoAction() //MailingAction()
        {
            currentAction = "sendEmail";
            DoListsFioGroupsMailings();

            string sender = "";
            string recipient = "";
            string gproupsReport = "";
            string nameReport = "";
            string descriptionReport = "";
            string period = "";
            string status = "";
            string typeReport = "";
            string dayReport = "";

            List<MailingStructure> mailingList = new List<MailingStructure>();

            using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
            {
                sqlConnection.Open();

                using (var sqlCommand = new SQLiteCommand("SELECT * FROM Mailing;", sqlConnection))
                {
                    using (var reader = sqlCommand.ExecuteReader())
                    {
                        foreach (DbDataRecord record in reader)
                        {
                            if (record["SenderEmail"]?.ToString()?.Length > 0 && record["RecipientEmail"]?.ToString()?.Length > 0)
                            {
                                sender = record["SenderEmail"].ToString();
                                recipient = record["RecipientEmail"].ToString();
                                gproupsReport = record["GroupsReport"].ToString();
                                nameReport = record["NameReport"].ToString();
                                descriptionReport = record["Description"].ToString();
                                period = record["Period"].ToString();
                                status = record["Status"].ToString().ToLower();
                                typeReport = record["TypeReport"].ToString();
                                dayReport = record["DayReport"].ToString();

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
                                        _typeReport = typeReport,
                                        _status = status,
                                        _dayReport = dayReport
                                    });
                                }
                            }
                        }
                    }
                }
            }

            string str = "";
            foreach (MailingStructure mailng in mailingList)
            {
                _toolStripStatusLabelBackColor(StatusLabel2, SystemColors.Control);
                DateTime dt = DateTime.Now;
                int daySendReport = 0;
                bool isDayReport = false;
                isDayReport = Int32.TryParse(mailng._dayReport, out daySendReport);

                if (isDayReport && daySendReport == dt.Day) //send selected report only on inputed day
                {
                    _toolStripStatusLabelSetText(StatusLabel2, "Готовлю отчет " + mailng._nameReport);

                    str = "UPDATE 'Mailing' SET SendingLastDate='" + DateTime.Now.ToYYYYMMDDHHMM() +
                        "' WHERE RecipientEmail='" + mailng._recipient +
                        "' AND NameReport='" + mailng._nameReport +
                        "' AND GroupsReport ='" + mailng._groupsReport + "';";
                    logger.Info(str);
                    ExecuteSql(str, databasePerson);

                    GetRegistrationAndSendReport(
                        mailng._groupsReport, mailng._nameReport, mailng._descriptionReport, mailng._period, mailng._status,
                        mailng._typeReport, mailng._dayReport, true, mailng._recipient, mailServerUserName);
                }
            }

            ShowDataTableDbQuery(databasePerson, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
            "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
            "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
            " ORDER BY RecipientEmail asc, DateCreated desc; ");

            sender = null; recipient = null; gproupsReport = null; nameReport = null; descriptionReport = null; period = null; status = null;
            mailingList = null;
        }

        public string GetSafeFilename(string filename)
        {
            return string.Join("_", filename.Split(System.IO.Path.GetInvalidFileNameChars()));
        }

        private string SelectWholeCurrentMonth() //format of result: firstDayAndTime + "|" + lastDayAndTime
        {
            string result = "";
            var dt = DateTime.Now;

            var lastDayPrevMonth = new DateTime(dt.Year, dt.Month, dt.Day);

            result =
                lastDayPrevMonth.ToString("yyyy-MM") + "-01" + " 00:00:00" +
                "|" +
                lastDayPrevMonth.ToString("yyyy-MM-dd") + " 23:59:59";

            return result;
        }

        private string SelectWholePreviousMonth() //format of result: firstDay + "|" + lastDay
        {
            string result = "";
            var dt = DateTime.Now;
            var lastDayPrevMonth = new DateTime(dt.Year, dt.Month, 1).AddDays(-1);

            result =
                lastDayPrevMonth.ToString("yyyy-MM") + "-01" + " 00:00:00" +
                "|" +
                lastDayPrevMonth.ToString("yyyy-MM-dd") + " 23:59:59";

            return result;
        }

        private string SelectedDatetimePickersPeriodMonth() //format of result: firstDay + "|" + lastDay
        {
            return _dateTimePickerStart() + "|" + _dateTimePickerEnd();
        }

        private void MailingAction(string mainAction, string recipientEmail, string senderEmail, string groupsReport, string nameReport, string description, string period, string status, string typeReport, string dayReport)
        {
            _toolStripStatusLabelBackColor(StatusLabel2, SystemColors.Control);

            switch (mainAction)
            {
                case "saveEmail":
                    {
                        SaveMailing(mailServerUserName, senderEmail, groupsReport, nameReport, description, period, status, DateTime.Now.ToYYYYMMDDHHMM(), "", typeReport, dayReport);

                        ShowDataTableDbQuery(databasePerson, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
                        "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
                        "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
                        " ORDER BY RecipientEmail asc, DateCreated desc; ");
                        break;
                    }
                case "sendEmail":
                    {
                        CheckAliveIntellectServer(sServer1, sServer1UserName, sServer1UserPassword);

                        if (bServer1Exist)
                        {
                            GetRegistrationAndSendReport(groupsReport, nameReport, description, period, status, typeReport, dayReport, true, recipientEmail, senderEmail);
                        }
                        break;
                    }
                default:
                    break;
            }
        }

        private void GetRegistrationAndSendReport(string groupsReport, string nameReport, string description, string period, string status, string typeReport, string dayReport, bool sendReport, string recipientEmail, string senderEmail)
        {
            DataTable dtTempIntermediate = dtPeople.Clone();
            DateTime dtCurrent = DateTime.Today;
            PersonFull person = new PersonFull();

            GetNamePoints();  //Get names of the registration' points

            if (period.ToLower().Contains("текущий") || period.ToLower().Contains("текущая"))
            {
                reportStartDay = SelectWholeCurrentMonth().Split('|')[0];
                reportLastDay = SelectWholeCurrentMonth().Split('|')[1];
            }
            else if (period.ToLower().Contains("предыдущий") || period.ToLower().Contains("предыдущий"))
            {
                reportStartDay = SelectWholePreviousMonth().Split('|')[0];
                reportLastDay = SelectWholePreviousMonth().Split('|')[1];
            }
            else
            {
                reportStartDay = SelectedDatetimePickersPeriodMonth().Split('|')[0];
                reportLastDay = SelectedDatetimePickersPeriodMonth().Split('|')[1];
            }

            DateTime dt = DateTime.Parse(reportStartDay);
            int[] startPeriod = { dt.Year, dt.Month, dt.Day };
            dt = DateTime.Parse(reportLastDay);
            int[] endPeriod = { dt.Year, dt.Month, dt.Day };
            SeekAnualDays(ref dtTempIntermediate, ref person, false, startPeriod, endPeriod, ref myBoldedDates, ref workSelectedDays);
            logger.Trace(reportStartDay + " - " + reportLastDay);

            string nameGroup = "";
            string selectedPeriod = reportStartDay.Split(' ')[0] + " - " + reportLastDay.Split(' ')[0];

            string titleOfbodyMail = "";
            string[] groups = groupsReport.Split('+');

            foreach (string groupName in groups)
            {
                _toolStripStatusLabelBackColor(StatusLabel2, SystemColors.Control);
                nameGroup = groupName.Trim();
                if (nameGroup.Length > 0)
                {
                    dtPersonRegistrationsFullList?.Clear();
                    GetRegistrations(groupName, reportStartDay, reportLastDay, "sendEmail");//typeReport== only one group

                    dtTempIntermediate?.Clear();
                    dtTempIntermediate = dtPeople.Clone();

                    dtPersonTemp?.Clear();
                    dtPersonTemp = dtPeople.Clone();

                    person = new PersonFull();

                    dtPeopleGroup.Clear();
                    LoadGroupMembersFromDbToDataTable(nameGroup, ref dtPeopleGroup);

                    foreach (DataRow row in dtPeopleGroup.Rows)
                    {
                        if (row[FIO]?.ToString()?.Length > 0 && (row[GROUP]?.ToString() == nameGroup || (@"@" + row[DEPARTMENT_ID]?.ToString()) == nameGroup))
                        {
                            person = new PersonFull()

                            {
                                FIO = row[FIO].ToString(),
                                NAV = row[CODE].ToString(),

                                GroupPerson = row[GROUP].ToString(),
                                Department = row[DEPARTMENT].ToString(),
                                PositionInDepartment = row[EMPLOYEE_POSITION].ToString(),
                                DepartmentId = row[DEPARTMENT_ID].ToString(),
                                City = row[PLACE_EMPLOYEE].ToString(),

                                ControlInSeconds = ConvertStringTimeHHMMToSeconds(row[DESIRED_TIME_IN].ToString()),
                                ControlOutSeconds = ConvertStringTimeHHMMToSeconds(row[DESIRED_TIME_OUT].ToString()),
                                ControlInHHMM = row[DESIRED_TIME_IN].ToString(),
                                ControlOutHHMM = row[DESIRED_TIME_OUT].ToString(),

                                Comment = row[EMPLOYEE_SHIFT_COMMENT].ToString(),
                                Shift = row[EMPLOYEE_SHIFT].ToString()
                            };

                            FilterDataByNav(ref person, ref dtPersonRegistrationsFullList, ref dtTempIntermediate, typeReport);
                        }
                    }

                    logger.Trace("dtTempIntermediate: " + dtTempIntermediate.Rows.Count);
                    dtPersonTemp = GetDistinctRecords(dtTempIntermediate, orderColumnsFinacialReport);
                    dtPersonTemp.SetColumnsOrder(orderColumnsFinacialReport);
                    logger.Trace("dtPersonTemp: " + dtPersonTemp.Rows.Count);

                    if (dtPersonTemp.Rows.Count > 0)
                    {
                        string nameFile = nameReport + " " + reportStartDay.Split(' ')[0] + "-" + reportLastDay.Split(' ')[0] + " " + groupName + " от " + DateTime.Now.ToYYYYMMDDHHMM();
                        string illegal = GetSafeFilename(nameFile) + @".xlsx";
                        filePathExcelReport = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(filePathApplication), illegal);

                        logger.Trace("Подготавливаю отчет: " + filePathExcelReport);
                        ExportDatatableSelectedColumnsToExcel(ref dtPersonTemp, nameReport, ref filePathExcelReport);

                        if (sendReport)
                        {
                            if (reportExcelReady)
                            {
                                titleOfbodyMail = "с " + reportStartDay.Split(' ')[0] + " по " + reportLastDay.Split(' ')[0];
                                _toolStripStatusLabelSetText(StatusLabel2, "Выполняю отправку отчета адресату: " + recipientEmail);

                                SendEmail(senderEmail, recipientEmail, titleOfbodyMail, description, filePathExcelReport, Properties.Resources.LogoRYIK, productName);

                                _toolStripStatusLabelBackColor(StatusLabel2, Color.PaleGreen);
                                _toolStripStatusLabelSetText(StatusLabel2, DateTime.Now.ToYYYYMMDDHHMM() + " Отчет '" + nameReport + "'(" + groupName + ") подготовлен и отправлен " + recipientEmail);
                            }
                            else
                            {
                                _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                                _toolStripStatusLabelSetText(StatusLabel2, DateTime.Now.ToYYYYMMDDHHMM() + " Ошибка экспорта в файл отчета: " + nameReport + "(" + groupName + ")");
                            }
                        }
                    }
                    else
                    {
                        _toolStripStatusLabelSetText(StatusLabel2, DateTime.Now.ToYYYYMMDDHHMM() + "Ошибка получения данных для отчета: " + nameReport);
                        _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                    }
                }
            }
            titleOfbodyMail = null;

            dtTempIntermediate?.Dispose();
            dtPersonTemp?.Clear();
            selectedPeriod = null;
            nameGroup = null;
        }

        private static void SendEmail(string sender, string recipient, string period, string department, string pathToFile, Bitmap myLogo, string messageAfterPicture) //Compose and send e-mail
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
                    newMail.BodyEncoding = Encoding.UTF8;

                    newMail.AlternateViews.Add(getEmbeddedImage(period, department, myLogo, messageAfterPicture));
                    // отправитель - устанавливаем адрес и отображаемое в письме имя
                    newMail.From = new System.Net.Mail.MailAddress(sender, NAME_OF_SENDER_REPORTS);

                    // отправитель - устанавливаем адрес и отображаемое в письме имя
                    newMail.ReplyToList.Add(sender);

                    // Скрытая копия
                    // newMail.Bcc.Add("ry@ais.com.ua");

                    // кому отправляем
                    newMail.To.Add(new System.Net.Mail.MailAddress(recipient));
                    // тема письма
                    newMail.Subject = "Отчет по посещаемости за период: " + period;
                    // Проверка успешности отправки
                    smtpClient.SendCompleted += new System.Net.Mail.SendCompletedEventHandler(SendCompletedCallback);
                    // добавляем вложение
                    newMail.Attachments.Add(new System.Net.Mail.Attachment(pathToFile));
                    // логин и пароль
                    smtpClient.Credentials = new System.Net.NetworkCredential(sender.Split('@')[0], "");

                    // отправка письма
                    try
                    {
                        // async sending has a problem with sending
                        // smtpClient.SendAsync(newMail);
                        smtpClient.Send(newMail);
                        logger.Info("SendEmail: " + recipient + " - Ok");
                    } catch (Exception expt) { logger.Warn("SendEmail, Error: " + expt.Message); }

                    if (mailSent == false)
                    {
                        smtpClient.SendAsyncCancel();
                    }
                }
            }
        }

        static bool mailSent = false;

        //for async sending
        private static void SendCompletedCallback(object sender, System.ComponentModel.AsyncCompletedEventArgs e)
        {
            // Get the unique identifier for this asynchronous operation.
            String token = (string)e.UserState;

            if (e.Cancelled)
            { logger.Info("Send canceled " + token); }
            if (e.Error != null)
            { logger.Info("Send error " + e.Error.ToString()); }
            else
            { logger.Info("Message sent"); }
            mailSent = true;
        }

        private static System.Net.Mail.AlternateView getEmbeddedImage(string period, string department, Bitmap bmp, string messageAfterPicture)
        {
            //convert embedded resources into memorystream
            Bitmap b = new Bitmap(bmp, new Size(50, 50));
            ImageConverter ic = new ImageConverter();
            Byte[] ba = (Byte[])ic.ConvertTo(b, typeof(Byte[]));
            System.IO.MemoryStream logo = new System.IO.MemoryStream(ba);

            System.Net.Mail.LinkedResource res = new System.Net.Mail.LinkedResource(logo, "image/jpeg");
            res.ContentId = Guid.NewGuid().ToString();

            string message = @"<style type = 'text/css'> A {text - decoration: none;}</ style >" +
                @"<p><font size='3' color='black' face='Arial'>Здравствуйте,</p>Во вложении «Отчет по учету рабочего времени сотрудников».<p>" +

                @"<b>Период: </b>" +
                period +

                @"<br/><b>Подразделение: </b>'" +
                department +
                @"'<br/><p>Уважаемые руководители,</p><p>согласно Приказу ГК АИС «О функционировании процессов кадрового делопроизводства»,<br/><br/>" +
                @"<b>Внесите,</b> пожалуйста, <b>до конца текущего месяца</b> по сотрудникам подразделения " +
                @"информацию о командировках, больничных, отпусках, прогулах и т.п. <b>на сайт</b> www.ais .<br/><br/>" +
                @"Руководители <b>подразделений</b> ЦОК, <b>не отображающихся на сайте,<br/>вышлите, </b>пожалуйста, <b>Табель</b> учета рабочего времени<br/>" +
                @"<b>в отдел компенсаций и льгот до последнего рабочего дня месяца.</b><br/></p>" +
                @"<font size='3' color='black' face='Arial'>С, Уважением,<br/>" +
                NAME_OF_SENDER_REPORTS +
                @"</font><br/><br/><font size='2' color='black' face='Arial'><i>" +
                @"Данное сообщение и отчет созданы автоматически<br/>программой по учету рабочего времени сотрудников." +
                @"</i></font><br/><font size='1' color='red' face='Arial'><br/>" +
                DateTime.Now.ToYYYYMMDDHHMM() +
                @"</font></p><hr><img alt='ASTA' src='cid:" +
                res.ContentId +
                @"'/><br/><a href='mailto:ryik.yuri@gmail.com'>_</a>";

            System.Net.Mail.AlternateView alternateView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(message, null, System.Net.Mime.MediaTypeNames.Text.Html);
            alternateView.LinkedResources.Add(res);
            return alternateView;
        }

        //---  End. Schedule Functions ---//




        //---  Start. Block Encryption-Decryption ---//  

        private void TestCryptionItem_Click(object sender, EventArgs e)
        {
            string original = "Here is some data to encrypt!";
            MessageBox.Show("Original:   " + original);

            using (RijndaelManaged myRijndael = new RijndaelManaged())
            {
                myRijndael.Key = btsMess1;
                myRijndael.IV = btsMess2;
                // Encrypt the string to an array of bytes.
                byte[] encrypted = EncryptionDecryptionCriticalData.EncryptStringToBytes(original, myRijndael.Key, myRijndael.IV);

                StringBuilder s = new StringBuilder();
                foreach (byte item in encrypted)
                {
                    s.Append(item.ToString("X2") + " ");
                }
                MessageBox.Show("Encrypted:   " + s);

                // Decrypt the bytes to a string.
                string decrypted = EncryptionDecryptionCriticalData.DecryptStringFromBytes(encrypted, btsMess1, btsMess2);

                //Display the original data and the decrypted data.
                MessageBox.Show("Decrypted:    " + decrypted);
            }
        }

        //---  End. Block Encryption-Decryption ---//  





        private void ListBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            Font font = (sender as ListBox).Font;
            Brush backgroundColor;
            Brush textColor;

            if (e.Index % 2 != 0)
            {
                if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                {
                    backgroundColor = Brushes.SandyBrown;
                    textColor = Brushes.Black;
                }
                else
                {
                    backgroundColor = SystemBrushes.Window;
                    textColor = Brushes.Black;
                }
            }
            else
            {
                if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                {
                    backgroundColor = Brushes.SandyBrown;// SystemBrushes.Highlight;
                    textColor = textColor = Brushes.Black;// SystemBrushes.HighlightText;
                }
                else
                {
                    backgroundColor = Brushes.PaleTurquoise; // SystemBrushes.Window;
                    textColor = Brushes.Black;// SystemBrushes.WindowText;
                }
            }
            e.Graphics.FillRectangle(backgroundColor, e.Bounds);
            e.Graphics.DrawString((sender as ListBox).Items[e.Index].ToString(), font, textColor, e.Bounds);
        }

        private void ComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            Font font = (sender as ComboBox).Font;
            Brush backgroundColor;
            Brush textColor;

            if (e.Index % 2 != 0)
            {
                if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                {
                    backgroundColor = Brushes.SandyBrown;
                    textColor = Brushes.Black;
                }
                else
                {
                    backgroundColor = SystemBrushes.Window;
                    textColor = Brushes.Black;
                }
            }
            else
            {
                if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                {
                    backgroundColor = Brushes.SandyBrown;// SystemBrushes.Highlight;
                    textColor = textColor = Brushes.Black;// SystemBrushes.HighlightText;
                }
                else
                {
                    backgroundColor = Brushes.PaleTurquoise; // SystemBrushes.Window;
                    textColor = Brushes.Black;// SystemBrushes.WindowText;
                }
            }
            e.Graphics.FillRectangle(backgroundColor, e.Bounds);
            e.Graphics.DrawString((sender as ComboBox).Items[e.Index].ToString(), font, textColor, e.Bounds);
        }

        private void RemoveClickEvent(Button b) //clear all events in the button
        {
            System.Reflection.FieldInfo f1 = typeof(Control).GetField("EventClick",
                System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
            object obj = f1.GetValue(b);
            System.Reflection.PropertyInfo pi = b.GetType().GetProperty("Events",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            System.ComponentModel.EventHandlerList list = (System.ComponentModel.EventHandlerList)pi.GetValue(b, null);
            list.RemoveHandler(obj, list[obj]);
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

        private string _dataGridView1CellValue() //from other threads
        {
            string sDgv = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    try
                    {
                        sDgv = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[dataGridView1.CurrentCell.ColumnIndex]?.Value?.ToString()?.Trim();
                    } catch { sDgv = ""; }
                }));
            else
                try { sDgv = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[dataGridView1.CurrentCell.ColumnIndex]?.Value?.ToString()?.Trim(); } catch { sDgv = ""; }
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
                    if (dt?.Columns?.Count > 0 && dt?.Rows?.Count > 0)
                    {
                        dataGridView1.DataSource = dt;
                        dataGridView1.Visible = true;
                    }
                    else
                    {
                        System.Collections.ArrayList Empty = new System.Collections.ArrayList();
                        dataGridView1.DataSource = Empty;
                        dataGridView1?.Refresh();
                    }
                }));
            }
            else
            {
                if (dt?.Columns?.Count > 0 && dt?.Rows?.Count > 0)
                {
                    dataGridView1.DataSource = dt;
                    dataGridView1.Visible = true;
                }
                else
                {
                    System.Collections.ArrayList Empty = new System.Collections.ArrayList();
                    dataGridView1.DataSource = Empty;
                    dataGridView1?.Refresh();
                }
            }
        }


        private string _textBoxReturnText(TextBox txtBox) //add string into  from other threads
        {
            string tBox = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { tBox = txtBox?.Text?.ToString()?.Trim(); }));
            else
                tBox = txtBox?.Text?.ToString()?.Trim();
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
                    comboBx?.Items.Clear();
                    comboBx.SelectedText = "";
                    comboBx.Text = "";
                }));
            else
            {
                comboBx?.Items.Clear();
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

        private int _comboBoxCountItems(ComboBox comboBx) //from other threads
        {
            int count = 0;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { count = comboBx.Items.Count; }));
            else
            { count = comboBx.Items.Count; }
            return count;
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

        private string _listBoxReturnSelected(ListBox listBox) //from other threads
        {
            string result = "";
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate
                {
                    result = listBox.SelectedIndex == -1
                     ? listBox.Text.Trim()
                         : listBox.SelectedItem.ToString();
                }));
            }
            else
            {
                result = listBox.SelectedIndex == -1
                               ? listBox.Text.Trim()
                                   : listBox.SelectedItem.ToString();
            }
            return result;
        }

        private void _numUpDownSet(NumericUpDown numericUpDown, decimal i) //add string into comboBoxTargedPC from other threads
        {
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate { numericUpDown.Value = i; }));
            }
            else
            {
                numericUpDown.Value = i;
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
                    stringDT = dateTimePickerStart?.Value.Year.ToString("0000") + "-" + dateTimePickerStart?.Value.Month.ToString("00") + "-" + dateTimePickerStart.Value.Day.ToString("00") + " 00:00:00";
                }));
            else
            {
                stringDT = dateTimePickerStart?.Value.Year.ToString("0000") + "-" + dateTimePickerStart?.Value.Month.ToString("00") + "-" + dateTimePickerStart.Value.Day.ToString("00") + " 00:00:00";
            }
            return stringDT;
        }

        private string _dateTimePickerStartReturnMonth() //add string into  from other threads
        {
            string stringDT = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    stringDT = dateTimePickerStart?.Value.ToMonthName() + " " + dateTimePickerStart?.Value.Year;
                }));
            else
            {
                stringDT = dateTimePickerStart?.Value.ToMonthName() + " " + dateTimePickerStart?.Value.Year;
            }
            return stringDT;
        }
        private string _dateTimePickerEnd() //add string into  from other threads
        {
            string stringDT = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    stringDT = dateTimePickerEnd?.Value.Year.ToString("0000") + "-" + dateTimePickerEnd?.Value.Month.ToString("00") + "-" + dateTimePickerEnd.Value.Day.ToString("00") + " 23:59:59";
                }));
            else
            {
                stringDT = dateTimePickerEnd?.Value.Year.ToString("0000") + "-" + dateTimePickerEnd?.Value.Month.ToString("00") + "-" + dateTimePickerEnd.Value.Day.ToString("00") + " 23:59:59";
            }
            return stringDT;
        }

        private string _dateTimePickerReturn(DateTimePicker dateTimePicker) //add string into  from other threads
        {
            string result = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                { result = dateTimePicker?.Value.ToString(); }
                ));
            else
                result = dateTimePicker?.Value.ToString();
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
            stimerPrev = s;
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
                    panelView?.ResumeLayout();
                }));
            else
            {
                panelView?.ResumeLayout();
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
                    if (panelView?.Parent?.Height > 0)
                        height = panelView.Parent.Height;
                }));
            else
            {
                if (panelView?.Parent?.Height > 0)
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
                    if (panelView?.Height > 0)
                        height = panelView.Height;
                }));
            else
            {
                if (panelView?.Height > 0)
                    height = panelView.Height;
            }
            return height;
        }

        private int _panelWidthReturn(Panel panel) //access from other threads
        {
            int width = 0;
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate
                 {
                     if (panelView?.Width > 0)
                         width = panelView.Width;
                 }));
            }
            else
            {
                if (panelView?.Width > 0)
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
                    if (panelView?.Controls?.Count > 0)
                        count = panelView.Controls.Count;
                }));
            else
            {
                if (panelView?.Controls?.Count > 0)
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

        private string _MenuItemReturnText(ToolStripMenuItem menuItem) //access from other threads
        {
            string name = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    name = menuItem?.Text;
                }));
            else
            {
                name = menuItem?.Text;
            }
            return name;
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
                    if (control != null)
                        control.Dispose();
                }));
            else
            {
                if (control != null)
                    control.Dispose();
            }
        }

        private void _changeControlBackColor(Control control, Color color) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { control.BackColor = color; }));
            else
                control.BackColor = color; ;
        }

        private void timer1_Tick(object sender, EventArgs e) //Change a Color of the Font on Status by the Timer
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    if (StatusLabel2.ForeColor == Color.DarkBlue)
                    {
                        StatusLabel2.ForeColor = Color.DarkRed;
                        StatusLabel2.Text = stimerCurr;
                    }
                    else
                    {
                        StatusLabel2.ForeColor = Color.DarkBlue;
                        StatusLabel2.Text = stimerPrev;
                    }
                }));
            else
            {
                if (StatusLabel2.ForeColor == Color.DarkBlue)
                {
                    StatusLabel2.ForeColor = Color.DarkRed;
                    StatusLabel2.Text = stimerCurr;
                }
                else
                {
                    StatusLabel2.ForeColor = Color.DarkBlue;
                    StatusLabel2.Text = stimerPrev;
                }
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

        private int ConvertDecimalSeparatedTimeToSeconds(decimal decimalHour, decimal decimalMinute)
        {
            int result = Convert.ToInt32(decimalHour * 60 * 60 + decimalMinute * 60);
            return result;
        }

        private decimal ConvertStringsTimeToDecimal(string hour, string minute)
        {
            decimal result = TryParseStringToDecimal(hour) + TryParseStringToDecimal(TimeSpan.FromMinutes(TryParseStringToDouble(minute)).TotalHours.ToString());
            return result;
        }

        private string ConvertSecondsToStringHHMM(int seconds)
        {
            string result;
            int hours = seconds / 3600;
            int minutes = (seconds % 3600) / 60;

            result = string.Format("{0:d2}:{1:d2}", hours, minutes);
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

            return String.Format("{0:d2}:{1:d2}", h, m);
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
            decimal[] result = new decimal[5];
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
            result[4] = 60 * 60 * result[0] + 60 * result[1];                                    // result in seconds       1335

            return result;
        }

        private int ConvertStringTimeHHMMToSeconds(string timeInHHMM) //time HH:MM converted to decimal value
        {
            string hours = "0";
            string minutes = "0";
            string seconds = "0";
            int length = timeInHHMM.Split(':').Length;

            if (length > 2)
            {
                string[] time = timeInHHMM.Split(':');
                hours = time[0];
                minutes = time[1];
                seconds = time[2];
            }
            else if (length == 2)
            {
                string[] time = timeInHHMM.Split(':');
                hours = time[0];
                minutes = time[1];
            }
            else if (length == 1)
            {
                hours = timeInHHMM;
            }

            return (60 * 60 * Convert.ToInt32(hours) + 60 * Convert.ToInt32(minutes) + Convert.ToInt32(seconds));
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

        private int[] ConvertStringDateToIntArray(string dateYYYYmmDD) //date "YYYY-MM-DD HH:MM" to int array values
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


        public static string DateTimeToYYYYMMDD(string date = "")
        {
            if (date.Length > 0)
            { return DateTime.Parse(date).ToString("yyyy-MM-dd"); }
            else { return DateTime.Now.ToString("yyyy-MM-dd"); }
        }

        public static string DateTimeToYYYYMMDDHHMM(string date = "")
        {
            if (date.Length > 0)
            { return DateTime.Parse(date).ToString("yyyy-MM-dd HH:mm"); }
            else { return DateTime.Now.ToString("yyyy-MM-dd HH:mm"); }
        }
        //---- End. Convertors of data types ----//



        private void testADToolStripMenuItem_Click(object sender, EventArgs e)
        { GetUsersFromAD(); }

        static List<StaffAD> staffAD = new List<StaffAD>();

        private void GetUsersFromAD()
        {
            logger.Trace("GetUsersFromAD: ");
            string user = null;
            string password = null;
            string domain = null;
            string server = null;

            ActiveDirectoryGetData ad;
            StaffListStore staffListStore = new StaffListStore();
            ListStaffSender listStaffSender = new ListStaffSender();
            staffAD = new List<StaffAD>();

            listParameters = new List<ParameterConfig>();
            ParameterOfConfigurationInSQLiteDB parameters = new ParameterOfConfigurationInSQLiteDB();
            parameters.databasePerson = databasePerson;
            listParameters = parameters.GetParameters("%%").FindAll(x => x.isExample == "no"); //load only real data

            try
            {
                user = listParameters.FindLast(x => x.parameterName == @"UserName").parameterValue;
                password = listParameters.FindLast(x => x.parameterName == @"UserPassword").parameterValue;
                server = listParameters.FindLast(x => x.parameterName == @"ServerURI").parameterValue;
                domain = listParameters.FindLast(x => x.parameterName == @"DomainOfUser").parameterValue;
            } catch (Exception expt) { logger.Info(expt.ToString()); }
            logger.Trace("user, domain, password, server: " + user + " |" + domain + " |" + password + " |" + server);

            if (user.Length > 0 && password.Length > 0 && domain.Length > 0 && server.Length > 0)
            {
                ad = new ActiveDirectoryGetData(user, domain, password, server);
                staffListStore.Story.Push(ad.SaveListStaff());
                listStaffSender.RestoreListStaff(staffListStore.Story.Pop());

                staffAD = listStaffSender.GetListStaff();

                logger.Trace("GetUsersFromAD: Store list ");
                //передать дальше в обработку:
                foreach (var person in staffAD)
                {
                    logger.Trace(person.fio + " |" + person.login + " |" + person.code + " |" + person.mail);
                }
            }
            ad = null; listStaffSender = null; staffListStore = null; parameters = null; listParameters = null;
        }

        /*
        private void GetConfigTableFromDB(string tableName, ref List<ObjectsOfConfig> parametersList)
        {
            string parameter, value;

            if (databasePerson.Exists)
            {
                using (var sqlConnection = new SQLiteConnection($"Data Source={databasePerson};Version=3;"))
                {
                    sqlConnection.Open();
                    using (var sqlCommand = new SQLiteCommand("SELECT ParameterName, Value FROM ConfigDB;", sqlConnection))
                    {
                        using (var reader = sqlCommand.ExecuteReader())
                        {
                            foreach (DbDataRecord record in reader)
                            {
                                parameter = record["ParameterName"]?.ToString();
                                value = record["Value"]?.ToString();

                                if (parameter?.Length > 0 && value?.Length > 0)
                                {
                                    parametersList.Add(
                                        new ObjectsOfConfig
                                        {
                                            _parameter = parameter,
                                            _value = value
                                        });
                                }
                            }
                        }
                    }
                }
            }
            parameter = value = null;
        }
        */

    }
}
