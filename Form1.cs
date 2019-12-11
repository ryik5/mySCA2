//using NLog;
//Project\Control NuGet\console
//install-package nlog
//install-package nlog.config

using ASTA.Classes;
using ASTA.Classes.People;
using ASTA.Classes.Security;
using ASTA.Classes.Updating;
using AutoUpdaterDotNET; //Updater
using MimeKit; //Mailing
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.Data.Common;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ASTA
{
    public partial class WinFormASTA : Form
    {
        //todo!!!!!!!!!
        //Check of All variables, const and controls
        //they will be needed to Remove if they are not needed
        private DataGridViewOperations dgvo;

        //logging
        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        // static string method = System.Reflection.MethodBase.GetCurrentMethod().Name;
        // logger.Trace("-= " + method + " =-");

        //System settings
        private static readonly string guid = System.Runtime.InteropServices.Marshal.GetTypeLibGuidForAssembly(System.Reflection.Assembly.GetExecutingAssembly()).ToString(); // получаем GIUD приложения// получаем GIUD приложения

        private static readonly string appVersionAssembly = System.Reflection.Assembly.GetEntryAssembly().GetName().Version.ToString();//Application.ProductVersion;//

        private static readonly string appCopyright = Application.CompanyName;// appFileVersionInfo.LegalCopyright;
        private static readonly string appName = Application.ProductName;// appFileVersionInfo.ProductName;;
        private static readonly string fileNameToUpdateXML = $"{appName}.xml";
        private static readonly string fileNameToUpdateZip = $"{appName}.zip";

        private static readonly string filePathApp = Application.ExecutablePath;// System.Reflection.Assembly.GetEntryAssembly().Location;// Application.ExecutablePath;
        private static readonly System.Diagnostics.FileVersionInfo appFileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(filePathApp);

        private static readonly string localAppFolderPath = Application.StartupPath; //Environment.CurrentDirectory
        private static readonly string pathToQueryToCreateMainDb = System.IO.Path.Combine(localAppFolderPath, $"{appName}.sql"); //System.IO.Path.GetFileNameWithoutExtension(appFilePath)
        private static readonly string appFolderTempPath = System.IO.Path.Combine(localAppFolderPath, "Temp");
        private static readonly string appFolderUpdatePath = System.IO.Path.Combine(localAppFolderPath, "Update");
        private static readonly string appFolderBackUpPath = System.IO.Path.Combine(localAppFolderPath, "Backup");

        private static readonly string[] appAllFiles =  {
                @"ASTA.exe", @"NLog.config", @"NLog.dll", @"ASTA.sql",
                @"MailKit.dll", @"MimeKit.dll", @"MySql.Data.dll",
                @"x64\SQLite.Interop.dll", @"x86\SQLite.Interop.dll", @"System.Data.SQLite.dll", 
                @"AutoUpdater.NET.dll", @"Google.Protobuf.dll",
                @"Renci.SshNet.dll", @"BouncyCastle.Crypto.dll",

                @"EntityFramework.dll",
                @"EntityFramework.SqlServer.dll",
                @"Microsoft.Data.SqlClient.dll",
                @"Microsoft.DotNet.PlatformAbstractions.dll",
                @"System.Diagnostics.DiagnosticSource.dll",
                @"System.ComponentModel.Annotations.dll",
                @"System.Buffers.dll",
                @"System.Memory.dll",
                @"System.Runtime.CompilerServices.Unsafe.dll",
                @"System.Threading.Tasks.Extensions.dll",

                //Abstraction for static System.IO library
                @"System.IO.Abstractions.dll",
                //Analysing packages
                @"ASTA.exe.config"
                        };

        private static readonly string appRegistryKey = @"SOFTWARE\RYIK\ASTA";

        private static readonly byte[] keyEncryption = Convert.FromBase64String(@"OCvesunvXXsxtt381jr7vp3+UCwDbE4ebdiL1uinGi0="); //Key Encrypt
        private static readonly byte[] keyDencryption = Convert.FromBase64String(@"NO6GC6Zjl934Eh8MAJWuKQ=="); //Key Decrypt

        private static readonly System.IO.FileInfo dbApplication = new System.IO.FileInfo(System.IO.Path.Combine(localAppFolderPath, @"main.db"));
        private static readonly string appDbPath = dbApplication.FullName;
        private static readonly string appDbName = System.IO.Path.GetFileName(appDbPath);

        private static readonly string sqLiteLocalConnectionString = $"Data Source = {appDbPath}; Version=3;"; ////$"Data Source={dbApplication.FullName};Version=3;" ////$"Data Source={dbApplication.FullName};Version=3;"

        private static string sqlServerConnectionString;// = "Data Source=" + serverName + "\\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID=" + userName + ";Password=" + userPasswords + "; Connect Timeout=5";
        private static System.Data.Entity.Core.EntityClient.EntityConnectionStringBuilder sqlServerConnectionStringEF;
        private static string mysqlServerConnectionString;//@"server=" + mysqlServer + @";User=" + mysqlServerUserName + @";Password=" + mysqlServerUserPassword + @";database=wwwais;convert zero datetime=True;Connect Timeout=60";

        private static readonly string queryCheckSystemDdSCA = "SELECT database_id FROM sys.databases WHERE Name ='intellect'";

        private static string remoteFolderUpdateURL; // format - server.domain.subdomain/folder  or   server/folder

        private static bool uploadingStatus = false;
        private string appFileMD5;
        private string domain;

        //todo
        //Все константы в локальную БД
        //todo
        //преобразовать в дикшенери сущностей с дефолтовыми значениями и описаниями параметров
        //  Dictionary<string, ParameterOfConfiguration> config = new Dictionary<string, ParameterOfConfiguration>();
        private static string sLastSelectedElement;

        private static string nameOfLastTable;
        private static string currentAction = "";
        private static bool currentModeAppManual;

        // taskbar and logo
        private static Bitmap bmpLogo;

        private NotifyIcon notifyIcon;
        private ContextMenu contextMenu;
        private bool buttonAboutForm;
        private static Byte[] byteLogo;

        //context Menu on Datagrid
        private ContextMenu mRightClick;

        private int iCounterLine = 0;

        //collecting of data
        private static List<Employee> listFIO = new List<Employee>(); // List of FIO and identity of data

        //Controls "NumUpDown"
        private decimal numUpHourStart = 9;

        private decimal numUpMinuteStart = 0;
        private decimal numUpHourEnd = 18;
        private decimal numUpMinuteEnd = 0;

        //View of registrations
        private PictureBox pictureBox1 = new PictureBox();

        private Bitmap bmp = new Bitmap(1, 1);

        //reports
        private const int offsetTimeIn = 60;    // смещение времени прихода после учетного, в секундах, в течении которого не учитывается опоздание

        private const int offsetTimeOut = 60;   // смещение времени ухода до учетного, в секундах в течении которого не учитывается ранний уход
        private string[] myBoldedDates;
        private string[] workSelectedDays;
        private static string reportStartDay = "";
        private static string reportLastDay = "";
        private bool reportExcelReady = false;
        private string filePathExcelReport;

        //mailing
        private const string NAME_OF_SENDER_REPORTS = @"Отдел компенсаций и льгот";

        private const string START_OF_MONTH = @"START_OF_MONTH";
        private const string MIDDLE_OF_MONTH = @"MIDDLE_OF_MONTH";
        private const string END_OF_MONTH = @"END_OF_MONTH";
        private static System.Threading.Timer timer;
        private static readonly object synclock = new object();
        private static bool sent = false;
        private static string DEFAULT_DAY_OF_SENDING_REPORT = @"END_OF_MONTH";
        private static int ShiftDaysBackOfSendingFromLastWorkDay = 3; //shift back of sending email before a last working day within the month

        //Page of Mailing

        private Label labelMailServerName;
        private TextBox textBoxMailServerName;
        private Label labelMailServerUserName;
        private TextBox textBoxMailServerUserName;
        private Label labelMailServerUserPassword;
        private TextBox textBoxMailServerUserPassword;
        private static string mailServer = "";
        private static string mailServerDB = "";
        private static int mailServerSMTPPort = 25;
        private static string mailServerSMTPPortDB = "";
        private static string mailSenderAddress = "";
        private static string mailsOfSenderOfNameDB = "";
        private static string mailsOfSenderOfPassword = "";
        private static string mailsOfSenderOfPasswordDB = "";

        private static MailServer _mailServer;
        private static MailUser _mailUser;
        private static string mailJobReportsOfNameOfReceiver = ""; //Receiver of job reports
        private static List<Mailing> resultOfSendingReports = new List<Mailing>();
        //  static System.Net.Mail.LinkedResource mailLogo;

        //Page of "Settings' Programm"
        private bool bServer1Exist = false;

        private Label labelServer1;
        private TextBox textBoxServer1;
        private Label labelServer1UserName;
        private TextBox textBoxServer1UserName;
        private Label labelServer1UserPassword;
        private TextBox textBoxServer1UserPassword;
        private string sServer1;
        private string sServer1Registry;
        private string sServer1DB;
        private string sServer1UserName;
        private string sServer1UserNameRegistry;
        private string sServer1UserNameDB;
        private string sServer1UserPassword;
        private string sServer1UserPasswordRegistry;
        private string sServer1UserPasswordDB;

        private Label labelmysqlServer;
        private TextBox textBoxmysqlServer;
        private Label labelmysqlServerUserName;
        private TextBox textBoxmysqlServerUserName;
        private Label labelmysqlServerUserPassword;
        private TextBox textBoxmysqlServerUserPassword;
        private static string mysqlServer;
        private static string mysqlServerRegistry;
        private static string mysqlServerDB;
        private static string mysqlServerUserName;
        private static string mysqlServerUserNameRegistry;
        private static string mysqlServerUserNameDB;
        private static string mysqlServerUserPassword;
        private static string mysqlServerUserPasswordRegistry;
        private static string mysqlServerUserPasswordDB;

        private Label listComboLabel;
        private ComboBox listCombo;

        private Label periodComboLabel;
        private ListBox listboxPeriod;

        private Label labelSettings9;
        private ComboBox comboSettings9;

        private Label labelSettings15; //type report
        private ComboBox comboSettings15;

        private Label labelSettings16;
        private TextBox textBoxSettings16;

        private CheckBox checkBox1;
        private static List<ParameterConfig> listParameters = new List<ParameterConfig>();

        private static Color clrRealRegistration = Color.PaleGreen;
        private Color clrRealRegistrationRegistry = Color.PaleGreen;

        private int countGroups = 0;
        private int countUsers = 0;
        private int countMailers = 0;

        private int numberPeopleInLoading = 1;

        //todo
        private static string stimerPrev = "";

        private static string stimerCurr = "Ждите!";

        private static List<ADUserFullAccount> listUsersAD = new List<ADUserFullAccount>(); //Users of AD. Got data from Domain

        private static List<PeopleOutReasons> outResons = new List<PeopleOutReasons>();
        private static List<OutPerson> outPerson = new List<OutPerson>();
        private static List<PeopleShift> peopleShifts = new List<PeopleShift>();

        //DataTables with people data
        private DataTable dtPeople = new DataTable("People");

        private DataColumn[] dcPeople =
           {
                                  new DataColumn(Names.NPP,typeof(int)),//0
                                  new DataColumn(Names.FIO,typeof(string)),//1
                                  new DataColumn(Names.CODE,typeof(string)),//2
                                  new DataColumn(Names.GROUP,typeof(string)),//3
                                  new DataColumn(Names.N_ID,typeof(int)), //6
                                  new DataColumn(Names.N_ID_STRING,typeof(string)), //6
                                  new DataColumn(Names.DEPARTMENT,typeof(string)),//7
                                  new DataColumn(Names.PLACE_EMPLOYEE,typeof(string)),//8
                                  new DataColumn(Names.DATE_REGISTRATION,typeof(string)),//9
                                  new DataColumn(Names.TIME_REGISTRATION,typeof(int)), //10
                                  new DataColumn(Names.TIME_REGISTRATION_STRING,typeof(string)), //10
                                  new DataColumn(Names.REAL_TIME_IN,typeof(string)),//16
                                  new DataColumn(Names.REAL_TIME_OUT,typeof(string)), //17
                                  new DataColumn(Names.SERVER_SKD,typeof(string)), //11
                                  new DataColumn(Names.CHECKPOINT_NAME,typeof(string)), //12
                                  new DataColumn(Names.CHECKPOINT_DIRECTION,typeof(string)), //12
                                  new DataColumn(Names.CHECKPOINT_ACTION,typeof(string)), //12
                                  new DataColumn(Names.DESIRED_TIME_IN,typeof(string)),//14
                                  new DataColumn(Names.DESIRED_TIME_OUT,typeof(string)),//15
                                  new DataColumn(Names.EMPLOYEE_TIME_SPENT,typeof(int)), //18
                                  new DataColumn(Names.EMPLOYEE_PLAN_TIME_WORKED,typeof(string)), //19
                                  new DataColumn(Names.EMPLOYEE_BEING_LATE,typeof(string)),                    //20
                                  new DataColumn(Names.EMPLOYEE_EARLY_DEPARTURE,typeof(string)),                 //21
                                  new DataColumn(Names.EMPLOYEE_VACATION,typeof(string)),                 //22
                                  new DataColumn(Names.EMPLOYEE_TRIP,typeof(string)),                 //23
                                  new DataColumn(Names.DAY_OF_WEEK,typeof(string)),                 //24
                                  new DataColumn(Names.EMPLOYEE_SICK_LEAVE,typeof(string)),                 //25
                                  new DataColumn(Names.EMPLOYEE_ABSENCE,typeof(string)),     //26
                                  new DataColumn(Names.GROUP_DECRIPTION,typeof(string)),            //27
                                  new DataColumn(Names.EMPLOYEE_SHIFT_COMMENT,typeof(string)),                 //28
                                  new DataColumn(Names.EMPLOYEE_POSITION,typeof(string)),                 //29
                                  new DataColumn(Names.EMPLOYEE_SHIFT,typeof(string)),                 //30
                                  new DataColumn(Names.EMPLOYEE_HOOKY,typeof(string)),                 //31
                                  new DataColumn(Names.DEPARTMENT_ID,typeof(string)), //32
                                  new DataColumn(Names.CHIEF_ID,typeof(string)), //33
                                  new DataColumn(Names.CARD_STATE,typeof(string)) //34
                };

        private DataTable dtPersonTemp = new DataTable("PersonTemp");
        private DataTable dtPersonTempAllColumns = new DataTable("PersonTempAllColumns");
        private DataTable dtPersonRegistrationsFullList = new DataTable("PersonRegistrationsFullList");
        private DataTable dtPeopleGroup = new DataTable("PeopleGroup");
        private DataTable dtPeopleListLoaded = new DataTable("PeopleLoaded");

        //Color of User's Control elements which depend on the selected MenuItem
        private Color labelGroupCurrentBackColor;

        private Color textBoxGroupCurrentBackColor;
        private Color labelGroupDescriptionCurrentBackColor;
        private Color textBoxGroupDescriptionCurrentBackColor;
        private Color comboBoxFioCurrentBackColor;
        private Color textBoxFIOCurrentBackColor;
        private Color textBoxNavCurrentBackColor;

        private static CollectionOfPassagePoints collectionOfPassagePoints;

        //Drawing //DrawFullWorkedPeriodRegistration
        //  int iPanelBorder = 2;
        private static readonly int iShiftStart = 300;

        private static readonly int iShiftHeightAll = 36;
        private static readonly int iOffsetBetweenHorizontalLines = 19; //смещение между горизонтальными линиями
        private static readonly int iOffsetBetweenVerticalLines = 60; //смещение между "часовыми" линиями
        private static readonly int iNumbersOfHoursInDay = 24;        //количество часов в сутках(количество вертикальных часовых линий)
        private static readonly int iHeightLineWork = 4; //толщина линии рабочего времени на графике
        private static readonly int iHeightLineRealWork = 14; //толщина линии фактически отработанного веремени на графике

        public WinFormASTA()
        { InitializeComponent(); }

        private void Form1_Load(object sender, EventArgs e)
        { Form1Load(); }

        private void Form1Load()
        {
            string statusInfo = appName + " ver." + appVersionAssembly + " by " + appCopyright;

            sLastSelectedElement = "MainForm";
            nameOfLastTable = "PersonRegistrationsList";
            SetStatusLabelText(StatusLabel1, "");
            SetStatusLabelText(StatusLabel2, "");

            //Начало лога
            logger.Info("");
            logger.Info("");
            logger.Info("-= " + statusInfo + " =-");
            logger.Info("");
            logger.Info("");

            //Блок проверки уровня настройки логгирования
            logger.Info("Test Info message");
            logger.Trace("Test1 Trace message");
            logger.Debug("Test2 Debug message");
            logger.Warn("Test3 Warn message");
            logger.Error("Test4 Error message");
            logger.Fatal("Test5 Fatal message");

            //Настройка отображаемых пунктов меню и других элементов интерфеса
            currentModeAppManual = true;
            System.IO.Directory.CreateDirectory(appFolderBackUpPath);
            System.IO.Directory.CreateDirectory(appFolderTempPath);
            System.IO.Directory.CreateDirectory(appFolderUpdatePath);

            ClearItemsInFolder(appFolderTempPath);   //Clear temporary folder
            ClearItemsInFolder(appFolderUpdatePath); //System.IO.Path.Combine(appFolderUpdatePath, @"*.*")
            ClearItemsInFolder(System.IO.Path.Combine(localAppFolderPath, fileNameToUpdateZip));

            //Make archives:
            //1. from app's *.exe and main lib files of the app
            MakeZip(appAllFiles, fileNameToUpdateZip);
            if (System.IO.File.Exists(fileNameToUpdateZip))
            {
                System.IO.File.Move(fileNameToUpdateZip, System.IO.Path.Combine(appFolderBackUpPath, appName + "." + GetSafeFilename(DateTime.Now.ToYYYYMMDDHHMMSS(), "") + @".zip"));
            }
            //refresh temp folder
            ClearItemsInFolder(appFolderTempPath); //+ @"\*.*"

            //2. from the main DB of application
            string dbZipPath = appDbName + "." + GetSafeFilename(DateTime.Now.ToYYYYMMDDHHMMSS(), "") + @".zip";
            ClearItemsInFolder(System.IO.Path.Combine(localAppFolderPath, dbZipPath));//localAppFolderPath + @"\" + dbZipPath
            MakeZip(appDbName, dbZipPath);
            System.IO.File.Move(dbZipPath, System.IO.Path.Combine(appFolderBackUpPath, System.IO.Path.GetFileName(dbZipPath)));

            //refresh temp folder
            ClearItemsInFolder(appFolderTempPath); //+ @"\*.*"

            //Check local DB Configuration
            logger.Trace("Проверка локальной БД");
            if (!System.IO.File.Exists(dbApplication.FullName))
            {
                logger.Info("Создаю файл локальной DB");
                SQLiteConnection.CreateFile(dbApplication.FullName);
            }

            DbSchema schemaDB = DbSchema.LoadDB(dbApplication.FullName);
            bool currentDbEmpty = schemaDB?.Tables?.Count > 0 ? false : true;
            logger.Trace("current db has: " + schemaDB?.Tables?.Count.ToString() + " tables");
            foreach (var table in schemaDB.Tables)
            { logger.Trace("the table is existed: " + table.Key); }

            string txtSQLs = String.Join(" ", ReadTXTFile(pathToQueryToCreateMainDb)); //List<string> -> a single string
            DbSchema schemaTXT = DbSchema.ParseSql(txtSQLs);
            logger.Trace("txtSQL is wanted to make: " + schemaTXT?.Tables?.Count + " tables from: " + pathToQueryToCreateMainDb);

            foreach (var table in schemaTXT.Tables)
            { logger.Trace("the table is wanted: " + table.Key); }

            bool equalDB = schemaTXT.Equals(schemaDB);
            if (equalDB) { logger.Trace("Схема конфигурации текущей БД соответствует схеме загруженной с файла: " + pathToQueryToCreateMainDb); }
            else { logger.Info("Схема конфигурации текущей БД Отличается от схеме загруженной с файла: " + pathToQueryToCreateMainDb); }

            logger.Trace("tables in loaded DB: " + schemaDB?.Tables?.Count + ", " + " must be tables: " + schemaTXT?.Tables?.Count);

            if (currentDbEmpty || !equalDB || !(schemaDB.Tables.Equals(schemaTXT.Tables)))
            {
                logger.Info("Заполняю схему локальной DB");
                TryMakeLocalDB(pathToQueryToCreateMainDb);
            }

            //Refresh Configuration of the application
            AddExceptedParametersIntoConfigurationDb();

            logger.Info("Загружаю/проверяю настройки программы...");

            LoadPreviouslySavedParameters();
            logger.Info("Вычисляю ближайшие праздничные и выходные дни...");
            DataTable dtEmpty = new DataTable();
            EmployeeFull personEmpty = new EmployeeFull();
            var startDay = DateTime.Now.AddDays(-60).ToIntYYYYMMDD();
            var endDay = DateTime.Now.AddDays(30).ToIntYYYYMMDD();

            SeekAnualDays(ref dtEmpty, ref personEmpty, false, startDay, endDay,
                ref myBoldedDates, ref workSelectedDays);

            dtEmpty?.Dispose();
            personEmpty = null;

            dgvo = new DataGridViewOperations();
            monthCalendar.SelectionStart = DateTime.Now;
            monthCalendar.SelectionEnd = DateTime.Now;
            monthCalendar.Update();
            monthCalendar.Refresh();

            logger.Info("Настраиваю переменные....");
            //Prepare DataTables
            dtPeople.Columns.AddRange(dcPeople);
            dtPeople.DefaultView.Sort = Names.GROUP + ", " + Names.FIO + ", " + Names.DATE_REGISTRATION + ", " + Names.TIME_REGISTRATION + ", " + Names.REAL_TIME_IN + ", " + Names.REAL_TIME_OUT + " ASC";
            // dtPeople.DefaultView.Sort = "[Группа], [Фамилия Имя Отчество], [Дата регистрации], [Время регистрации], [Фактич. время прихода ЧЧ:ММ:СС], [Фактич. время ухода ЧЧ:ММ:СС] ASC";

            //Clone default column name and structure from 'dtPeople' to other DataTables
            dtPersonTemp = dtPeople.Clone();  //Copy only structure(Name of columns)
            dtPersonRegistrationsFullList = dtPeople.Clone();  //Copy only structure(Name of columns)
            dtPeopleGroup = dtPeople.Clone();  //Copy only structure(Name of columns)

            logger.Trace("SetTechInfoIntoDB");
            SetTechInfoIntoDB();

            logger.Info("Настраиваю интерфейс....");
            bmpLogo = Properties.Resources.LogoRYIK;
            MakeByteLogo(bmpLogo); //logo for mailing

            this.Icon = Icon.FromHandle(bmpLogo.GetHicon());
            this.Text = appFileVersionInfo.Comments;

            contextMenu = new ContextMenu();  //Context Menu on notify Icon
            contextMenu.MenuItems.Add("About", AboutSoft);
            contextMenu.MenuItems.Add("-", AboutSoft);
            contextMenu.MenuItems.Add("Exit", ApplicationExit);

            notifyIcon = new NotifyIcon
            {
                Icon = this.Icon,
                Visible = true,
                BalloonTipText = "Developed by " + appCopyright
            };
            notifyIcon.ShowBalloonTip(500);

            // notifyIcon.Text = Application.ProductName + "\nv." + Application.ProductVersion + " (" + appFileVersionInfo.FileVersion + ")" + "\n" + Application.CompanyName;
            notifyIcon.Text = appName + "\nv." + appVersionAssembly + " (" + appFileVersionInfo.FileVersion + ")" + "\n" + appFileVersionInfo.CompanyName;
            notifyIcon.ContextMenu = contextMenu;

            //Set up Starting values
            dateTimePickerStart.CustomFormat = "yyyy-MM-dd";
            dateTimePickerEnd.CustomFormat = "yyyy-MM-dd";
            dateTimePickerStart.Format = DateTimePickerFormat.Custom;
            dateTimePickerEnd.Format = DateTimePickerFormat.Custom;
            dateTimePickerStart.Value = DateTime.Now.FirstDayOfMonth();
            dateTimePickerEnd.Value = DateTime.Now.LastDayOfMonth();

            numUpDownHourStart.Value = 9;
            numUpDownMinuteStart.Value = 0;
            numUpDownHourEnd.Value = 18;
            numUpDownMinuteEnd.Value = 0;

            StatusLabel1.Text = statusInfo;
            StatusLabel1.Alignment = ToolStripItemAlignment.Right;

            /////////////////////////////////
            ///SetUp Menu Items
            //todo
            //rewrite to access from other threads
            CheckBoxesFiltersAll_Enable(false);

            EnableMenuItem(AddAnualDateItem, false);

            MembersGroupItem.Enabled = false;
            AddPersonToGroupItem.Enabled = false;
            CreateGroupItem.Enabled = false;

            DeleteGroupItem.Visible = false;
            TableModeItem.Visible = false;
            VisualModeItem.Visible = false;
            ChangeColorMenuItem.Visible = false;
            TableExportToExcelItem.Visible = false;
            listFioItem.Visible = false;
            groupBoxProperties.Visible = false;

            dataGridView1.ShowCellToolTips = true;

            //Colorizing of Menu Items
            GetFioItem.BackColor = Color.LightSkyBlue;
            HelpAboutItem.BackColor = Color.PaleGreen;
            ExitItem.BackColor = Color.DarkOrange;
            DeleteGroupItem.BackColor = Color.DarkOrange;
            OpenMenuAsLocalAdminItem.BackColor = Color.DarkOrange;
            UpdateItem.BackColor = Color.PaleGreen;

            if (currentDbEmpty)
            {
                logger.Warn("Form loading is finishing, but the local db is still empty!");
                VisibleOfAdminMenuItems(true);
            }
            else
            {
                VisibleOfAdminMenuItems(false);
                if (currentModeAppManual)
                {
                    nameOfLastTable = "ListFIO";
                    SeekAndShowMembersOfGroup("");
                    logger.Info("Программа запущена в интерактивном режиме...");
                }
                else
                {
                    // nameOfLastTable = "Mailing";
                    EnableControl(comboBoxFio, false);
                    logger.Info("Стартовый режим - автоматический...");

                    ShowDataTableDbQuery(dbApplication, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
                "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
                "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
                " ORDER BY RecipientEmail asc, DateCreated desc; ");
                    ExecuteAutoMode(true);
                }

                if (mailServer?.Length > 0 && mailServerSMTPPort > 0)
                {
                    _mailServer = new MailServer(mailServer, mailServerSMTPPort);
                }
                else
                {
                    logger.Warn("mailServer: " + mailServer + ", mailServerSMTPPort: " + mailServerSMTPPort);
                    MessageBox.Show("Проверьте параметры конфигурации почтового сервера в базе\n'mailServer', 'mailServerSMTPPort'.\nТекущие параметры:\n" +
                        "URL почтового сервера: " + mailServer +
                        "\nПорт сервера для отправки писем: " + mailServerSMTPPort,
                        "Функция рассылки отчетов не работает!",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                if (mailSenderAddress != null && mailSenderAddress.Contains('@'))
                {
                    _mailUser = new MailUser(NAME_OF_SENDER_REPORTS, mailSenderAddress);
                }
                else
                {
                    logger.Warn("mailSenderAddress: " + mailSenderAddress);
                    MessageBox.Show("Проверьте адрес отправителя почты в конфигурации почтового сервера в базе\n'mailSenderAddress'.\nТекущие параметры:\n" +
                        "адрес отправителя почты: " + mailSenderAddress,
                        "Функция рассылки отчетов не работает!",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                if (remoteFolderUpdateURL?.Length > 5)
                {
                    //Run Autoupdate function
                    Task.Run(() => CheckUpdates());
                }
                else
                {
                    logger.Warn("RemoteFolderUpdateURL: " + remoteFolderUpdateURL);
                    MessageBox.Show("Проверьте URL адрес сервера обновлений в базе\n'RemoteFolderUpdateURL'.\nТекущие параметры:\n" +
                        "URL адрес сервера обновлений: " + remoteFolderUpdateURL,
                        "Функция работа с обновлениями ПО не работает!",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                //loading parameters of configuration Application
                listParameters = GetConfigOfASTA();
                List<ParameterConfig> parameters = ReturnListParametersWithEmptyValue(listParameters);
                if (parameters?.Count > 0)
                {
                    string resultParameters = null;
                    foreach (var p in parameters)
                    {
                        resultParameters += (p.name + " is empty\n\r");
                    }

                    logger.Warn($"Empty parameters in local config db: {resultParameters}");
                }
            }

            if (ReturnComboBoxCountItems(comboBoxFio) > 0)
            {
                VisibleMenuItem(listFioItem, true);
                SetComboBoxIndex(comboBoxFio, 0);
            }
            comboBoxFio.DrawMode = DrawMode.OwnerDrawFixed;
            comboBoxFio.DrawItem += new DrawItemEventHandler(Set_ComboBox_DrawItem);

            //Naming of Menu Items
            SetMenuItemText(ModeItem, "Включить режим автоматических e-mail рассылок");
            SetMenuItemTooltip(ModeItem, "Включен интерактивный режим. Все рассылки остановлены.");
            SetMenuItemText(EditAnualDaysItem, Names.DAY_OFF_OR_WORK);
            SetMenuItemTooltip(EditAnualDaysItem, Names.DAY_OFF_OR_WORK_EDIT);
            SetMenuItemText(PersonOrGroupItem, Names.WORK_WITH_A_PERSON);
            SetMenuItemsTextAfterClosingDateTimePicker();

            //Naming of Control Items
            SetControlToolTip(textBoxGroup, "Создать или добавить в группу");
            SetControlToolTip(textBoxGroupDescription, "Изменить описание группы");

            logger.Info("");
            logger.Info("Загрузка и настройка интерфейса ПО завершена....");
            logger.Info("");
        }

        private void AboutSoft(object sender, EventArgs e) //Кнопка "О программе"
        { AboutSoft(); }

        private void AboutSoft()
        {
            this.Hide();
            AboutBox1 aboutBox = new AboutBox1();
            aboutBox.ShowDialog();
            buttonAboutForm = aboutBox.OKButtonClicked;
            this.Show();
            aboutBox?.Close();
            aboutBox?.Dispose();
        }

        private void ApplicationExit(object sender, EventArgs e)
        { ApplicationExit(); }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        { ApplicationExit(); }

        private void ApplicationExit()
        {
            logger.Info("");
            logger.Info("");
            logger.Info("-=-=  Завершение работы ПО  =-=-");
            logger.Info("-----------------------------------------");
            logger.Info("");

            dgvo = null;

            bmp?.Dispose();
            bmpLogo?.Dispose();

            dtPersonTemp?.Dispose();
            dtPersonTempAllColumns?.Dispose();
            dtPersonRegistrationsFullList?.Dispose();
            dtPeopleGroup?.Dispose();
            dtPeopleListLoaded?.Dispose();
            dtPeople?.Dispose();

            notifyIcon?.Dispose();
            contextMenu?.Dispose();
            mRightClick?.Dispose();

            //taskkill /F /IM ASTA.exe
            Text = @"Closing application...";
            System.Threading.Thread.Sleep(500);
            Application.Exit();
        }

        private List<string> ReadTXTFile(string fpath = null, int listMaxLength = 10000)
        {
            List<string> txt = new List<string>(listMaxLength);
            string query = string.Empty, s = string.Empty, result = string.Empty, log = string.Empty;

            Cursor = Cursors.WaitCursor;

            MethodInvoker mi = delegate
            {
                string prevText = ReturnStatusLabelText(StatusLabel2);
                Color prevColor = ReturnStatusLabelBackColor(StatusLabel2);
                bool readOk = true;

                if (!(fpath?.Length > 0))
                {
                    OpenFileDialogExtentions filePath = new OpenFileDialogExtentions("Выберите файл", "SQL файлы (*.sql)|*.sql|Все files (*.*)|*.*");
                     fpath = filePath.FilePath;

                    if (fpath == null) return;
                }

                logger.Trace("ReadTXTFile");
                SetStatusLabelText(StatusLabel2, $"Читаю файл: {fpath}");
                try
                {
                    using (System.IO.StreamReader Reader =
                            new System.IO.StreamReader(fpath, Encoding.GetEncoding(1251)))
                    {
                        while ((s = Reader.ReadLine()) != null)
                        {
                            if (s?.Trim()?.Length > 0)
                                txt.Add(s);
                        }//while
                    }//using
                }
                catch (Exception err)
                {
                    readOk = false;
                    logger.Warn($"Не могу прочитать файл: {fpath} \nException: {err.ToString()}");
                }

                if (!readOk)
                {
                    SetStatusLabelText(StatusLabel2, $"Не могу прочитать файл: {fpath}", true);
                }
                else
                {
                    SetStatusLabelBackColor(StatusLabel2, prevColor);
                    if (prevText?.Length > 0 && !nameof(StatusLabel2).Equals(prevText))
                    { SetStatusLabelText(StatusLabel2, prevText); }
                }
            };
            this.Invoke(mi);

            MethodInvoker mi2 = delegate
            {
                Cursor = Cursors.Default;
            };
            this.Invoke(mi2);

            return txt;
        }

        private void TryMakeLocalDB(string fpath = null)
        {
            List<string> txt = ReadTXTFile(fpath);
            string query = string.Empty;

            SetStatusLabelText(StatusLabel2, $"Создаю таблицы в БД на основе запроса из текстового файла: {fpath}");
            using (SqLiteDbWrapper dbWriter = new SqLiteDbWrapper(sqLiteLocalConnectionString, dbApplication))
            {
                dbWriter.Status += AddLoggerTraceText;

                dbWriter.Execute("begin");
                foreach (var s in txt)
                {
                    if (s.StartsWith("CREATE TABLE"))
                    { query = s.Trim(); }
                    else { query += s.Trim(); }

                    logger.Trace($"query: {query}");
                    if (s.EndsWith(";"))
                    {
                        dbWriter.Execute(query);
                    }
                }//foreach

                dbWriter.Execute("end");

                dbWriter.Status -= AddLoggerTraceText;

                SetStatusLabelText(StatusLabel2, "Таблицы в БД созданы.");
            }
        }

        private void SetTechInfoIntoDB() //Write Technical Info in DB
        {
            string query = "INSERT OR REPLACE INTO 'TechnicalInfo' (PCName, POName, POVersion, LastDateStarted, CurrentUser, FreeRam, GuidAppication) " +
                      " VALUES (@PCName, @POName, @POVersion, @LastDateStarted, @CurrentUser, @FreeRam, @GuidAppication)";

            using (SqLiteDbWrapper dbWriter = new SqLiteDbWrapper(sqLiteLocalConnectionString, dbApplication))
            {
                logger.Trace($"query: {query}");
                dbWriter.Status += AddLoggerTraceText;
                using (SQLiteCommand SqlQuery = new SQLiteCommand(query, dbWriter.sqlConnection))
                {
                    SqlQuery.Parameters.Add("@PCName", DbType.String).Value = Environment.MachineName + "|" + Environment.OSVersion;
                    SqlQuery.Parameters.Add("@POName", DbType.String).Value = Application.ExecutablePath + "(" + Application.ProductName + ")"; // appFileVersionInfo.FileName + "(+ appName + ")"
                    SqlQuery.Parameters.Add("@POVersion", DbType.String).Value = Application.ProductVersion;// appFileVersionInfo.FileVersion;
                    SqlQuery.Parameters.Add("@LastDateStarted", DbType.String).Value = DateTime.Now.ToYYYYMMDDHHMM();
                    SqlQuery.Parameters.Add("@CurrentUser", DbType.String).Value = Environment.UserName;
                    SqlQuery.Parameters.Add("@FreeRam", DbType.String).Value = "RAM: " + Environment.WorkingSet.ToString();
                    SqlQuery.Parameters.Add("@GuidAppication", DbType.String).Value = guid;

                    dbWriter.Execute(SqlQuery);
                }
                dbWriter.Status -= AddLoggerTraceText;
            }
        }

        private void LoadPreviouslySavedParameters()   //Select Previous Data from DB and write it into the combobox and Parameters
        {
            logger.Trace("-= LoadPreviouslySavedParameters =-");

            string modeApp = "";

            //Get data from registry
            using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(appRegistryKey, Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree, System.Security.AccessControl.RegistryRights.ReadKey))
            {
                try { sServer1Registry = EvUserKey?.GetValue("SKDServer")?.ToString(); }
                catch { logger.Trace("Can't get value of SCA server's name from Registry"); }
                try { sServer1UserNameRegistry = EncryptionDecryptionCriticalData.DecryptBase64ToString(EvUserKey?.GetValue("SKDUser")?.ToString(), keyEncryption, keyDencryption).ToString(); }
                catch { logger.Trace("Can't get value of SCA User from Registry"); }
                try { sServer1UserPasswordRegistry = EncryptionDecryptionCriticalData.DecryptBase64ToString(EvUserKey?.GetValue("SKDUserPassword")?.ToString(), keyEncryption, keyDencryption).ToString(); }
                catch { logger.Trace("Can't get value of SCA UserPassword from Registry"); }

                try { mysqlServerRegistry = EvUserKey?.GetValue("MySQLServer")?.ToString(); }
                catch { logger.Trace("Can't get value of MySQL Name from Registry"); }
                try { mysqlServerUserNameRegistry = EvUserKey?.GetValue("MySQLUser")?.ToString(); }
                catch { logger.Trace("Can't get value of MySQL User from Registry"); }
                try { mysqlServerUserPasswordRegistry = EncryptionDecryptionCriticalData.DecryptBase64ToString(EvUserKey?.GetValue("MySQLUserPassword")?.ToString(), keyEncryption, keyDencryption).ToString(); }
                catch { logger.Trace("Can't get value of MySQL UserPassword from Registry"); }

                try { modeApp = EvUserKey?.GetValue("ModeApp")?.ToString(); }
                catch { logger.Trace("Can't get value of ModeApp from Registry"); }
            }

            //Get data from local DB
            string query;
            int count = 0;
            using (SqLiteDbWrapper dbReader = new SqLiteDbWrapper(sqLiteLocalConnectionString, dbApplication))
            {
                System.Data.SQLite.SQLiteDataReader data = null;
                query = "SELECT ComboList FROM LastTakenPeopleComboList;";
                try
                {
                    data = dbReader?.GetData(query);
                }
                catch { logger.Info("LoadPreviouslySavedParameters: no any info in 'LastTakenPeopleComboList'"); }

                if (data != null)
                {
                    foreach (DbDataRecord record in data)
                    {
                        try
                        {
                            if (record["ComboList"]?.ToString()?.Trim()?.Length > 0)
                            {
                                AddComboBoxItem(comboBoxFio, record["ComboList"].ToString().Trim());
                                count++;
                            }
                        }
                        catch (Exception err) { logger.Info(err.ToString()); }
                    }
                }
                logger.Trace($"LoadPreviouslySavedParameters: query: {query}\n{count} rows loaded from 'LastTakenPeopleComboList'");

                query = "SELECT FIO FROM PeopleGroup;";
                count = 0;
                try
                {
                    data = dbReader.GetData(query);
                }
                catch { logger.Info("LoadPreviouslySavedParameters: no any info in 'PeopleGroup'"); }
                if (data != null)
                    foreach (DbDataRecord record in data)
                    {
                        if (record["FIO"]?.ToString()?.Length > 0)
                        {
                            count++;
                        }
                    }
                logger.Trace($"LoadPreviouslySavedParameters: query: {query}\n{count} rows loaded from 'PeopleGroup'");
            }

            //loading parameters
            listParameters = GetConfigOfASTA().FindAll(x => x?.isExample == "no"); //load only real data

            DEFAULT_DAY_OF_SENDING_REPORT = GetValueOfConfigParameter(listParameters, @"DEFAULT_DAY_OF_SENDING_REPORT", END_OF_MONTH);

            int.TryParse(GetValueOfConfigParameter(listParameters, @"ShiftDaysBackOfSendingFromLastWorkDay", ""), out int days);

            if (days < 0 || days > 27)
            { ShiftDaysBackOfSendingFromLastWorkDay = 3; }
            else
            { ShiftDaysBackOfSendingFromLastWorkDay = days; }

            clrRealRegistrationRegistry = Color.FromName(GetValueOfConfigParameter(listParameters, @"clrRealRegistration", "PaleGreen"));

            sServer1DB = GetValueOfConfigParameter(listParameters, @"SKDServer", null);
            sServer1UserNameDB = GetValueOfConfigParameter(listParameters, @"SKDUser", null);
            sServer1UserPasswordDB = GetValueOfConfigParameter(listParameters, @"SKDUserPassword", null);

            mysqlServerDB = GetValueOfConfigParameter(listParameters, @"MySQLServer", null);
            mysqlServerUserNameDB = GetValueOfConfigParameter(listParameters, @"MySQLUser", null);
            mysqlServerUserPasswordDB = GetValueOfConfigParameter(listParameters, @"MySQLUserPassword", null);

            mailServerDB = GetValueOfConfigParameter(listParameters, @"MailServer", null);
            mailServerSMTPPortDB = GetValueOfConfigParameter(listParameters, @"MailServerSMTPport", null);
            mailsOfSenderOfNameDB = GetValueOfConfigParameter(listParameters, @"MailUser", null);
            mailsOfSenderOfPasswordDB = GetValueOfConfigParameter(listParameters, @"MailUserPassword", null);

            mailJobReportsOfNameOfReceiver = GetValueOfConfigParameter(listParameters, @"JobReportsReceiver", null);
            string defaultURL = remoteFolderUpdateURL;
            remoteFolderUpdateURL = GetValueOfConfigParameter(listParameters, @"RemoteFolderUpdateURL", defaultURL);

            domain = GetValueOfConfigParameter(listParameters, @"DomainOfUser", null);

            //set app's variables
            {
                sServer1 = sServer1Registry?.Length > 0 ? sServer1Registry : sServer1DB;
                sServer1UserName = sServer1UserNameRegistry?.Length > 0 ? sServer1UserNameRegistry : sServer1UserNameDB;
                sServer1UserPassword = sServer1UserPasswordRegistry?.Length > 0 ? sServer1UserPasswordRegistry : sServer1UserPasswordDB;

                mailServer = mailServerDB?.Length > 0 ? mailServerDB : null;
                int.TryParse(mailServerSMTPPortDB, out mailServerSMTPPort);
                mailSenderAddress = mailsOfSenderOfNameDB?.Length > 0 ? mailsOfSenderOfNameDB : null;
                mailsOfSenderOfPassword = mailsOfSenderOfPasswordDB?.Length > 0 ? mailsOfSenderOfPasswordDB : null;

                mysqlServer = mysqlServerRegistry?.Length > 0 ? mysqlServerRegistry : mysqlServerDB;
                mysqlServerUserName = mysqlServerUserNameRegistry?.Length > 0 ? mysqlServerUserNameRegistry : mysqlServerUserNameDB;
                mysqlServerUserPassword = mysqlServerUserPasswordRegistry?.Length > 0 ? mysqlServerUserPasswordRegistry : mysqlServerUserPasswordDB;

                sqlServerConnectionString = $"Data Source={sServer1}\\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID={sServer1UserName};Password={sServer1UserPassword};Connect Timeout=30";
                mysqlServerConnectionString = $"server={mysqlServer};User={mysqlServerUserName};Password={mysqlServerUserPassword};database=wwwais;convert zero datetime=True;Connect Timeout=60";

                clrRealRegistration = clrRealRegistrationRegistry != Color.PaleGreen ? clrRealRegistrationRegistry : Color.PaleGreen;
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

            logger.Trace($"LoadPreviouslySavedParameters: {nameof(modeApp)} - {modeApp}, {nameof(currentModeAppManual)} - {currentModeAppManual}");
        }

        private string GetValueOfConfigParameter(List<ParameterConfig> listOfParameters, string nameParameter, string defaultValue)
        {
            return listOfParameters.FindLast(x => x?.name?.Trim() == nameParameter)?.value?.Trim() != null ?
                   listParameters.FindLast(x => x?.name?.Trim() == nameParameter)?.value?.Trim() :
                   defaultValue;
        }

        private List<ParameterConfig> ReturnListParametersWithEmptyValue(List<ParameterConfig> listOfParameters)
        {
            List<ParameterConfig> parameterConfigs = new List<ParameterConfig>();
            foreach (var parameter in listOfParameters)
            {
                if (string.IsNullOrWhiteSpace(parameter.value))
                {
                    parameterConfigs.Add(parameter);
                }
            }
            return parameterConfigs;
        }

        private void RefreshConfigInMainDBItem_Click(object sender, EventArgs e)
        { AddExceptedParametersIntoConfigurationDb(); }

        private void AddExceptedParametersIntoConfigurationDb()    //add not existed example of parameters into ConfigTable in the Main Local DB
        {
            SetStatusLabelText(StatusLabel2, "Проверяю список параметров конфигурации локальной БД...");

            ConfigurationOfASTA config = new ConfigurationOfASTA(dbApplication);
            config.status += AddLoggerTraceText;

            ParameterOfConfiguration parameterOfConfiguration;
            listParameters = config.GetParameters("%%");//.FindAll(x => x.isExample == "no");//update work parameters

            foreach (string sParameter in Names.allParametersOfConfig)
            {
                logger.Trace($"looking for: {sParameter} in local DB");
                if (!listParameters.Any(x => x?.name == sParameter))
                {
                    parameterOfConfiguration = new ParameterOfConfigurationBuilder().
                        SetName(sParameter).
                        SetValue("").
                        SetDescription("").
                        SetIsSecret(false).
                        SetIsExample("yes");

                    string resultSaving = config.SaveParameter(parameterOfConfiguration);
                    logger.Info($"Попытка добавить новый параметр в конфигурацию: {resultSaving}");
                }
            }
            config.status -= AddLoggerTraceText;
            config = null;

            SetStatusLabelText(StatusLabel2, "Обновление параметров конфигурации локальной БД завершено");
        }

        private void AddParameterInConfigItem_Click(object sender, EventArgs e)
        {
            AddParameterInConfig();
        }

        private void AddParameterInConfig()
        {
            EnableMainMenuItems(false);

            logger.Trace($"-= {nameof(AddParameterInConfigItem_Click)} =-");

            VisibleControl(panelView, false);
            ClearButtonClickEvent(btnPropertiesSave);
            SetControlText(btnPropertiesSave, "Сохранить параметр");
            btnPropertiesSave.Click += new EventHandler(ButtonPropertiesSave_inConfig);

            listParameters = GetConfigOfASTA();

            foreach (string sParameter in Names.allParametersOfConfig)
            {
                if (!(listParameters.FindLast(x => x?.name?.Trim() == sParameter)?.value?.Length > 0))
                {
                    listParameters.Add(new ParameterConfig()
                    {
                        name = sParameter,
                        description = "Example",
                        value = "",
                        isSecret = false,
                        isExample = "yes"
                    });
                }
            }

            // listParameters = parameter.GetParameters("%%").FindAll(x => x.isExample != "no"); //load only real data
            InitializeParameterFormSettings(listParameters);
        }

        private List<ParameterConfig> GetConfigOfASTA(string parameterName = "%%")
        {
            List<ParameterConfig> list = new List<ParameterConfig>();
            ConfigurationOfASTA config = new ConfigurationOfASTA(dbApplication);
            config.status += AddLoggerTraceText;

            list = config.GetParameters(parameterName);

            config.status -= AddLoggerTraceText;
            config = null;
            return list;
        }

        private void InitializeParameterFormSettings(List<ParameterConfig> listParameters)
        {
            logger.Trace($"-= {nameof(InitializeParameterFormSettings)} =-");

            panelViewResize(numberPeopleInLoading);

            listboxPeriod = new ListBox
            {
                Location = new Point(20, 120),
                Size = new Size(590, 100),
                Parent = groupBoxProperties,
                DrawMode = DrawMode.OwnerDrawFixed,
                Sorted = true
            };

            listboxPeriod.DrawItem += new DrawItemEventHandler(Set_ListBox_DrawItem);
            listboxPeriod.DataSource = listParameters.Select(x => x.name).ToList();
            if (listParameters.Count > 0) listboxPeriod.SelectedIndex = 0;
            toolTip1.SetToolTip(listboxPeriod, "Перечень параметров");

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
            listboxPeriod.SelectedIndexChanged += new EventHandler(ListBox_SelectedIndexChanged);
            textBoxSettings16.KeyPress += new KeyPressEventHandler(textboxDate_KeyPress);

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
            string result = ReturnListBoxSelectedItem(sender as ListBox);
            string tooltip = "";

            checkBox1.Checked = false;
            labelServer1.Text = "";
            labelSettings9.Text = "";
            textBoxSettings16.Text = "";
            toolTip1.SetToolTip(textBoxSettings16, tooltip);

            checkBox1.Checked = listParameters.FindLast(x => x.name == result).isSecret;
            labelServer1.Text = result;
            labelSettings9.Text = listParameters.FindLast(x => x.name == result)?.description;
            textBoxSettings16.Text = listParameters.FindLast(x => x.name == result)?.value;
            tooltip = listParameters.FindLast(x => x.name == result)?.description;
            toolTip1.SetToolTip(textBoxSettings16, tooltip);
        }

        private void textboxDate_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)//если нажата Enter
            {
              string  inputed = ReturnStrongNameDayOfSendingReports((sender as TextBox).Text);
                (sender as TextBox).Text = inputed;
            }
        }

        private string SaveParameterInConfigASTA(ParameterConfig parameter)
        {
            ConfigurationOfASTA config = new ConfigurationOfASTA(dbApplication);
            config.status += AddLoggerTraceText;

            ParameterOfConfiguration parameterOfConfiguration = new ParameterOfConfigurationBuilder()
                .SetParameter(parameter);

            string result = config.SaveParameter(parameterOfConfiguration);
            config.status -= AddLoggerTraceText;
            config = null;

            return result;
        }

        private void ButtonPropertiesSave_inConfig(object sender, EventArgs e) //SaveProperties()
        {
            logger.Trace($"-= {nameof(ButtonPropertiesSave_inConfig)} =-");

            string textLabel = ReturnTextOfControl(labelSettings9);

            ParameterConfig parameter = new ParameterConfig()
            {
                description = textLabel?.ToLower() == "example" ? "" : textLabel,
                name = ReturnTextOfControl(labelServer1),
                value = ReturnTextOfControl(textBoxSettings16),
                isSecret = ReturnCheckboxCheckedStatus(checkBox1),
                isExample = "no"
            };

            string resultSaving = SaveParameterInConfigASTA(parameter);
            MessageBox.Show(parameter.name + " обновлен новым значением - " + parameter.value + "\n" + resultSaving);
            parameter = null;

            DisposeTemporaryControls();
            VisibleControl(panelView, true);

            if (mailServer?.Length > 0 && mailServerSMTPPort > 0)
            {
                _mailServer = new MailServer(mailServer, mailServerSMTPPort);
            }
            if (mailSenderAddress != null && mailSenderAddress.Contains('@'))
            {
                _mailUser = new MailUser(NAME_OF_SENDER_REPORTS, mailSenderAddress);
            }

            ShowDataTableDbQuery(dbApplication, "ConfigDB", "SELECT ParameterName AS 'Имя параметра', " +
            "Value AS 'Данные', Description AS 'Описание', DateCreated AS 'Дата создания/модификации'",
            " ORDER BY ParameterName asc, DateCreated desc; ");

            EnableMainMenuItems(true);
        }

        //void ShowDataTableDbQuery(
        private void ShowDataTableDbQuery(System.IO.FileInfo dbApplication, string myTable, string mySqlQuery, string mySqlWhere) //Query data from the Table of the DB
        {
            logger.Trace($"-= {nameof(ShowDataTableDbQuery)} =-");

            string query = $"{mySqlQuery} FROM '{myTable}' {mySqlWhere};";
            using (SqLiteDbWrapper dbReader = new SqLiteDbWrapper(sqLiteLocalConnectionString, dbApplication))
            {
                using (DataTable dt = dbReader.GetDataTable(query))
                {
                    iCounterLine = dt.Rows.Count;

                    dgvo.Show(dataGridView1, false);
                    dgvo.AddDataTable(dataGridView1, dt);
                    dgvo.Show(dataGridView1, true);
                }
            }
            nameOfLastTable = myTable;
            logger.Trace($"ShowDataTableDbQuery: {iCounterLine}");
            sLastSelectedElement = "dataGridView";
        }
      
        private void ShowDatatableOnDatagridview(DataTable dt, string nameLastTable, string[] nameHidenColumnsArray1 = null) //Query data from the Table of the DB
        {
            logger.Trace($"-= {nameof(ShowDatatableOnDatagridview)} =-");

            using (DataTable dataTable = dt.Copy())
            {
                for (int i = 0; i < nameHidenColumnsArray1?.Length; i++)
                {
                    if (nameHidenColumnsArray1[i]?.Length > 0)
                        try { dataTable.Columns[nameHidenColumnsArray1[i]].ColumnMapping = MappingType.Hidden; } catch { }
                }

                SetStatusLabelText(StatusLabel2, $"Всего записей: {dataTable.Rows.Count}");

                dgvo.Show(dataGridView1, false);
                dgvo.AddDataTable(dataGridView1, dataTable);
                dgvo.Show(dataGridView1, true);
            }
            nameOfLastTable = nameLastTable;
            sLastSelectedElement = "dataGridView";
        }

        private async Task ExecuteQueryOnLocalDB(System.IO.FileInfo dbApplication, string query) //Delete All data from the selected Table of the DB (both parameters are string)
        {
            if (dbApplication.Exists)
            {
                using (SqLiteDbWrapper dbWriter = new SqLiteDbWrapper(sqLiteLocalConnectionString, dbApplication))
                {
                    logger.Trace($"query: {query}");
                    dbWriter.Status += AddLoggerTraceText;

                    using (SQLiteCommand SqlQuery = new SQLiteCommand(query, dbWriter.sqlConnection))
                    {
                        await Task.Run(() =>
                        {
                            dbWriter.Execute(SqlQuery);
                        });
                    }
                    dbWriter.Status -= AddLoggerTraceText;
                }
            }
        }

        private async Task DeleteDataTableQueryParameters(System.IO.FileInfo dbApplication, string myTable, string sqlParameter1, string sqlData1,
            string sqlParameter2 = "", string sqlData2 = "", string sqlParameter3 = "", string sqlData3 = "",
            string sqlParameter4 = "", string sqlData4 = "", string sqlParameter5 = "", string sqlData5 = "", string sqlParameter6 = "", string sqlData6 = "") //Delete data from the Table of the DB by NAV (both parameters are string)
        {
            string result = string.Empty;
            string query = $"DELETE FROM '{myTable}' Where {sqlParameter1}= @{sqlParameter1} " +
                           $"AND {sqlParameter2}= @{sqlParameter2} AND {sqlParameter3}= @{sqlParameter3} " +
                           $"AND {sqlParameter4}= @{sqlParameter4} AND {sqlParameter5}= @{sqlParameter5} AND {sqlParameter6}= @{sqlParameter6};";

            await Task.Run(() =>
            {
                using (SqLiteDbWrapper dbWriter = new SqLiteDbWrapper(sqLiteLocalConnectionString, dbApplication))
                {
                    dbWriter.Status += AddLoggerTraceText;

                    SQLiteCommand sqlCommand = null;
                    if (sqlParameter1.Length > 0 && sqlParameter2.Length > 0 && sqlParameter3.Length > 0 && sqlParameter4.Length > 0
                    && sqlParameter5.Length > 0 && sqlParameter6.Length > 0)
                    {
                        sqlCommand = new SQLiteCommand(query, dbWriter.sqlConnection);

                        sqlCommand.Parameters.Add("@" + sqlParameter1, DbType.String).Value = sqlData1;
                        sqlCommand.Parameters.Add("@" + sqlParameter2, DbType.String).Value = sqlData2;
                        sqlCommand.Parameters.Add("@" + sqlParameter3, DbType.String).Value = sqlData3;
                        sqlCommand.Parameters.Add("@" + sqlParameter4, DbType.String).Value = sqlData4;
                        sqlCommand.Parameters.Add("@" + sqlParameter5, DbType.String).Value = sqlData5;
                        sqlCommand.Parameters.Add("@" + sqlParameter6, DbType.String).Value = sqlData6;
                    }
                    else if (sqlParameter1.Length > 0 && sqlParameter2.Length > 0 && sqlParameter3.Length > 0 && sqlParameter4.Length > 0
                        && sqlParameter5.Length > 0)
                    {
                        query = $"DELETE FROM '{myTable}' Where {sqlParameter1}= @{sqlParameter1} " +
                                $"AND {sqlParameter2}= @{sqlParameter2} AND {sqlParameter3}= @{sqlParameter3} " +
                                $"AND {sqlParameter4}= @{sqlParameter4} AND {sqlParameter5}= @{sqlParameter5} ;";

                        sqlCommand = new SQLiteCommand(query, dbWriter.sqlConnection);
                        sqlCommand.Parameters.Add("@" + sqlParameter1, DbType.String).Value = sqlData1;
                        sqlCommand.Parameters.Add("@" + sqlParameter2, DbType.String).Value = sqlData2;
                        sqlCommand.Parameters.Add("@" + sqlParameter3, DbType.String).Value = sqlData3;
                        sqlCommand.Parameters.Add("@" + sqlParameter4, DbType.String).Value = sqlData4;
                        sqlCommand.Parameters.Add("@" + sqlParameter5, DbType.String).Value = sqlData5;
                    }
                    else if (sqlParameter1.Length > 0 && sqlParameter2.Length > 0 && sqlParameter3.Length > 0 && sqlParameter4.Length > 0)
                    {
                        sqlCommand = new SQLiteCommand(query, dbWriter.sqlConnection);
                        sqlCommand.Parameters.Add("@" + sqlParameter1, DbType.String).Value = sqlData1;
                        sqlCommand.Parameters.Add("@" + sqlParameter2, DbType.String).Value = sqlData2;
                        sqlCommand.Parameters.Add("@" + sqlParameter3, DbType.String).Value = sqlData3;
                        sqlCommand.Parameters.Add("@" + sqlParameter4, DbType.String).Value = sqlData4;
                    }
                    else if (sqlParameter1.Length > 0 && sqlParameter2.Length > 0 && sqlParameter3.Length > 0)
                    {
                        query = $"DELETE FROM '{myTable}' Where {sqlParameter1}= @{sqlParameter1} " +
                                $"AND {sqlParameter2}= @{sqlParameter2} AND {sqlParameter3}= @{sqlParameter3};";

                        sqlCommand = new SQLiteCommand(query, dbWriter.sqlConnection);
                        sqlCommand.Parameters.Add("@" + sqlParameter1, DbType.String).Value = sqlData1;
                        sqlCommand.Parameters.Add("@" + sqlParameter2, DbType.String).Value = sqlData2;
                        sqlCommand.Parameters.Add("@" + sqlParameter3, DbType.String).Value = sqlData3;
                    }
                    else if (sqlParameter1.Length > 0 && sqlParameter2.Length > 0)
                    {
                        query = $"DELETE FROM '{myTable}' Where {sqlParameter1}= @{sqlParameter1} AND {sqlParameter2}= @{sqlParameter2};";

                        sqlCommand = new SQLiteCommand(query, dbWriter.sqlConnection);
                        sqlCommand.Parameters.Add("@" + sqlParameter1, DbType.String).Value = sqlData1;
                        sqlCommand.Parameters.Add("@" + sqlParameter2, DbType.String).Value = sqlData2;
                    }
                    else if (sqlParameter1.Length > 0)
                    {
                        query = $"DELETE FROM '{myTable}' Where {sqlParameter1}= @{sqlParameter1};";

                        sqlCommand = new SQLiteCommand(query, dbWriter.sqlConnection);
                        sqlCommand.Parameters.Add("@" + sqlParameter1, DbType.String).Value = sqlData1;
                    }
                    dbWriter.Execute(sqlCommand);

                    dbWriter.Status -= AddLoggerTraceText;
                    sqlCommand?.Dispose();
                }
            });
            logger.Trace($"Удалены данные из таблицы {myTable} - query: {query}");
        }

        private void CheckAliveIntellectServer(string serverName, string connectDB, string queryCheckDb) //Check alive the SKD Intellect-server and its DB's 'intellect'
        {
            //stop checking last registrations
            checkInputsOutputs = false;

            stimerCurr = $"Проверка доступности {serverName}. Ждите окончания процесса...";
            bServer1Exist = false;

            AddLoggerTraceText($"CheckAliveIntellectServer: query: {queryCheckDb}");

            try
            {
                using (SqlDbReader sqlDbTableReader = new SqlDbReader(connectDB))
                {
                    sqlDbTableReader.Status += AddLoggerTraceText;
                    sqlDbTableReader.GetData(queryCheckDb);
                    bServer1Exist = true;
                    AddLoggerTraceText($"Сервер {serverName} с SQL БД СКД-сервера доступен");
                }
            }
            catch (Exception err)
            {
                AddLoggerTraceText($"Ошибка доступа к {serverName} SQL БД СКД-сервера! {err.ToString()}");
            }
            stimerCurr = "Ждите!";
        }

        private async void GetADUsersItem_Click(object sender, EventArgs e)
        { await Task.Run(() => GetUsersFromAD()); }

        private IADUser ReturnADUserFromLocalDb()
        {
            IADUser user = new ADUser();
            listParameters = GetConfigOfASTA().FindAll(x => x.isExample == "no"); //load only real data

            user.Login = listParameters.FindLast(x => x?.name == @"UserName")?.value;
            user.Password = listParameters.FindLast(x => x?.name == @"UserPassword")?.value;
            user.DomainControllerPath = listParameters.FindLast(x => x?.name == @"DomainController")?.value;
            user.Domain = listParameters.FindLast(x => x?.name == @"DomainOfUser")?.value;

            return user;
        }

        private void GetUsersFromAD()
        {
            logger.Trace($"-= {nameof(GetUsersFromAD)} =-");
            string prevStatusLabel1Text = ReturnStatusLabelText(StatusLabel1);

            stimerCurr = null;
            listUsersAD = new List<ADUserFullAccount>();
            IADUser ad = ReturnADUserFromLocalDb();

            logger.Trace($"user, domain, password, server: {ad.Login} |{ad.Domain} |{ad.Password} |{ad.DomainControllerPath}");

            if (ad.Login?.Length > 0 && ad.Password?.Length > 0 && ad.Domain?.Length > 0 && ad.DomainControllerPath?.Length > 0)
            {
                SetStatusLabelText(StatusLabel1, $"Получаю данные из домена: '{ad.Domain}'");

                ADUsers users = new ADUsers(ad);
                users.Info += SetStatusLabelText;
                users.Trace += AddLoggerTraceText;
                //    ad.ADUsersCollection.CollectionChanged += Users_CollectionChanged;
                listUsersAD = users.GetADUsersFromDomain();

                users.Info -= SetStatusLabelText;
                users.Trace -= AddLoggerTraceText; ;
                users = null;

                listUsersAD.Sort();

                logger.Trace($"GetUsersFromAD: count users in list: {listUsersAD?.Count}");
                AddLoggerTraceText($"Из домена {domain} получено {listUsersAD?.Count} учетных записей");

                //передать дальше в обработку:
                foreach (var person in listUsersAD)
                {
                    AddLoggerTraceText($"{person?.fio} |{person?.Login} |{person?.Code} |{person?.mail} |{person?.lastLogon}");
                }
            }
            else
            {
                logger.Error($"Ошибка доступа к домену. user: {ad.Login} |domain: {ad.Domain} |password: {ad.Password} |server: {ad.DomainControllerPath}");

                SetStatusLabelText(
                    StatusLabel2,
                    $"Ошибка доступа к домену {domain}",
                    true);
            }
            stimerCurr = "Ждите!";

            SetStatusLabelText(StatusLabel1, prevStatusLabel1Text);
        }

        //  уведомление о количестве и последнем полученном из AD пользователей
        //private void Users_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        //{
        //    switch (e.Action)
        //    {
        //        case NotifyCollectionChangedAction.Add: // если добавление
        //            UserAD newUser = e.NewItems[0] as UserAD;
        //            stimerPrev = "Получено из AD: " + newUser.id + " пользователей, последний: " + newUser.fio.ConvertFullNameToShortForm();
        //            break;
        //        /*   case NotifyCollectionChangedAction.Remove: // если удаление
        //               UserAD oldUser = e.OldItems[0] as UserAD;
        //               break;
        //           case NotifyCollectionChangedAction.Replace: // если замена
        //               UserAD replacedUser = e.OldItems[0] as UserAD;
        //               UserAD replacingUser = e.NewItems[0] as UserAD;
        //               break;*/
        //        default:
        //            break;
        //    }
        //}

        private async void GetFio_Click(object sender, EventArgs e)  //DoListsFioGroupsMailings()
        {
            ProgressBar1Start();
            currentAction = "GetFIO";
            CheckBoxesFiltersAll_SetState(false);
            CheckBoxesFiltersAll_Enable(false);
            EnableMenuItem(LoadDataItem, false);
            EnableMenuItem(FunctionMenuItem, false);
            EnableMenuItem(SettingsMenuItem, false);
            EnableMenuItem(GroupsMenuItem, false);
            EnableMenuItem(GetFioItem, false);
            EnableControl(dataGridView1, false);

            await Task.Run(() => CheckAliveIntellectServer(sServer1, sqlServerConnectionString, queryCheckSystemDdSCA));

            if (bServer1Exist)
            {
                await Task.Run(() => DoListsFioGroupsMailings());

                VisibleMenuItem(listFioItem, true);
                EnableMenuItem(SettingsMenuItem, true);
                EnableMenuItem(GetFioItem, true);
                EnableMenuItem(FunctionMenuItem, true);
                EnableMenuItem(LoadDataItem, true);
                EnableMenuItem(GroupsMenuItem, true);

                EnableControl(dataGridView1, true);
                VisibleControl(dataGridView1, true);
                VisibleControl(pictureBox1, false);
                EnableControl(comboBoxFio, true);

                SetStatusLabelForeColor(StatusLabel1, Color.Black);
                SetStatusLabelForeColor(StatusLabel2, Color.Black);
            }
            else
            {
                EnableMenuItem(SettingsMenuItem, true);
            }
            ProgressBar1Stop();
        }

        private void DoListsFioGroupsMailings()  //  GetDataFromRemoteServers()  ImportTablePeopleToTableGroupsInLocalDB()
        {
            logger.Trace($"-= {nameof(DoListsFioGroupsMailings)} =-");

            countGroups = 0;
            countUsers = 0;
            countMailers = 0;

            if (currentAction != "sendEmail")
            { SetStatusLabelText(StatusLabel2, "Получаю данные с серверов..."); }

            //Get list people FIO and mail from AD
            GetUsersFromAD();

            using (DataTable dtTempIntermediate = dtPeople.Clone())
            {
                //Get list FIO + chief from www and SCA
                GetDataFromRemoteServers(dtTempIntermediate, peopleShifts);

                if (currentAction != "sendEmail")
                { SetStatusLabelText(StatusLabel2, "Формирую и записываю группы в локальную базу..."); }
                WriteGroupsMailingsInLocalDb(dtTempIntermediate, peopleShifts);

                if (currentAction != "sendEmail")
                { SetStatusLabelText(StatusLabel2, "Записываю ФИО в локальную базу..."); }

                WritePeopleInLocalDB(dbApplication.ToString(), dtTempIntermediate);

                if (currentAction != "sendEmail")
                {
                    dtPersonTemp = LeaveAndOrderColumnsOfDataTable(dtTempIntermediate, Names.orderColumnsFIO);

                    ShowDatatableOnDatagridview(dtPersonTemp, "ListFIO");

                    SetStatusLabelText(StatusLabel2, $"Записано в локальную базу: {countUsers} ФИО, {countGroups} групп и {countMailers} рассылок");
                }
            }
        }

        //Get the list of registered users
        private void GetDataFromRemoteServers(DataTable dataTablePeople, List<PeopleShift> peopleShifts)
        {
            logger.Trace($"-= {nameof(GetDataFromRemoteServers)} =-");

            EmployeeFull personFromServer = new EmployeeFull();
            DataRow row;

            string query;
            string fio = "";
            string nav = "";
            string groupName = "";
            string idGroup = "";
            string depDescription = "";
            string depBossCode = "";

            string timeStart = ConvertDecimalTimeToStringHHMM(ReturnNumUpDown(numUpDownHourStart), ReturnNumUpDown(numUpDownMinuteStart));
            string timeEnd = ConvertDecimalTimeToStringHHMM(ReturnNumUpDown(numUpDownHourEnd), ReturnNumUpDown(numUpDownMinuteEnd));
            string dayStartShift = "";
            string dayStartShift_ = "";

            listFIO = new List<Employee>();
            Dictionary<string, Department> departments = new Dictionary<string, Department>();
            Department departmentFromDictionary;

            ClearComboBox(comboBoxFio);
            SetStatusLabelText(StatusLabel2, $"Запрашиваю данные с {sServer1}. Ждите окончания процесса...");

            try
            {
                string confitionToLoad = "";

                using (var sqlConnection = new SQLiteConnection(sqLiteLocalConnectionString))
                {
                    sqlConnection.Open();

                    AddLoggerTraceText("Получаю из локальной базы список городов для построения условий загрузки ФИО из MySQL(www) базы...");
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

                // import users and group from SCA server
                using (SqlDbReader sqlDbTableReader = new SqlDbReader(sqlServerConnectionString))
                {
                    query = "SELECT id,level_id,name,owner_id,parent_id,region_id,schedule_id FROM OBJ_DEPARTMENT";
                    AddLoggerTraceText($"from SCA server, query: {query}");
                    sqlDbTableReader.Status += AddLoggerTraceText;

                    //import group from MSSQL SCA server
                    using (System.Data.SqlClient.SqlDataReader sqlData = sqlDbTableReader.GetData(query))
                    {
                        foreach (DbDataRecord record in sqlData)
                        {
                            idGroup = record["id"]?.ToString()?.Trim();
                            groupName = record?["name"]?.ToString()?.Trim();
                            if (groupName?.Length > 0 && idGroup?.Length > 0 && !departments.ContainsKey(idGroup))
                            {
                                departments.Add(idGroup, new Department()
                                {
                                    DepartmentId = idGroup,
                                    DepartmentDescr = groupName,
                                    DepartmentBossCode = sServer1
                                });
                            }
                            ProgressWork1Step();
                        }
                    }
                    sqlDbTableReader.Status -= AddLoggerTraceText;
                }

                //import users from с SCA server
                using (SqlDbReader sqlDbTableReader = new SqlDbReader(sqlServerConnectionString))
                {
                    query = "SELECT id, name, surname, patronymic, post, tabnum, parent_id, facility_code, card FROM OBJ_PERSON WHERE is_locked = '0' AND facility_code NOT LIKE '' AND card NOT LIKE '' ";
                    AddLoggerTraceText($"query: {query}");
                    sqlDbTableReader.Status += AddLoggerTraceText;

                    using (System.Data.SqlClient.SqlDataReader sqlData = sqlDbTableReader.GetData(query))
                    {
                        foreach (DbDataRecord record in sqlData)
                        {
                            if (record["name"]?.ToString()?.Trim()?.Length > 0)
                            {
                                row = dataTablePeople.NewRow();
                                fio = (record["name"]?.ToString()?.Trim() + " " + record["surname"]?.ToString()?.Trim() + " " + record["patronymic"]?.ToString()?.Trim()).Replace(@"  ", @" ");
                                groupName = record["parent_id"]?.ToString()?.Trim();
                                nav = record["tabnum"]?.ToString()?.Trim()?.ToUpper();

                                if (departments.TryGetValue(groupName, out departmentFromDictionary))
                                {
                                    depDescription = departmentFromDictionary.DepartmentDescr;
                                }
                                else { depDescription = ""; }

                                row[Names.N_ID] = Convert.ToInt32(record["id"]?.ToString()?.Trim());
                                row[Names.FIO] = fio;
                                row[Names.CODE] = nav;

                                row[Names.GROUP] = groupName;
                                row[Names.DEPARTMENT] = depDescription;
                                row[Names.DEPARTMENT_ID] = sServer1.IndexOf('.') > -1 ? sServer1.Remove(sServer1.IndexOf('.')) : sServer1;

                                row[Names.EMPLOYEE_POSITION] = record["post"]?.ToString()?.Trim();

                                row[Names.DESIRED_TIME_IN] = timeStart;
                                row[Names.DESIRED_TIME_OUT] = timeEnd;

                                dataTablePeople.Rows.Add(row);

                                listFIO.Add(new Employee { fio = fio, Code = nav });

                                ProgressWork1Step();
                            }
                        }
                    }
                    sqlDbTableReader.Status -= AddLoggerTraceText;
                }

                // import users, shifts and group from web DB
                int tmpSeconds = 0;
                groupName = mysqlServer;
                SetStatusLabelText(StatusLabel2, $"Запрашиваю данные с {mysqlServer}. Ждите окончания процесса...");

                // import departments from web DB
                using (MySqlDbReader mysqlDbTableReader = new MySqlDbReader(mysqlServerConnectionString))
                {
                    query = "SELECT id, parent_id, name, boss_code FROM dep_struct ORDER by id";

                    AddLoggerTraceText($"query: {query}");
                    mysqlDbTableReader.Status += AddLoggerTraceText;

                    using (MySql.Data.MySqlClient.MySqlDataReader mysqlData = mysqlDbTableReader.GetData(query))
                    {
                        while (mysqlData.Read())
                        {
                            idGroup = mysqlData?.GetString(@"id");

                            if (mysqlData?.GetString(@"name")?.Length > 0 && idGroup?.Length > 0 && !departments.ContainsKey(idGroup))
                            {
                                departments.Add(idGroup, new Department()
                                {
                                    DepartmentId = idGroup,
                                    DepartmentDescr = mysqlData.GetString(@"name"),
                                    DepartmentBossCode = mysqlData?.GetString(@"boss_code")
                                });
                            }
                            ProgressWork1Step();
                        }
                    }
                    mysqlDbTableReader.Status -= AddLoggerTraceText;
                }

                // import individual shifts of people from web DB
                using (MySqlDbReader mysqlDbTableReader = new MySqlDbReader(mysqlServerConnectionString))
                {
                    query = "Select code,start_date,mo_start,mo_end,tu_start,tu_end,we_start,we_end,th_start,th_end,fr_start,fr_end, " +
                                    "sa_start,sa_end,su_start,su_end,comment FROM work_time ORDER by start_date";
                    AddLoggerTraceText($"query: {query}");
                    mysqlDbTableReader.Status += AddLoggerTraceText;

                    using (MySql.Data.MySqlClient.MySqlDataReader mysqlData = mysqlDbTableReader.GetData(query))
                    {
                        while (mysqlData.Read())
                        {
                            if (mysqlData.GetString(@"code")?.Length > 0)
                            {
                                try { dayStartShift = DateTime.Parse(mysqlData.GetMySqlDateTime(@"start_date").ToString()).ToYYYYMMDD(); }
                                catch
                                { dayStartShift = DateTime.Parse("1980-01-01").ToYYYYMMDD(); }

                                peopleShifts.Add(new PeopleShift()
                                {
                                    code = mysqlData.GetString(@"code").Replace('C', 'S'),
                                    dayStartShift = dayStartShift,
                                    MoStart = Convert.ToInt32(mysqlData.GetString(@"mo_start")),
                                    MoEnd = Convert.ToInt32(mysqlData.GetString(@"mo_end")),
                                    TuStart = Convert.ToInt32(mysqlData.GetString(@"tu_start")),
                                    TuEnd = Convert.ToInt32(mysqlData.GetString(@"tu_end")),
                                    WeStart = Convert.ToInt32(mysqlData.GetString(@"we_start")),
                                    WeEnd = Convert.ToInt32(mysqlData.GetString(@"we_end")),
                                    ThStart = Convert.ToInt32(mysqlData.GetString(@"th_start")),
                                    ThEnd = Convert.ToInt32(mysqlData.GetString(@"th_end")),
                                    FrStart = Convert.ToInt32(mysqlData.GetString(@"fr_start")),
                                    FrEnd = Convert.ToInt32(mysqlData.GetString(@"fr_end")),
                                    SaStart = Convert.ToInt32(mysqlData.GetString(@"sa_start")),
                                    SaEnd = Convert.ToInt32(mysqlData.GetString(@"sa_end")),
                                    SuStart = Convert.ToInt32(mysqlData.GetString(@"su_start")),
                                    SuEnd = Convert.ToInt32(mysqlData.GetString(@"su_end")),
                                    Status = "",
                                    Comment = mysqlData.GetString(@"comment")
                                });
                                ProgressWork1Step();
                            }
                            try
                            {
                                dayStartShift = peopleShifts.FindLast((x) => x.code == "0").dayStartShift;
                                dayStartShift_ = "Общий график с " + dayStartShift;

                                tmpSeconds = peopleShifts.FindLast((x) => x.code == "0" && x.dayStartShift == dayStartShift).MoStart;
                                timeStart = (tmpSeconds.ConvertSecondsIntoStringsHHmmArray())[2];

                                tmpSeconds = peopleShifts.FindLast((x) => x.code == "0" && x.dayStartShift == dayStartShift).MoEnd;
                                timeEnd = (tmpSeconds.ConvertSecondsIntoStringsHHmmArray())[2];

                                logger.Trace("Общий график с " + dayStartShift);
                            }
                            catch {/*nothing todo on any errors */ }
                        }
                    }
                    mysqlDbTableReader.Status -= AddLoggerTraceText;
                }

                // import people from web DB
                using (MySqlDbReader mysqlDbTableReader = new MySqlDbReader(mysqlServerConnectionString))
                {
                    query = "Select code, family_name, first_name, last_name, vacancy, department, boss_id, city FROM personal " + confitionToLoad;//where hidden=0

                    AddLoggerTraceText($"query: {query}");
                    mysqlDbTableReader.Status += AddLoggerTraceText;

                    using (MySql.Data.MySqlClient.MySqlDataReader mysqlData = mysqlDbTableReader.GetData(query))
                    {
                        while (mysqlData.Read())
                        {
                            if (mysqlData.GetString(@"family_name")?.Trim()?.Length > 0)
                            {
                                row = dataTablePeople.NewRow();
                                personFromServer = new EmployeeFull();

                                fio = (mysqlData.GetString(@"family_name")?.Trim() + " " + mysqlData.GetString(@"first_name")?.Trim() + " " + mysqlData.GetString(@"last_name")?.Trim())?.Replace(@"  ", @" ");

                                personFromServer.fio = fio.Replace("&acute;", "'");
                                personFromServer.Code = mysqlData.GetString(@"code")?.Trim()?.ToUpper()?.Replace('C', 'S');
                                personFromServer.DepartmentId = mysqlData.GetString(@"department")?.Trim();
                                personFromServer.City = mysqlData.GetString(@"city")?.Trim();
                                personFromServer.PositionInDepartment = mysqlData.GetString(@"vacancy")?.Trim();

                                if (departments.TryGetValue(personFromServer?.DepartmentId, out departmentFromDictionary))
                                {
                                    depDescription = departmentFromDictionary?.DepartmentDescr;
                                    depBossCode = departmentFromDictionary?.DepartmentBossCode;
                                }

                                personFromServer.Department = depDescription ?? personFromServer?.DepartmentId;

                                personFromServer.DepartmentBossCode = depBossCode ?? mysqlData.GetString(@"boss_id")?.Trim();

                                personFromServer.GroupPerson = groupName;
                                personFromServer.Shift = dayStartShift_;
                                personFromServer.ControlInHHMM = timeStart;
                                personFromServer.ControlOutHHMM = timeEnd;

                                dayStartShift = peopleShifts.FindLast((x) => x.code == personFromServer.Code).dayStartShift;
                                if (dayStartShift?.Length > 0)
                                {
                                    personFromServer.Shift = $"Индивидуальный график с {dayStartShift}";

                                    tmpSeconds = peopleShifts.FindLast((x) => x.code == personFromServer.Code).MoStart;
                                    personFromServer.ControlInHHMM = tmpSeconds.ConvertSecondsIntoStringsHHmmArray()[2];

                                    tmpSeconds = peopleShifts.FindLast((x) => x.code == personFromServer.Code).MoEnd;
                                    personFromServer.ControlOutHHMM = tmpSeconds.ConvertSecondsIntoStringsHHmmArray()[2];

                                    personFromServer.Comment = peopleShifts.FindLast((x) => x.code == personFromServer.Code).Comment;

                                    AddLoggerTraceText($"Индивидуальный график с {dayStartShift} {personFromServer.Code} {personFromServer.Comment}");
                                }

                                row[Names.FIO] = personFromServer.fio;
                                row[Names.CODE] = personFromServer.Code;

                                row[Names.GROUP] = personFromServer.GroupPerson;
                                row[Names.PLACE_EMPLOYEE] = personFromServer.City;

                                row[Names.DEPARTMENT] = personFromServer.Department;
                                row[Names.DEPARTMENT_ID] = personFromServer.DepartmentId;
                                row[Names.EMPLOYEE_POSITION] = personFromServer.PositionInDepartment;
                                row[Names.CHIEF_ID] = personFromServer.DepartmentBossCode;

                                row[Names.EMPLOYEE_SHIFT] = personFromServer.Shift;

                                row[Names.DESIRED_TIME_IN] = personFromServer.ControlInHHMM;
                                row[Names.DESIRED_TIME_OUT] = personFromServer.ControlOutHHMM;

                                dataTablePeople.Rows.Add(row);

                                listFIO.Add(new Employee { fio = personFromServer.fio, Code = personFromServer.Code });

                                ProgressWork1Step();
                            }
                        }
                    }
                    mysqlDbTableReader.Status -= AddLoggerTraceText;
                }

                dataTablePeople.AcceptChanges();

                AddLoggerTraceText($"departments.count: {departments.Count}");
                SetStatusLabelText(StatusLabel2, "ФИО и наименования департаментов получены.");
            }
            catch (Exception err)
            {
                AddLoggerTraceText($"departments.count: {err.ToString()}");

                SetStatusLabelText(
                    StatusLabel2,
                    "Возникла ошибка во время получения данных с серверов.",
                    true);
            }

            query = fio = nav = groupName = depDescription = depBossCode = timeStart = timeEnd = dayStartShift = dayStartShift_ = null;
            row = null; departments = null; personFromServer = null;
        }

        private void WriteGroupsMailingsInLocalDb(DataTable dataTablePeople, List<PeopleShift> peopleShifts)
        {
            logger.Trace($"-= {nameof(WriteGroupsMailingsInLocalDb)} =-");

            SetStatusLabelText(StatusLabel2, "Формирую обновленные списки ФИО, департаментов и рассылок...");

            logger.Info("Приступаю к формированию списков ФИО и департаментов...");
            string query;
            string depId = "";
            string depName = "";
            string depBoss = "";
            string depDescr = "";
            string depBossEmail = "";

            // HashSet<DepartmentFull> groups = new HashSet<DepartmentFull>();
            Dictionary<string, DepartmentFull> groups = new Dictionary<string, DepartmentFull>();
            HashSet<Department> departmentsUniq = new HashSet<Department>();
            HashSet<DepartmentFull> departmentsEmailUniq = new HashSet<DepartmentFull>();
            ProgressWork1Step();

            string skdName = sServer1.Split('.')[0];
            int iSKD = 0;
            int iMysql = 0;
            foreach (var dr in dataTablePeople.AsEnumerable())
            {
                depId = dr[Names.DEPARTMENT_ID]?.ToString();
                depBossEmail = listUsersAD.Find((x) => x.Code == dr[Names.CHIEF_ID]?.ToString())?.mail;
                if (depId?.Length > 0 && !groups.ContainsKey("@" + depId))
                {
                    if (depId == skdName && iSKD < 1)
                    {
                        groups.Add("@" + depId, new DepartmentFull()
                        {
                            DepartmentId = "@" + depId,
                            DepartmentDescr = "skd",
                            DepartmentBossCode = dr[Names.CHIEF_ID]?.ToString(),
                            DepartmentBossEmail = depBossEmail
                        });
                        iSKD++;
                    }
                    else if (depId != skdName)
                    {
                        groups.Add("@" + depId, new DepartmentFull()
                        {
                            DepartmentId = "@" + depId,
                            DepartmentDescr = dr[Names.DEPARTMENT]?.ToString(),
                            DepartmentBossCode = dr[Names.CHIEF_ID]?.ToString(),
                            DepartmentBossEmail = depBossEmail
                        });
                    }
                }

                depId = dr[Names.GROUP]?.ToString();
                if (depId?.Length > 0 && !groups.ContainsKey(depId))
                {
                    if (depId == mysqlServer && iMysql < 1)
                    {
                        groups.Add(depId, new DepartmentFull()
                        {
                            DepartmentId = depId,
                            DepartmentDescr = "web",
                            DepartmentBossCode = "GetCodeFromDB",
                            DepartmentBossEmail = "GetEmailFromDB"
                        });
                        iMysql++;
                    }
                    else if (depId != mysqlServer)
                    {
                        groups.Add(depId, new DepartmentFull()
                        {
                            DepartmentId = depId,
                            DepartmentDescr = dr[Names.DEPARTMENT]?.ToString(),
                            DepartmentBossCode = "GetCodeFromDB",
                            DepartmentBossEmail = "GetEmailFromDB"
                        });
                    }
                }
                ProgressWork1Step();
            }
            foreach (var dep in groups)
            {
                logger.Trace($"{dep.Value.DepartmentId} {dep.Value.DepartmentDescr} {dep.Value.DepartmentBossCode} {dep.Value.DepartmentBossEmail}");
            }
            logger.Trace($"groups.count: {groups.Distinct().Count()}");

            foreach (var strDepartment in groups)
            {
                if (strDepartment.Value?.DepartmentId?.Length > 0)
                {
                    departmentsUniq.Add(new Department
                    {
                        DepartmentId = strDepartment.Value.DepartmentId,
                        DepartmentDescr = strDepartment.Value.DepartmentDescr,
                        DepartmentBossCode = strDepartment.Value.DepartmentBossCode
                    });

                    if (strDepartment.Value?.DepartmentBossEmail?.Length > 0)
                    {
                        departmentsEmailUniq.Add(new DepartmentFull
                        {
                            DepartmentId = strDepartment.Value.DepartmentId,
                            DepartmentDescr = strDepartment.Value.DepartmentDescr,
                            DepartmentBossEmail = strDepartment.Value.DepartmentBossEmail
                        });
                    }
                }
            }
            ProgressWork1Step();

            logger.Trace("Чищу базу от старых списков с ФИО...");

            ExecuteQueryOnLocalDB(dbApplication, "DELETE FROM 'LastTakenPeopleComboList';").GetAwaiter().GetResult();

            foreach (var department in departmentsUniq?.ToList()?.Distinct())
            {
                DeleteDataTableQueryParameters(dbApplication, "PeopleGroup", "GroupPerson", department?.DepartmentId).GetAwaiter().GetResult();
                ProgressWork1Step();
            }
            ProgressWork1Step();

            using (var sqlConnection = new SQLiteConnection(sqLiteLocalConnectionString))
            {
                sqlConnection.Open();

                using (var sqlCommand1 = new SQLiteCommand("begin", sqlConnection))
                { sqlCommand1.ExecuteNonQuery(); }

                logger.Info("Готовлю списки исключений из рассылок...");
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
                                departmentsEmailUniq.RemoveWhere(x => x.DepartmentBossEmail == dbRecordTemp);
                            }
                        }
                    }
                }
                using (var sqlCommand1 = new SQLiteCommand("end", sqlConnection))
                { sqlCommand1.ExecuteNonQuery(); }

                logger.Info("Записываю новые группы ...");
                using (var sqlCommand1 = new SQLiteCommand("begin", sqlConnection))
                { sqlCommand1.ExecuteNonQuery(); }

                foreach (var deprtment in departmentsUniq.ToList().Distinct())
                {
                    depName = deprtment.DepartmentId;
                    depDescr = deprtment.DepartmentDescr;

                    depBoss = deprtment.DepartmentBossCode?.Length > 0 ? deprtment.DepartmentBossCode : "Default_Recepient_code_From_Db";
                    using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PeopleGroupDescription' (GroupPerson, GroupPersonDescription, Recipient) " +
                                           "VALUES (@GroupPerson, @GroupPersonDescription, @Recipient)", sqlConnection))
                    {
                        command.Parameters.Add("@GroupPerson", DbType.String).Value = depName;
                        command.Parameters.Add("@GroupPersonDescription", DbType.String).Value = depDescr;
                        command.Parameters.Add("@Recipient", DbType.String).Value = depBoss;
                        try { command.ExecuteNonQuery(); } catch { }
                    }

                    logger.Trace("CreatedGroup: " + depName + "(" + depDescr + ")");
                    ProgressWork1Step();
                }
                using (var sqlCommand1 = new SQLiteCommand("end", sqlConnection))
                { sqlCommand1.ExecuteNonQuery(); }

                logger.Info("Записываю новые рассылки по группам с учетом исключений...");
                string recipientEmail = "";
                using (var sqlCommand1 = new SQLiteCommand("begin", sqlConnection))
                { sqlCommand1.ExecuteNonQuery(); }
                foreach (var deprtment in departmentsEmailUniq?.ToList()?.Distinct())
                {
                    depName = deprtment?.DepartmentId;
                    depDescr = deprtment?.DepartmentDescr;
                    depBoss = deprtment?.DepartmentBossCode;
                    recipientEmail = deprtment?.DepartmentBossEmail;

                    if (depName.StartsWith("@") && recipientEmail.Contains('@'))
                    {
                        using (SQLiteCommand sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'Mailing' (RecipientEmail, GroupsReport, NameReport, Description, Period, Status, DateCreated, SendingLastDate, TypeReport, DayReport)" +
                           " VALUES (@RecipientEmail, @GroupsReport, @NameReport, @Description, @Period, @Status, @DateCreated, @SendingLastDate, @TypeReport, @DayReport)", sqlConnection))
                        {
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

                        logger.Trace($"SaveMailing: {recipientEmail} {depName} {depDescr}");
                    }
                    ProgressWork1Step();
                }
                using (var sqlCommand1 = new SQLiteCommand("end", sqlConnection))
                { sqlCommand1.ExecuteNonQuery(); }

                logger.Info("Записываю новые индивидуальные графики...");
                using (var sqlCommand1 = new SQLiteCommand("begin", sqlConnection))
                { sqlCommand1.ExecuteNonQuery(); }
                foreach (var shift in peopleShifts?.ToArray())
                {
                    if (shift.code?.Length > 0)
                    {
                        using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'ListOfWorkTimeShifts' (NAV, DayStartShift, MoStart, MoEnd, " +
                            "TuStart, TuEnd, WeStart, WeEnd, ThStart, ThEnd, FrStart, FrEnd, SaStart, SaEnd, SuStart, SuEnd, Status, Comment, DayInputed) " +
                        " VALUES (@NAV, @DayStartShift, @MoStart, @MoEnd, @TuStart, @TuEnd, @WeStart, @WeEnd, @ThStart, @ThEnd, @FrStart, @FrEnd, " +
                        " @SaStart, @SaEnd, @SuStart, @SuEnd, @Status, @Comment, @DayInputed)", sqlConnection))
                        {
                            sqlCommand.Parameters.Add("@NAV", DbType.String).Value = shift.code;
                            sqlCommand.Parameters.Add("@DayStartShift", DbType.String).Value = shift.dayStartShift;
                            sqlCommand.Parameters.Add("@MoStart", DbType.Int32).Value = shift.MoStart;
                            sqlCommand.Parameters.Add("@MoEnd", DbType.Int32).Value = shift.MoEnd;
                            sqlCommand.Parameters.Add("@TuStart", DbType.Int32).Value = shift.TuStart;
                            sqlCommand.Parameters.Add("@TuEnd", DbType.Int32).Value = shift.TuEnd;
                            sqlCommand.Parameters.Add("@WeStart", DbType.Int32).Value = shift.WeStart;
                            sqlCommand.Parameters.Add("@WeEnd", DbType.Int32).Value = shift.WeEnd;
                            sqlCommand.Parameters.Add("@ThStart", DbType.Int32).Value = shift.ThStart;
                            sqlCommand.Parameters.Add("@ThEnd", DbType.Int32).Value = shift.ThEnd;
                            sqlCommand.Parameters.Add("@FrStart", DbType.Int32).Value = shift.FrStart;
                            sqlCommand.Parameters.Add("@FrEnd", DbType.Int32).Value = shift.FrEnd;
                            sqlCommand.Parameters.Add("@SaStart", DbType.Int32).Value = shift.SaStart;
                            sqlCommand.Parameters.Add("@SaEnd", DbType.Int32).Value = shift.SaEnd;
                            sqlCommand.Parameters.Add("@SuStart", DbType.Int32).Value = shift.SuStart;
                            sqlCommand.Parameters.Add("@SuEnd", DbType.Int32).Value = shift.SuEnd;
                            sqlCommand.Parameters.Add("@Status", DbType.String).Value = shift.Status;
                            sqlCommand.Parameters.Add("@Comment", DbType.String).Value = shift.Comment;
                            sqlCommand.Parameters.Add("@DayInputed", DbType.String).Value = DateTime.Now.ToYYYYMMDDHHMM();
                            try { sqlCommand.ExecuteNonQuery(); } catch (Exception err) { MessageBox.Show(err.ToString()); }
                            ProgressWork1Step();
                        }
                    }
                }
                using (var sqlCommand1 = new SQLiteCommand("end", sqlConnection))
                { sqlCommand1.ExecuteNonQuery(); }

                int.TryParse(departmentsUniq?.ToArray()?.Distinct()?.Count().ToString(), out int countGroups);
                int.TryParse(departmentsEmailUniq?.ToArray()?.Distinct()?.Count().ToString(), out int countMailers);

                logger.Info("Записано групп: " + countGroups);
                logger.Info("Записано рассылок: " + countMailers);
            }

            departmentsUniq = null;
            departmentsEmailUniq = null;

            ProgressWork1Step();
            SetStatusLabelText(StatusLabel2, "Списки ФИО и департаментов получены.");
        }

        private void ShowListFioItem_Click(object sender, EventArgs e) //ListFioReturn()
        {
            nameOfLastTable = "ListFIO";

            EnableControl(comboBoxFio, true);
            SeekAndShowMembersOfGroup("");
        }

        private async void Export_Click(object sender, EventArgs e)
        {
            logger.Trace($"-= {nameof(Export_Click)} =-");

            ProgressBar1Start();
            EnableMenuItem(FunctionMenuItem, false);
            EnableMenuItem(SettingsMenuItem, false);
            EnableMenuItem(GroupsMenuItem, false);
            EnableControl(dataGridView1, false);

            filePathExcelReport = System.IO.Path.Combine(localAppFolderPath, $"InputOutputs {DateTime.Now.ToString("yyyyMMdd_HHmmss")}");

            await Task.Run(() => ExportDatatableSelectedColumnsToExcel(dtPersonTemp, "InputOutputsOfStaff", filePathExcelReport));
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("explorer.exe", " /select, " + filePathExcelReport + @".xlsx")); // //System.Reflection.Assembly.GetExecutingAssembly().Location)

            EnableMenuItem(FunctionMenuItem, true);
            EnableMenuItem(SettingsMenuItem, true);
            EnableMenuItem(GroupsMenuItem, true);
            EnableControl(dataGridView1, true);
            ProgressBar1Stop();
        }

        private void ExportDatatableSelectedColumnsToExcel(DataTable dataTable, string nameReport, string filePath)  //Export DataTable to Excel
        {
            logger.Trace($"-= {nameof(ExportDatatableSelectedColumnsToExcel)} =-");
            logger.Trace($"{nameof(nameReport)}:{nameReport} |{nameof(filePath)}:{filePath} ");

            string pathToFile = BuilderFileName.BuildPath(filePath, "xlsx");

            reportExcelReady = false;

            dataTable.SetColumnsOrder(Names.orderColumnsFinacialReport);
            DataTable dtExport;
            string sortOrderInTheFirst = Names.DEPARTMENT + ", " + Names.FIO + ", " + Names.DATE_REGISTRATION + " ASC";
            string sortOrderInTheSecond = Names.GROUP + ", " + Names.FIO + ", " + Names.DATE_REGISTRATION + " ASC";

            // Sort order of collumns
            using (DataView dv = dataTable.DefaultView)
            {
                try
                {
                    dv.Sort = sortOrderInTheFirst;
                    dtExport = dv.ToTable();
                    logger.Trace($"Сортировка: {sortOrderInTheFirst}");
                }
                catch
                {
                    dv.Sort = sortOrderInTheSecond;
                    dtExport = dv.ToTable();
                    logger.Trace($"Сортировка: {sortOrderInTheSecond}");
                }
            }

            logger.Trace($"В таблице {dataTable.TableName} столбцов - {dtExport.Columns.Count}, строк - {dtExport.Rows.Count}");
            SetStatusLabelText(StatusLabel2, $"Генерирую Excel-файл по отчету: '{nameReport}'");
            ProgressWork1Step();

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
                ProgressWork1Step();

                int rows = 1;
                int rowsInTable = dtExport.Rows.Count;
                int columnsInTable = indexColumns.Length - 1; // p.Length;
                sheet.Name = nameReport;
                //sheet.Names.Add("next", "=" + Path.GetFileNameWithoutExtension(filePathExcelReport) + "!$A$1", true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                ProgressWork1Step();

                //colourize background of column
                //the last column
                sheet.Columns[columnsInTable].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSilver;  // последняя колонка

                //other ways to colourize background:
                //sheet.Columns[columnsInTable].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Silver);
                // or
                //sheet.Columns[columnsInTable].Interior.Color = System.Drawing.Color.Silver;

                sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(Names.FIO)) + 1)]
               .Interior.Color = Color.DarkSeaGreen;

                try
                {
                    sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(Names.EMPLOYEE_BEING_LATE)) + 1)]
                    .Interior.Color = Color.SandyBrown;
                    sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(Names.EMPLOYEE_BEING_LATE)) + 1)]
                   .Font.Color = Color.Red;
                    //.Font.Color = ColorTranslator.ToOle(Color.Red);
                    //arranges text in the center of the column
                    Microsoft.Office.Interop.Excel.Range rangeColumnA = sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(Names.EMPLOYEE_BEING_LATE)) + 1)];
                    rangeColumnA.Cells.EntireColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(Names.EMPLOYEE_EARLY_DEPARTURE)) + 1)]
                    .Interior.Color = Color.SandyBrown;
                    sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(Names.EMPLOYEE_EARLY_DEPARTURE)) + 1)]
                   .Font.Color = Color.Red;
                    //arranges text in the center of the column
                    Microsoft.Office.Interop.Excel.Range rangeColumnB = sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(Names.EMPLOYEE_EARLY_DEPARTURE)) + 1)];
                    rangeColumnB.Cells.EntireColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                }
                catch (Exception err) { logger.Warn($"Нарушения: {err.ToString()}"); }
                ProgressWork1Step();

                try
                {
                    Microsoft.Office.Interop.Excel.Range rangeColumnC = sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(@"Отпуск")) + 1)];
                    rangeColumnC.Cells.EntireColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                }
                catch (Exception err) { logger.Warn($"Отпуск: {err.ToString()}"); }

                try
                {
                    Microsoft.Office.Interop.Excel.Range rangeColumnD = sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(Names.EMPLOYEE_HOOKY)) + 1)];
                    rangeColumnD.Cells.EntireColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                }
                catch (Exception err) { logger.Warn($"Отгул: {err.ToString()}"); }

                try
                {
                    Microsoft.Office.Interop.Excel.Range rangeColumnE = sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(Names.EMPLOYEE_SICK_LEAVE)) + 1)];
                    rangeColumnE.Cells.EntireColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                }
                catch (Exception err) { logger.Warn($"Больничный: {err.ToString()}"); }

                try
                {
                    Microsoft.Office.Interop.Excel.Range rangeColumnF = sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(Names.EMPLOYEE_ABSENCE)) + 1)];
                    rangeColumnF.Cells.EntireColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                }
                catch (Exception err) { logger.Warn($"Отсутствовал: {err.ToString()}"); }
                ProgressWork1Step();

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
                ProgressWork1Step();

                foreach (DataRow row in dtExport.Rows)
                {
                    rows++;
                    for (int column = 0; column < columnsInTable; column++)
                    {
                        sheet.Cells[rows, column + 1].Value = row[indexColumns[column]];
                    }
                    ProgressWork1Step();
                }

                //colourize parts of text in the selected cell by different colors
                /*
                Microsoft.Office.Interop.Excel.Range rng = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[1, (i + 1)];
                rng.Value = "Green Red";
                rng.Characters[1, 5].Font.Color = Color.Green;
                rng.Characters[7, 9].Font.Color = Color.Red;
                */

                Microsoft.Office.Interop.Excel.Range range = sheet.UsedRange;  //get the using range
                                                                               //sheet.Cells.Range["A1", GetExcelColumnName(columnsInTable) + rowsInTable];
                                                                               // Microsoft.Office.Interop.Excel.Range range = sheet.Range["A2", GetExcelColumnName(columnsInTable) + (rows - 1)];

                range.Cells.Font.Name = "Tahoma";           //Шрифт для диапазона
                range.Cells.Font.Size = 8;                  //Размер шрифта
                range.Cells.EntireColumn.AutoFit();         //ширина колонок - авто
                ProgressWork1Step();

                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                ProgressWork1Step();

                //Autofilter
                range.Select();
                range.AutoFilter(1, Type.Missing, Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);

                //save document
                workbook.SaveAs(pathToFile,
                    Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                    false, false,
                    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, true,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value);
                ProgressWork1Step();

                SetStatusLabelText(StatusLabel2, $"Отчет сохранен в файл: {pathToFile}.xlsx");

                filePath = pathToFile;
                SetStatusLabelForeColor(StatusLabel2, Color.Black);
                reportExcelReady = true;
                ReleaseObject(range);
                ReleaseObject(rangeColumnName);
            }
            catch (Exception err)
            {
                SetStatusLabelText(
                    StatusLabel2,
                    "Ошибка генерации файла. Проверьте наличие установленного Excel",
                    true);
                AddLoggerTraceText(err.ToString());
            }
            finally
            {
                //close document
                workbook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                workbooks.Close();

                //clear temporary objects
                ReleaseObject(sheet);
                ReleaseObject(workbook);
                ReleaseObject(workbooks);
                excel.Quit();
                ReleaseObject(excel);

                dtExport?.Dispose();
            }
            ProgressWork1Step();

            sLastSelectedElement = "ExportExcel";
        }

        private void ReleaseObject(object obj) //for function - ExportToExcel()
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                logger.Trace($"Exception releasing object of {nameof(obj)}: {ex.ToString()}");
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private string GetExcelColumnName(int number)
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
        { SetTextBoxFIOAndTextBoxNAVFromSelectedComboboxFio(); }

        private void SetTextBoxFIOAndTextBoxNAVFromSelectedComboboxFio()
        {
            logger.Trace($"-= {nameof(SetTextBoxFIOAndTextBoxNAVFromSelectedComboboxFio)} =-");

            textBoxNav.Text = "";
            CheckBoxesFiltersAll_Enable(false);
            string[] sComboboxFIO = comboBoxFio?.SelectedItem?.ToString()?.Trim()?.Split('|');

            if (sComboboxFIO?.Length > 0)
            {
                textBoxFIO.Text = sComboboxFIO[0]?.Trim();
                StatusLabel2.Text = @"Выбран: " + textBoxFIO?.Text?.ConvertFullNameToShortForm();

                if (sComboboxFIO?.Length > 1)
                {
                    textBoxNav.Text = sComboboxFIO[1]?.Trim();
                }
            }
            else { textBoxFIO.Text = ""; }

            if (comboBoxFio.SelectedIndex > -1)
            {
                LoadDataItem.BackColor = Color.PaleGreen;
                groupBoxPeriod.BackColor = Color.PaleGreen;
                groupBoxTimeStart.BackColor = Color.PaleGreen;
                groupBoxTimeEnd.BackColor = Color.PaleGreen;
                groupBoxFilterReport.BackColor = SystemColors.Control;
            }

            if (nameOfLastTable == "LastIputsOutputs")
            {
                _selectedEmployeeVisitor = new EmployeeVisitor
                {
                    fio = ReturnTextOfControl(textBoxFIO),
                    code = ReturnTextOfControl(textBoxNav)
                };

                logger.Trace($"{_selectedEmployeeVisitor.fio} { _selectedEmployeeVisitor.code}");
                AddVisitorsAtDataGridView(visitors, _selectedEmployeeVisitor);
            }
        }

        private void CreateGroupItem_Click(object sender, EventArgs e)
        {
            logger.Trace($"-= {nameof(CreateGroupItem_Click)} =-");

            string group = ReturnTextOfControl(textBoxGroup);
            string groupDescr = ReturnTextOfControl(textBoxGroupDescription);

            if (group?.Length > 0)
            {
                CreateGroupInDB(group, groupDescr);
            }

            PersonOrGroupItem.Text = Names.WORK_WITH_A_PERSON;
            nameOfLastTable = "PeopleGroup";
            ListGroups();
        }

        private void CreateGroupInDB(string nameGroup, string descriptionGroup)
        {
            logger.Trace($"-= {nameof(CreateGroupInDB)} =-");

            if (nameGroup?.Length > 0)
            {
                using (var connection = new SQLiteConnection(sqLiteLocalConnectionString))
                {
                    connection.Open();
                    using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PeopleGroupDescription' (GroupPerson, GroupPersonDescription) " +
                                            "VALUES (@GroupPerson, @GroupPersonDescription )", connection))
                    {
                        command.Parameters.Add("@GroupPerson", DbType.String).Value = nameGroup;
                        command.Parameters.Add("@GroupPersonDescription", DbType.String).Value = descriptionGroup;
                        try { command.ExecuteNonQuery(); } catch { }
                    }
                }
                SetStatusLabelText(StatusLabel2, $"Группа '{nameGroup}' создана");
            }
        }

        private void ListGroupsItem_Click(object sender, EventArgs e)
        {
            ListGroups();
        }

        private void ListGroups()
        {
            logger.Trace($"-= {nameof(ListGroups)} =-");

            VisibleControl(groupBoxProperties, false);
            VisibleControl(dataGridView1, false);

            UpdateAmountAndRecepientOfPeopleGroupDescription();

            ShowDataTableDbQuery(dbApplication, "PeopleGroupDescription",
                "SELECT GroupPerson AS 'Группа', GroupPersonDescription AS 'Описание группы', " +
                "AmountStaffInDepartment AS 'Колличество сотрудников в группе', " +
                "Recipient AS '" + Names.RECEPIENTS_OF_REPORTS + "' ",
                " group by GroupPerson ORDER BY GroupPerson asc; ");

            LoadDataItem.BackColor = Color.PaleGreen;
            groupBoxPeriod.BackColor = Color.PaleGreen;
            groupBoxTimeStart.BackColor = Color.PaleGreen;
            groupBoxTimeEnd.BackColor = Color.PaleGreen;
            groupBoxFilterReport.BackColor = SystemColors.Control;

            DeleteGroupItem.Visible = true;
            MembersGroupItem.Enabled = true;
            PersonOrGroupItem.Text = Names.WORK_WITH_A_PERSON;

            EnableControl(comboBoxFio, false);

            VisibleControl(dataGridView1, true);
            dataGridView1.Select();
        }

        private void UpdateAmountAndRecepientOfPeopleGroupDescription()
        {
            logger.Trace($"-= {nameof(UpdateAmountAndRecepientOfPeopleGroupDescription)} =-");

            logger.Trace("UpdateAmountAndRecepientOfPeopleGroupDescription");
            List<string> groupsUncount = new List<string>();
            List<GroupParameters> amounts = new List<GroupParameters>();
            HashSet<string> groupsUniq = new HashSet<string>();
            List<DepartmentFull> emails = new List<DepartmentFull>();
            string tmpRec = "";
            string query = "";

            using (var sqlConnection = new SQLiteConnection(sqLiteLocalConnectionString))
            {
                sqlConnection.Open();

                using (SQLiteCommand sqlCommand = new SQLiteCommand("begin", sqlConnection))
                { sqlCommand.ExecuteNonQuery(); }

                //set to empty for amounts and recepients in the PeopleGroupDescription
                query = "UPDATE 'PeopleGroupDescription' SET AmountStaffInDepartment='0';";
                using (var command = new SQLiteCommand(query, sqlConnection))
                { command.ExecuteNonQuery(); }

                query = "UPDATE 'PeopleGroupDescription' SET Recipient='';";
                using (var command = new SQLiteCommand(query, sqlConnection))
                { command.ExecuteNonQuery(); }
                logger.Trace(query);

                using (var sqlCommand = new SQLiteCommand("end", sqlConnection))
                { sqlCommand.ExecuteNonQuery(); }

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
                            if (record != null)
                            {
                                group = record["GroupPerson"]?.ToString();
                                idGroup = "@" + record["DepartmentId"]?.ToString();

                                if (group?.Length > 0)
                                {
                                    groupsUncount.Add(group);
                                }
                                if (idGroup?.Length > 1)
                                {
                                    groupsUncount.Add(idGroup);
                                }
                            }
                        }
                    }
                }

                query = "SELECT GroupPerson FROM PeopleGroupDescription;";
                logger.Trace(query);
                using (var sqlCommand = new SQLiteCommand(query, sqlConnection))
                {
                    using (var reader = sqlCommand.ExecuteReader())
                    {
                        foreach (DbDataRecord record in reader)
                        {
                            if (record != null)
                            {
                                tmpRec = record["GroupPerson"]?.ToString();
                                if (tmpRec?.Length > 0)
                                { groupsUniq.Add(tmpRec); }
                            }
                        }
                    }
                }

                if (groupsUniq?.Count > 0)
                {
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
                                    if (record != null)
                                    {
                                        if (tmpRec.Length == 0)
                                        { tmpRec += record["RecipientEmail"]?.ToString(); }
                                        else { tmpRec += ", " + record["RecipientEmail"]?.ToString(); }
                                    }
                                }
                            }
                        }
                        if (grp?.Length > 0)
                        {
                            emails.Add(new DepartmentFull
                            {
                                DepartmentId = grp,
                                DepartmentBossEmail = tmpRec
                            });
                        }
                    }
                }
                logger.Trace("groupsUncount: " + (new HashSet<string>(groupsUncount)).Count());
                if (groupsUncount?.Count > 0)
                {
                    foreach (string group in new HashSet<string>(groupsUncount))
                    {
                        if (group?.Length > 0)
                        {
                            amounts.Add(new GroupParameters()
                            {
                                GroupName = group,
                                AmountMembers = groupsUncount.Where(x => x == group).Count(),
                                Emails = emails.Find(x => x.DepartmentId == group)?.DepartmentBossEmail
                            });
                        }
                    }
                }

                using (var sqlCommand = new SQLiteCommand("begin", sqlConnection))
                { sqlCommand.ExecuteNonQuery(); }
                if (amounts?.Count > 0)
                {
                    foreach (var group in amounts.ToArray())
                    {
                        if (group != null)
                        {
                            query = "UPDATE 'PeopleGroupDescription' SET AmountStaffInDepartment='" + group?.AmountMembers + "' WHERE GroupPerson like '" + group?.GroupName + "';";
                            logger.Trace(query);
                            using (var command = new SQLiteCommand(query, sqlConnection))
                            { command.ExecuteNonQuery(); }

                            query = "UPDATE 'PeopleGroupDescription' SET Recipient='" + group.Emails + "' WHERE GroupPerson like '" + group.GroupName + "';";
                            logger.Trace(query);
                            using (var command = new SQLiteCommand(query, sqlConnection))
                            { command.ExecuteNonQuery(); }
                        }
                    }
                }

                using (var sqlCommand = new SQLiteCommand("end", sqlConnection))
                { sqlCommand.ExecuteNonQuery(); }
            }
        }

        private void MembersGroupItem_Click(object sender, EventArgs e)//SearchMembersSelectedGroup()
        { SearchMembersSelectedGroup(); }

        private void SearchMembersSelectedGroup()
        {
            if (nameOfLastTable == "PeopleGroup" || nameOfLastTable == "PeopleGroupDescription")
            {
                string[] cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                    Names.GROUP
                });

                SeekAndShowMembersOfGroup(cellValue[0]);
            }
            else if (nameOfLastTable == "Mailing")
            {
                string[] cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                    @"Отчет по группам"
                });
                SeekAndShowMembersOfGroup(cellValue[0]);
            }
        }

        private void SeekAndShowMembersOfGroup(string nameGroup)
        {
            logger.Trace($"-= {nameof(SeekAndShowMembersOfGroup)} =-");

            using (var dtTemp = dtPeople.Clone())
            {
                numberPeopleInLoading = 0;
                DataRow dataRow;
                string dprtmnt = "", query = ""; ;

                query = "Select FIO, NAV, GroupPerson, ControllingHHMM, ControllingOUTHHMM, Shift, Comment, Department, PositionInDepartment, DepartmentId, City, Boss FROM PeopleGroup ";
                if (!string.IsNullOrEmpty(nameGroup) && nameGroup.Contains("@"))
                { query += " where DepartmentId like '" + nameGroup.Remove(0, 1) + "'"; }
                else if (!string.IsNullOrEmpty(nameGroup))
                { query += " where GroupPerson like '" + nameGroup + "'"; }

                query += ";";
                logger.Trace("SeekAndShowMembersOfGroup: " + query);

                using (SqLiteDbWrapper dbReader = new SqLiteDbWrapper(sqLiteLocalConnectionString, dbApplication))
                {
                    try
                    {
                        using (System.Data.SQLite.SQLiteDataReader data = dbReader?.GetData(query))
                        {
                            foreach (DbDataRecord record in data)
                            {
                                if (record != null && record["FIO"]?.ToString()?.Length > 0 && record["NAV"]?.ToString()?.Length > 0)
                                {
                                    dprtmnt = record[@"Department"]?.ToString() ?? record[@"GroupPerson"]?.ToString();

                                    dataRow = dtTemp.NewRow();
                                    dataRow[Names.FIO] = record[@"FIO"].ToString();
                                    dataRow[Names.CODE] = record[@"NAV"].ToString();

                                    dataRow[Names.GROUP] = record[@"GroupPerson"]?.ToString();
                                    dataRow[Names.DEPARTMENT] = dprtmnt;
                                    dataRow[Names.DEPARTMENT_ID] = record[@"DepartmentId"]?.ToString();
                                    dataRow[Names.EMPLOYEE_POSITION] = record[@"PositionInDepartment"]?.ToString();
                                    dataRow[Names.PLACE_EMPLOYEE] = record[@"City"]?.ToString();
                                    dataRow[Names.CHIEF_ID] = record[@"Boss"]?.ToString();

                                    dataRow[Names.DESIRED_TIME_IN] = record[@"ControllingHHMM"]?.ToString();
                                    dataRow[Names.DESIRED_TIME_OUT] = record[@"ControllingOUTHHMM"]?.ToString();

                                    dataRow[Names.EMPLOYEE_SHIFT_COMMENT] = record["Comment"]?.ToString();
                                    dataRow[Names.EMPLOYEE_SHIFT] = record[@"Shift"]?.ToString();

                                    dtTemp.Rows.Add(dataRow);
                                }
                            }
                        }
                    }
                    catch
                    { logger.Info("SeekAndShowMembersOfGroup: no any fio in 'selected'"); }
                }

                if (dtTemp.Rows.Count > 0)
                {
                    dtPersonTemp = LeaveAndOrderColumnsOfDataTable(dtTemp, Names.orderColumnsFIO);
                    ShowDatatableOnDatagridview(dtPersonTemp, nameOfLastTable);
                }
            }
        }

        private void AddPersonToGroupItem_Click(object sender, EventArgs e) //AddPersonToGroup() //Add the selected person into the named group
        { AddPersonToGroup(); }

        private void AddPersonToGroup() //Add the selected person into the named group
        {
            logger.Trace($"-= {nameof(AddPersonToGroup)} =-");

            string group = ReturnTextOfControl(textBoxGroup);
            string groupDescription = ReturnTextOfControl(textBoxGroupDescription);
            logger.Trace("AddPersonToGroup: group " + group);
            if (DataGridViewOperations.RowsCount(dataGridView1) > -1)
            {
                string[] cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                    Names.FIO,
                    Names.CODE,
                    Names.DEPARTMENT,
                    Names.EMPLOYEE_POSITION,
                    Names.DESIRED_TIME_IN,
                    Names.DESIRED_TIME_OUT,
                    Names.CHIEF_ID,
                    Names.EMPLOYEE_SHIFT,
                    Names.DEPARTMENT_ID,
                    Names.PLACE_EMPLOYEE
                            });

                using (var connection = new SQLiteConnection(sqLiteLocalConnectionString))
                {
                    connection.Open();
                    if (group?.Length > 0)
                    {
                        using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'PeopleGroupDescription' (GroupPerson, GroupPersonDescription) " +
                                                "VALUES (@GroupPerson, @GroupPersonDescription)", connection))
                        {
                            sqlCommand.Parameters.Add("@GroupPerson", DbType.String).Value = group;
                            sqlCommand.Parameters.Add("@GroupPersonDescription", DbType.String).Value = groupDescription;
                            try { sqlCommand.ExecuteNonQuery(); } catch (Exception ept) { logger.Warn("PeopleGroupDescription" + ept.ToString()); }
                        }
                    }

                    if (group?.Length > 0 && cellValue[1]?.Length > 0)
                    {
                        using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'PeopleGroup' (FIO, NAV, GroupPerson, ControllingHHMM, ControllingOUTHHMM, Department, PositionInDepartment, Comment, Shift, DepartmentId, City, Boss) " +
                                                    "VALUES (@FIO, @NAV, @GroupPerson, @ControllingHHMM, @ControllingOUTHHMM, @Department, @PositionInDepartment, @Comment, @Shift, @DepartmentId, @City, @Boss)", connection))
                        {
                            sqlCommand.Parameters.Add("@FIO", DbType.String).Value = cellValue[0];
                            sqlCommand.Parameters.Add("@NAV", DbType.String).Value = cellValue[1];
                            sqlCommand.Parameters.Add("@GroupPerson", DbType.String).Value = group;
                            sqlCommand.Parameters.Add("@ControllingHHMM", DbType.String).Value = cellValue[4];
                            sqlCommand.Parameters.Add("@ControllingOUTHHMM", DbType.String).Value = cellValue[5];
                            sqlCommand.Parameters.Add("@Department", DbType.String).Value = cellValue[2];
                            sqlCommand.Parameters.Add("@PositionInDepartment", DbType.String).Value = cellValue[3];
                            sqlCommand.Parameters.Add("@Comment", DbType.String).Value = "";
                            sqlCommand.Parameters.Add("@DepartmentId", DbType.String).Value = cellValue[8];
                            sqlCommand.Parameters.Add("@City", DbType.String).Value = cellValue[9];
                            sqlCommand.Parameters.Add("@Boss", DbType.String).Value = cellValue[6];
                            sqlCommand.Parameters.Add("@Shift", DbType.String).Value = cellValue[7];
                            try { sqlCommand.ExecuteNonQuery(); } catch (Exception ept) { logger.Warn("PeopleGroup: " + ept.ToString()); }
                        }
                        SetStatusLabelText(StatusLabel2, "'" + cellValue[0]?.ConvertFullNameToShortForm() + "'" + " добавлен в группу '" + group + "'");
                        SetStatusLabelBackColor(StatusLabel2, SystemColors.Control);
                    }
                    else if (group?.Length > 0 && cellValue[1]?.Length == 0)
                    {
                        SetStatusLabelText(StatusLabel2, "Отсутствует NAV-код у: " + textBoxFIO.Text?.ConvertFullNameToShortForm());
                        SetStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                    }
                    else if (group?.Length == 0 && cellValue[1]?.Length > 0)
                    {
                        SetStatusLabelText(StatusLabel2, "Не указана группа, в которую нужно добавить!");
                        SetStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                    }
                }
            }

            SeekAndShowMembersOfGroup(group);

            labelGroup.BackColor = SystemColors.Control;
            SetMenuItemText(PersonOrGroupItem, Names.WORK_WITH_A_PERSON);
            nameOfLastTable = "PeopleGroup";
        }

        private void importPeopleInLocalDBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            logger.Trace($"-= {nameof(importPeopleInLocalDBToolStripMenuItem_Click)} =-");

            dtPersonTemp?.Clear();
            dtPersonTemp = dtPeople.Copy();
            HashSet<Department> departmentsUniq = new HashSet<Department>();

            ImportTextToTable(dtPersonTemp, ref departmentsUniq);
            WritePeopleInLocalDB(dbApplication.ToString(), dtPersonTemp);
            ImportListGroupsDescriptionInLocalDB(dbApplication.ToString(), departmentsUniq);
            departmentsUniq = null;
        }

        private void ImportTextToTable(DataTable dt, ref HashSet<Department> departmentsUniq) //Fill dtPeople
        {
            logger.Trace($"-= {nameof(ImportTextToTable)} =-");

            List<string> listRows = ReadTXTFile();

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

                    foreach (string strRow in listRows.Distinct())
                    {
                        string[] cell = strRow.Split('\t');
                        if (cell.Length == 7)
                        {
                            row[Names.FIO] = cell[0];
                            row[Names.CODE] = cell[1];
                            row[Names.GROUP] = cell[2];
                            row[Names.DEPARTMENT] = cell[3];
                            row[Names.DEPARTMENT_ID] = "";
                            row[Names.EMPLOYEE_POSITION] = cell[4];

                            departmentsUniq.Add(new Department
                            {
                                DepartmentId = cell[2],
                                DepartmentDescr = cell[3],
                            });

                            checkHourS = cell[5];
                            if (checkHourS.TryParseNumberAsStringToDecimal() == 0)
                            { checkHourS = numUpHourStart.ToString(); }

                            row[Names.DESIRED_TIME_IN] = ConvertStringsTimeToStringHHMM(checkHourS, cell[6]);

                            checkHourE = cell[7];
                            if (checkHourE.TryParseNumberAsStringToDecimal() == 0)
                            { checkHourE = numUpDownHourEnd.ToString(); }

                            row[Names.DESIRED_TIME_OUT] = ConvertStringsTimeToStringHHMM(checkHourE, cell[8]);

                            dt.Rows.Add(row);
                            row = dt.NewRow();
                        }
                    }
                }
            }
            else
            { MessageBox.Show("выбранный файл пустой, или \nне подходит для импорта."); }
        }

        //Write people in local DB
        private void WritePeopleInLocalDB(string pathToPersonDB, DataTable dtSource) //use listGroups /add reserv1 reserv2
        {
            logger.Trace($"{nameof(WritePeopleInLocalDB)}: table - {dtSource}, row - {dtSource.Rows.Count}");

            string query = "INSERT OR REPLACE INTO 'PeopleGroup' (FIO, NAV, GroupPerson, ControllingHHMM, ControllingOUTHHMM, Shift, Comment, Department, PositionInDepartment, DepartmentId, City, Boss) " +
                      "VALUES (@FIO, @NAV, @GroupPerson, @ControllingHHMM, @ControllingOUTHHMM, @Shift, @Comment, @Department, @PositionInDepartment, @DepartmentId, @City, @Boss)";

            using (SqLiteDbWrapper dbWriter = new SqLiteDbWrapper(sqLiteLocalConnectionString, dbApplication))
            {
                logger.Trace($"query: {query}");
                dbWriter.Status += AddLoggerTraceText;

                dbWriter.Execute("begin");
                foreach (var dr in dtSource.AsEnumerable())
                {
                    if (dr[Names.FIO]?.ToString()?.Length > 0 && dr[Names.CODE]?.ToString()?.Length > 0)
                    {
                        using (SQLiteCommand SqlQuery = new SQLiteCommand(query, dbWriter.sqlConnection))
                        {
                            SqlQuery.Parameters.Add("@FIO", DbType.String).Value = dr[Names.FIO]?.ToString();
                            SqlQuery.Parameters.Add("@NAV", DbType.String).Value = dr[Names.CODE]?.ToString();

                            SqlQuery.Parameters.Add("@GroupPerson", DbType.String).Value = dr[Names.GROUP]?.ToString();
                            SqlQuery.Parameters.Add("@Department", DbType.String).Value = dr[Names.DEPARTMENT]?.ToString();
                            SqlQuery.Parameters.Add("@DepartmentId", DbType.String).Value = dr[Names.DEPARTMENT_ID].ToString();
                            SqlQuery.Parameters.Add("@PositionInDepartment", DbType.String).Value = dr[Names.EMPLOYEE_POSITION]?.ToString();
                            SqlQuery.Parameters.Add("@City", DbType.String).Value = dr[Names.PLACE_EMPLOYEE]?.ToString();
                            SqlQuery.Parameters.Add("@Boss", DbType.String).Value = dr[Names.CHIEF_ID]?.ToString();

                            SqlQuery.Parameters.Add("@ControllingHHMM", DbType.String).Value = dr[Names.DESIRED_TIME_IN]?.ToString();
                            SqlQuery.Parameters.Add("@ControllingOUTHHMM", DbType.String).Value = dr[Names.DESIRED_TIME_OUT]?.ToString();

                            SqlQuery.Parameters.Add("@Shift", DbType.String).Value = dr[Names.EMPLOYEE_SHIFT]?.ToString();
                            SqlQuery.Parameters.Add("@Comment", DbType.String).Value = dr[Names.EMPLOYEE_SHIFT_COMMENT]?.ToString();

                            dbWriter.ExecuteBulk(SqlQuery);
                            ProgressWork1Step();
                        }
                    }
                }

                dbWriter.Execute("end");
                dbWriter.Execute("begin");

                query = "INSERT OR REPLACE INTO 'LastTakenPeopleComboList' (ComboList) VALUES (@ComboList)";
                logger.Trace($"query: {query}");
                foreach (var str in listFIO)
                {
                    if (str.fio?.Length > 0 && str.Code?.Length > 0)
                    {
                        using (SQLiteCommand SqlQuery = new SQLiteCommand(query, dbWriter.sqlConnection))
                        {
                            SqlQuery.Parameters.Add("@ComboList", DbType.String).Value = str.fio + "|" + str.Code;

                            dbWriter.ExecuteBulk(SqlQuery);
                            ProgressWork1Step();
                        }
                    }
                }
                dbWriter.Execute("end");

                dbWriter.Status -= AddLoggerTraceText;
            }

            foreach (var str in listFIO)
            { AddComboBoxItem(comboBoxFio, str.fio + "|" + str.Code); }
            ProgressWork1Step();
            if (ReturnComboBoxCountItems(comboBoxFio) > 0)
            { SetComboBoxIndex(comboBoxFio, 0); }
            ProgressWork1Step();

            int.TryParse(listFIO.Count.ToString(), out countUsers);
            logger.Info("Записано ФИО: " + countUsers);
        }

        private void ImportListGroupsDescriptionInLocalDB(string pathToPersonDB, HashSet<Department> departmentsUniq) //use listGroups
        {
            logger.Trace($"-= {nameof(ImportListGroupsDescriptionInLocalDB)} =-");

            string query = "INSERT OR REPLACE INTO 'PeopleGroupDescription' (GroupPerson, GroupPersonDescription, Recipient) " +
                                       "VALUES (@GroupPerson, @GroupPersonDescription, @Recipient)";

            using (SqLiteDbWrapper dbWriter = new SqLiteDbWrapper(sqLiteLocalConnectionString, dbApplication))
            {
                logger.Trace($"query: {query}");
                dbWriter.Status += AddLoggerTraceText;

                dbWriter.Execute("begin");
                foreach (var group in departmentsUniq)
                {
                    if (group?.DepartmentId?.Length > 0)
                    {
                        using (SQLiteCommand SqlQuery = new SQLiteCommand(query, dbWriter.sqlConnection))
                        {
                            SqlQuery.Parameters.Add("@GroupPerson", DbType.String).Value = group.DepartmentId;
                            SqlQuery.Parameters.Add("@GroupPersonDescription", DbType.String).Value = group.DepartmentDescr;
                            SqlQuery.Parameters.Add("@Recipient", DbType.String).Value = group.DepartmentBossCode;

                            dbWriter.ExecuteBulk(SqlQuery);
                        }
                    }
                }
                dbWriter.Execute("end");
                dbWriter.Status -= AddLoggerTraceText;
            }
        }

        private void GetNamesOfPassagePoints() //Get names of the pass by points
        {
            logger.Trace($"-= {nameof(GetNamesOfPassagePoints)} =-");

            collectionOfPassagePoints = new CollectionOfPassagePoints();

            string query = "Select id, name FROM OBJ_ABC_ARC_READER;";
            logger.Trace($"query: {query}");
            string idPoint, namePoint, direction;

            using (SqlDbReader sqlDbTableReader = new SqlDbReader(sqlServerConnectionString))
            using (System.Data.SqlClient.SqlDataReader sqlData = sqlDbTableReader.GetData(query))
            {
                foreach (DbDataRecord record in sqlData)
                {
                    namePoint = record?["name"]?.ToString()?.Trim();
                    idPoint = record["id"]?.ToString()?.Trim();
                    if (namePoint?.ToLower()?.Contains("выход") == true)
                    { direction = "Выход"; }
                    else
                    { direction = "Вход"; }

                    if (idPoint?.Length > 0 && namePoint?.Length > 0)
                    {
                        collectionOfPassagePoints.AddPoint(idPoint, namePoint, direction, sServer1);
                    }
                }
            }

            foreach (var tmp in collectionOfPassagePoints.GetCollection())
            {
                logger.Trace(tmp.Key + " " + tmp.Value._idPoint + " " + tmp.Value._namePoint + " " + tmp.Value._direction + " " + tmp.Value._connectedToServer);
            }
        }

        private void ResetFilterLoadLastIputsOutput_Click(object sender, EventArgs e)
        {
            _selectedEmployeeVisitor = null;
            AddVisitorsAtDataGridView(visitors);
        }

        private async void LoadLastIputsOutputs_Click(object sender, EventArgs e) //LoadInputsOutputsOfVisitors()
        {
            //how many times continiously to check registrations at the server
            int timesCheckingRegistration = 10;

            //clear painting;
            _paintedEmployeeVisitor = null;
            _selectedEmployeeVisitor = null;

            // status of repeatedly loading of registrations cards from server
            checkInputsOutputs = false;

            await Task.Run(() => LoadInputsOutputsOfVisitors(DateTime.Now.ToYYYYMMDD(), DateTime.Now.ToYYYYMMDD(), timesCheckingRegistration));
        }

        private async void LoadInputsOutputsItem_Click(object sender, EventArgs e)
        {
            //how many times continiously to check registrations at the server
            int timesCheckingRegistration = 1;

            //clear painting;
            _paintedEmployeeVisitor = null;
            _selectedEmployeeVisitor = null;

            //status of repeatedly loading of registrations cards from server
            checkInputsOutputs = false;

            string dayStart = ReturnDateTimePicker(dateTimePickerStart).ToYYYYMMDD();
            string dayEnd = ReturnDateTimePicker(dateTimePickerEnd).ToYYYYMMDD();

            await Task.Run(() => LoadInputsOutputsOfVisitors(dayStart, dayEnd, timesCheckingRegistration));
        }

        private async void LoadLastIputsOutputs_Update_Click(object sender, EventArgs e) //LoadInputsOutputsOfVisitors()
        {
            //how many times continiously to check registrations at the server
            int timesCheckingRegistration = 10;
            _selectedEmployeeVisitor = null;

            //clear painting;
            _paintedEmployeeVisitor = null;

            //status of repeatedly loading of registrations cards from server
            checkInputsOutputs = false;

            string dayStart = ReturnDateTimePicker(dateTimePickerStart).ToYYYYMMDD();
            string dayEnd = ReturnDateTimePicker(dateTimePickerEnd).ToYYYYMMDD();

            if (DateTime.Now.ToYYYYMMDD() != dayStart)
            { timesCheckingRegistration = 1; }

            await Task.Run(() => LoadInputsOutputsOfVisitors(dayStart, dayEnd, timesCheckingRegistration));
        }

        private static Visitors visitors = new Visitors();

        //lock to loading of registrations cards from server
        private static readonly object lockerToLoadInsOuts = new object();

        //lock to show data on datagridview
        private static readonly object lockerToShowData = new object();

        //status of repeatedly loading of registrations cards from server
        private static bool checkInputsOutputs = true;

        //Timer for waiting next loading
        private IStartStopTimer startStopTimer;

        private void LoadInputsOutputsOfVisitors(string startDay, string endDay, int timesChecking)
        {
            logger.Trace($"-= {nameof(LoadInputsOutputsOfVisitors)} =-");

            CheckAliveIntellectServer(sServer1, sqlServerConnectionString, queryCheckSystemDdSCA);

            if (bServer1Exist)
            {
                EnableControl(comboBoxFio, true);
                nameOfLastTable = "LastIputsOutputs";
                List<Visitor> visitorsTillNow;

                //Get names of the points
                GetNamesOfPassagePoints();

                visitors = new Visitors();
                //subscribe on action of changed data in collection of visitors
                visitors.collection.CollectionChanged += VisitorsCollectionChanged;

                checkInputsOutputs = true;
                firstStage = true;
                startStopTimer = new StartStopTimer(15);

                string startTime = "00:00:00";
                string endTime = "23:59:59";
                do
                {
                    lock (lockerToLoadInsOuts)
                    {
                        visitorsTillNow = new List<Visitor>();
                        dgvo.Show(dataGridView1, false);

                        logger.Trace("LoadInputsOutputsOfVisitors: " + startDay + " " + startTime + " - " + endDay + " " + endTime);

                        visitorsTillNow = GetInputsOutputs(ref startDay, ref startTime, ref endDay, ref endTime);
                        if (firstStage)
                        {
                            //put registrations in the list with order from the newest data in the top to the oldest ones in the end
                            visitors.collection.AddRange(visitorsTillNow);
                            firstStage = false;
                        }
                        else
                        {
                            if (visitorsTillNow?.Count > 0)
                            {
                                //put new registrations in the top of list
                                visitorsTillNow.Reverse();
                                foreach (var visitor in visitorsTillNow)
                                { visitors.Add(visitor, 0); }
                            }
                        }
                    }

                    dgvo.Show(dataGridView1, true);
                    timesChecking--;
                    if (timesChecking <= 0)
                    {
                        checkInputsOutputs = false;
                    }
                    else
                    {
                        SetStatusLabelText(StatusLabel2, $"Загружены данные о регистрации пропусков до: {startDay} {startTime}");
                        startStopTimer.WaitTime();
                    }
                } while (checkInputsOutputs);

                visitors.collection.CollectionChanged -= VisitorsCollectionChanged;
                SetStatusLabelText(StatusLabel2, $"Сбор данных регистрации пропусков за {startDay} завершен");
            }
            else
            {
                SetStatusLabelText(StatusLabel2, $"Сервер с базой регистрации пропусков {sServer1} не доступен");
            }
        }

        private bool firstStage = true;

        private List<Visitor> GetInputsOutputs(ref string startDay, ref string startTime, ref string endDay, ref string endTime)
        {
            ProgressBar1Start();

            List<Visitor> visitors = new List<Visitor>();
            bool startTimeNotSet = true;
            SideOfPassagePoint sideOfPassagePoint;
            string time, date, fio, action, action_descr, fac, card, idCardDescr;
            int idCard = 0;
            /*
                        string query = "SELECT p.param0 as param0, p.param1 as param1, p.objid as objid, p.objtype, p.action as action, " +
                        " pe.tabnum as nav, pe.facility_code as fac, pe.card as card, " +
                        " CONVERT(varchar, p.date, 120) AS date, CONVERT(varchar, p.time, 114) AS time" +
                        " FROM protocol p " +
                        " LEFT JOIN OBJ_PERSON pe ON  p.param1=pe.id " +
                        " where p.objtype like 'ABC_ARC_READER' " +
                        " AND p.param0 like '%%' " +
                        // " AND p.param1 like '" + person.idCard + "' " +
                        $" AND date > '{startDay} {startTime}' AND date <= '{endDay} {endTime}' " +
                        " ORDER BY date DESC, time DESC;"; //sorting

                        logger.Trace($"query: {query}");*/

            using (ProtocolConnector db = new ProtocolConnector(sqlServerConnectionString))
            {
                DateTime dtStart = DateTime.Parse(startDay + " " + startTime);
                DateTime dtEnd = DateTime.Parse(endDay + " " + endTime);

                // получаем объекты из бд
                //query
                /*   var RegisteredAction = (from registeredAction in db.ProtocolObjects
                                          join libraryUsers in db.PersonObjects
                                          on registeredAction.IdCard equals libraryUsers.id
                                          where dtStart < registeredAction.ActionDate && registeredAction.ActionDate < dtEnd
                                          orderby registeredAction.ActionTime descending
                                          select new
                                          {
                                              FIO = registeredAction.FIO,
                                              IdCard = registeredAction.IdCard,
                                              ActionDate = registeredAction.ActionDate,
                                              ActionTime = registeredAction.ActionTime,
                                              ActionDescr = registeredAction.ActionDescr,
                                              ActionType = registeredAction.ActionType,
                                              PointName = registeredAction.PointName,
                                              fac = libraryUsers.facility_code,
                                              code = libraryUsers.tabnum,
                                              card = libraryUsers.card
                                          }).ToList();*/

                //lambda
                var RegisteredAction = db.ProtocolObjects
                    .Join(
                        db.PersonObjects,
                        registeredAction => registeredAction.IdCard,
                        libraryUsers => libraryUsers.id,
                        (registeredAction, libraryUsers) => new
                        {
                            registeredAction.FIO,
                            registeredAction.IdCard,
                            registeredAction.ActionDate,
                            registeredAction.ActionTime,
                            registeredAction.ActionDescr,
                            registeredAction.ActionType,
                            registeredAction.PointName,
                            libraryUsers.facility_code,
                            libraryUsers.card
                        }
                        )
                    .Where(x => x.ActionDate > dtStart)
                    .Where(x => x.ActionDate <= dtEnd)
                    .Where(x => x.ActionType == "ABC_ARC_READER")
                    .OrderByDescending(x => x.ActionDate)
                    .ToList();

                foreach (var v in RegisteredAction)
                {
                    if (v != null)
                    {
                        date = v.ActionDate?.ToString()?.Split(' ')[0];
                        time = v.ActionTime?.ToString()?.Split(' ')[1];//?.ConvertTimeIntoStandartTime();

                        //look for FIO
                        fio = v.FIO?.Length > 0 ? v.FIO : sServer1;

                        //look for  idCard
                        idCard = 0;
                        int.TryParse(v.IdCard, out idCard);
                        fac = v.facility_code;
                        card = v.card;
                        idCardDescr = idCard != 0 ? $"№{idCard} ({fac},{card})" : (fio == sServer1 ? "" : "Пропуск не зарегистрирован");

                        //look for action with idCard
                        action = v.ActionDescr;
                        action_descr = null;
                        if (Names.CARD_REGISTERED_ACTION.TryGetValue(action, out action_descr))
                        { action = "Сервисное сообщение"; }
                        else if (fio == sServer1)
                        { action_descr = action; }

                        sideOfPassagePoint = collectionOfPassagePoints.GetPoint(v.PointName);

                        //write gathered data in the collection
                        try
                        {
                            logger.Trace($"{fio} {action_descr} {idCard} {idCardDescr} {v?.ActionDescr} {date} {time} {sideOfPassagePoint?._namePoint} {sideOfPassagePoint?._direction}");

                            visitors.Add(new Visitor(fio, idCardDescr, date, time, action_descr, sideOfPassagePoint));
                            if (startTimeNotSet) //set starttime into last time at once
                            {
                                startTime = time;
                                startTimeNotSet = false;
                                firstStage = false;
                            }
                        }
                        catch (Exception e)
                        {
                            logger.Warn($"{date} {time}");
                            logger.Warn(fio);
                            logger.Warn($"{action_descr} {idCard} {idCardDescr} {v?.ActionDescr}");
                            logger.Warn(sideOfPassagePoint?._namePoint + " " + sideOfPassagePoint?._direction);
                            logger.Warn(e.ToString());
                        }

                        ProgressWork1Step();
                    }
                }
            }

            /*  using (SqlDbReader sqlDbTableReader = new SqlDbReader(sqlServerConnectionString))
              {
                  System.Data.SqlClient.SqlDataReader sqlData = sqlDbTableReader.GetData(query);
                  foreach (DbDataRecord record in sqlData)
                  {
                      if (record != null)
                      {   //look for PassagePoint
                          fullPointName = record["objid"]?.ToString()?.Trim();
                          sideOfPassagePoint = collectionOfPassagePoints.GetPoint(fullPointName);

                          //look for FIO
                          fio = record["param0"]?.ToString()?.Trim()?.Length > 0 ? record["param0"]?.ToString()?.Trim() : sServer1;

                          //look for date and time
                          date = record["date"]?.ToString()?.Trim()?.Split(' ')[0];
                          time = ((string)record["time"]?.ToString()?.Trim()).ConvertTimeIntoStandartTime();

                          //look for  idCard
                          idCard = 0;
                          int.TryParse(record["param1"]?.ToString()?.Trim(), out idCard);
                          fac = record["fac"]?.ToString()?.Trim();
                          card = record["card"]?.ToString()?.Trim();
                          idCardDescr = idCard != 0 ? "№" + idCard + " (" + fac + "," + card + ")" : (fio == sServer1 ? "" : "Пропуск не зарегистрирован");

                          //look for action with idCard
                          action = record["action"]?.ToString()?.Trim();
                          action_descr = null;
                          if (Names.CARD_REGISTERED_ACTION.TryGetValue(action, out action_descr))
                          { action = "Сервисное сообщение"; }
                          else if (fio == sServer1)
                          { action_descr = action; }

                          //write gathered data in the collection
                           logger.Trace(fio + " " + action_descr + " " + idCard + " " + idCardDescr + " " + record["action"]?.ToString()?.Trim() + " " + date + " " + time + " " + sideOfPassagePoint._namePoint + " " + sideOfPassagePoint._direction);

                          visitors.Add(new Visitor(fio, idCardDescr, date, time, action_descr, sideOfPassagePoint));

                          if (startTimeNotSet) //set starttime into last time at once
                          {
                              startTime = time;
                              startTimeNotSet = false;
                          }

                          _ProgressWork1Step();
                      }
                  }
                   logger.Trace("visitors added: " + visitors.Count());
              }*/

            stimerPrev = "";
            ProgressBar1Stop();
            return visitors;
        }

        private void VisitorsCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (nameOfLastTable != "LastIputsOutputs")//Visitor selected=null
            { checkInputsOutputs = false; }

            lock (lockerToShowData)
            { AddVisitorsAtDataGridView(visitors, _selectedEmployeeVisitor); }
        }

        private void AddVisitorsAtDataGridView(Visitors visitors_, Visitor selected = null)
        {
            using (DataTable dtTemp = dtPeople.Clone())
            {
                SendListLastRegistrationsToDataTable(visitors_.collection, dtTemp, selected);

                using (DataTable dt = LeaveAndOrderColumnsOfDataTable(dtTemp, Names.orderColumnsLastRegistrations))
                {
                    dgvo.AddDataTable(dataGridView1, dt);
                    nameOfLastTable = "LastIputsOutputs";

                    if (_paintedEmployeeVisitor != null)
                    { dgvo.Paint(dataGridView1, Names.FIO, _paintedEmployeeVisitor.fio); }
                }
            }
        }

        private void SendListLastRegistrationsToDataTable(ObservableRangeCollection<Visitor> _visitors, DataTable dt, Visitor selected = null)
        {
            DataRow row;
            foreach (var visitor in _visitors.ToArray())
            {
                if (visitor != null)
                {
                    row = dt.NewRow();
                    row[Names.FIO] = visitor.fio;
                    row[Names.DATE_REGISTRATION] = visitor.date;
                    row[Names.TIME_REGISTRATION_STRING] = visitor.time;
                    row[Names.N_ID_STRING] = visitor.idCard;
                    row[Names.CHECKPOINT_DIRECTION] = visitor.sideOfPassagePoint;
                    row[Names.CHECKPOINT_ACTION] = visitor.action;

                    if (selected == null || selected.fio == visitor.fio)
                    { dt.Rows.Add(row); }
                }
            }
        }

        private static EmployeeVisitor _paintedEmployeeVisitor; //painted visitor in datagridview1
        private static EmployeeVisitor _selectedEmployeeVisitor; //painted visitor in datagridview1

        private void PaintRowsActionItem_Click(object sender, EventArgs e)
        {
            _paintedEmployeeVisitor = LookForSelectedVisitor(dataGridView1);
            dgvo.Paint(dataGridView1, Names.CHECKPOINT_ACTION, _paintedEmployeeVisitor.action);
        }

        private void PaintRowsFioItem_Click(object sender, EventArgs e)
        {
            _paintedEmployeeVisitor = LookForSelectedVisitor(dataGridView1);
            dgvo.Paint(dataGridView1, Names.FIO, _paintedEmployeeVisitor.fio);
        }

        private EmployeeVisitor LookForSelectedVisitor(DataGridView dgv)
        {
            EmployeeVisitor visitor = new EmployeeVisitor();
            string[] cellValue = dgvo.FindValuesInCurrentRow(dgv, new string[] {
                Names.FIO,
                Names.N_ID_STRING,
                Names.CHECKPOINT_ACTION
                    });
            visitor.fio = cellValue[0];
            visitor.idCard = cellValue[1];
            visitor.action = cellValue[2];

            return visitor;
        }

        private void GetDataOfGroup_Click(object sender, EventArgs e) //LoadIdCardRegistrations()
        {
            string group = ReturnTextOfControl(textBoxGroup);

            dateTimePickerStart.Value = DateTime.Now.FirstDayOfMonth();

            nameOfLastTable = @"ListFIO";

            LoadIdCardRegistrations(group);
        }

        private void GetDataOfPerson_Click(object sender, EventArgs e) //LoadIdCardRegistrations()
        {
            dateTimePickerStart.Value = DateTime.Now.FirstDayOfMonth();

            LoadIdCardRegistrations(null);
        }

        private async void LoadIdCardRegistrations(string _group) //GetData()
        {
            logger.Trace($"-= {nameof(LoadIdCardRegistrations)} =-");

            ProgressBar1Start();

            await Task.Run(() => CheckAliveIntellectServer(sServer1, sqlServerConnectionString, queryCheckSystemDdSCA));

            ChangeControlBackColor(groupBoxPeriod, SystemColors.Control);
            ChangeControlBackColor(groupBoxTimeStart, SystemColors.Control);
            ChangeControlBackColor(groupBoxTimeEnd, SystemColors.Control);
            SetMenuItemBackColor(LoadDataItem, SystemColors.Control);

            EnableMenuItem(LoadDataItem, false);
            EnableMenuItem(FunctionMenuItem, false);
            EnableMenuItem(SettingsMenuItem, false);
            EnableMenuItem(GroupsMenuItem, false);
            CheckBoxesFiltersAll_Enable(false);

            VisibleControl(dataGridView1, false);

            if (bServer1Exist)
            {
                reportStartDay = ReturnDateTimePicker(dateTimePickerStart).ToYYYYMMDD();
                reportLastDay = ReturnDateTimePicker(dateTimePickerEnd).ToYYYYMMDD();

                await Task.Run(() => GetData(_group, reportStartDay, reportLastDay));

                SetStatusLabelForeColor(StatusLabel2, Color.Black);
                SetMenuItemBackColor(LoadDataItem, SystemColors.Control);
                SetMenuItemBackColor(TableExportToExcelItem, Color.PaleGreen);

                EnableMenuItem(LoadDataItem, true);
                EnableMenuItem(FunctionMenuItem, true);
                EnableMenuItem(SettingsMenuItem, true);
                EnableMenuItem(GroupsMenuItem, true);
                EnableMenuItem(VisualModeItem, true);
                VisibleMenuItem(VisualModeItem, true);
                EnableMenuItem(TableModeItem, true);
                EnableMenuItem(TableExportToExcelItem, true);

                VisibleControl(dataGridView1, true);

                EnableControl(checkBoxReEnter, true);
                EnableControl(checkBoxTimeViolations, false);
                EnableControl(checkBoxWeekend, false);
                EnableControl(checkBoxCelebrate, false);
                CheckBoxesFiltersAll_SetState(false);

                panelViewResize(numberPeopleInLoading);
                ChangeControlBackColor(groupBoxFilterReport, Color.PaleGreen);
            }
            else
            {
                GetInfoSetup();
                EnableMenuItem(SettingsMenuItem, true);
            }
            stimerPrev = "";
            ProgressBar1Stop();

            if (dtPersonTemp?.Rows.Count > 0)
            {
                VisibleMenuItem(TableExportToExcelItem, true);
                SetStatusLabelText(StatusLabel2, "Данные регистрации пропусков загружены");
            }
        }

        private void GetData(string _group, string _reportStartDay, string _reportLastDay)
        {
            logger.Trace($"-= {nameof(GetData)} =-");

            //Clear work tables
            dtPersonRegistrationsFullList.Clear();

            dtPeople.DefaultView.Sort = Names.GROUP + ", " + Names.FIO + ", " + Names.DATE_REGISTRATION + ", " + Names.TIME_REGISTRATION + ", " + Names.REAL_TIME_IN + ", " + Names.REAL_TIME_OUT + " ASC";

            //Clone default column name and structure from 'dtPeople' to other DataTables
            dtPersonTemp = dtPeople.Clone();  //Copy only structure(Name of columns)
            dtPersonRegistrationsFullList = dtPeople.Clone();  //Copy only structure(Name of columns)
            dtPeopleGroup = dtPeople.Clone();  //Copy only structure(Name of columns)

            logger.Trace("GetData: " + _group + " " + _reportStartDay + " " + _reportStartDay);

            //Get names of the points
            GetNamesOfPassagePoints();

            //Load Records of registrations of Id cards
            LoadRecords(_group,
                ReturnDateTimePicker(dateTimePickerStart).ToYYYYMMDD() + " 00:00:00",
                ReturnDateTimePicker(dateTimePickerEnd).ToYYYYMMDD() + " 23:59:59",
                "");

            dtPersonTemp = LeaveAndOrderColumnsOfDataTable(dtPersonRegistrationsFullList.Copy(), Names.orderColumnsRegistrations);

            //show selected data  within the selected collumns
            ShowDatatableOnDatagridview(dtPersonTemp, "PeopleGroup");
        }

        private void LoadRecords(string nameGroup, string startDate, string endDate, string doPostAction)
        {
            logger.Trace($"-= {nameof(LoadRecords)} =-");

            EmployeeFull person = new EmployeeFull();
            outPerson = new List<OutPerson>();
            outResons = new List<PeopleOutReasons>
            {
                new PeopleOutReasons()
                { id = "0", hourly = 1, nameReason = @"Unknow", visibleNameReason = @"Неизвестная" }
            };

            string query = "Select id, name,hourly, visibled_name FROM out_reasons";
            logger.Trace(query);
            using (MySqlDbReader mysqlDbTableReader = new MySqlDbReader(mysqlServerConnectionString))
            {
                MySql.Data.MySqlClient.MySqlDataReader mysqlData = mysqlDbTableReader.GetData(query);

                while (mysqlData.Read())
                {
                    if (mysqlData?.GetString(@"id")?.Length > 0 && mysqlData?.GetString(@"name")?.Length > 0)
                    {
                        outResons.Add(new PeopleOutReasons()
                        {
                            id = mysqlData.GetString(@"id"),
                            hourly = Convert.ToInt32(mysqlData.GetString(@"hourly")),
                            nameReason = mysqlData.GetString(@"name"),
                            visibleNameReason = mysqlData.GetString(@"visibled_name")
                        });
                    }
                }
            }
            ProgressWork1Step();

            string date = "";
            string resonId = "";
            query = "Select * FROM out_users where reason_date >= '" + startDate.Split(' ')[0] + "' AND reason_date <= '" + endDate.Split(' ')[0] + "' ";
            logger.Trace(query);
            using (MySqlDbReader mysqlDbTableReader = new MySqlDbReader(mysqlServerConnectionString))
            {
                MySql.Data.MySqlClient.MySqlDataReader mysqlData = mysqlDbTableReader.GetData(query);

                while (mysqlData.Read())
                {
                    if (mysqlData?.GetString(@"reason_id")?.Length > 0 && mysqlData?.GetString(@"user_code")?.Length > 0)
                    {
                        resonId = outResons.Find((x) => x.id == mysqlData.GetString(@"reason_id")).id;
                        try { date = DateTime.Parse(mysqlData.GetString(@"reason_date")).ToYYYYMMDD(); } catch { date = ""; }

                        outPerson.Add(new OutPerson()
                        {
                            id = resonId,
                            code = mysqlData.GetString(@"user_code").ToUpper()?.Replace('C', 'S'),
                            date = date,
                            from = ConvertStringsTimeToSeconds(mysqlData.GetString(@"from_hour"), mysqlData.GetString(@"from_min")),
                            to = ConvertStringsTimeToSeconds(mysqlData.GetString(@"to_hour"), mysqlData.GetString(@"to_min")),
                            hourly = 0
                        });
                    }
                }
                logger.Trace("Всего с " + startDate.Split(' ')[0] + " по " + endDate.Split(' ')[0] + " на сайте есть - " + outPerson.Count + " записей с отсутствиями");
            }
            ProgressWork1Step();

            if ((nameOfLastTable == "PeopleGroupDescription" || nameOfLastTable == "PeopleGroup" || nameOfLastTable == "Mailing" ||
                nameOfLastTable == "ListFIO" || doPostAction == "sendEmail") && nameGroup?.Length > 0)
            {
                SetStatusLabelText(StatusLabel2, "Получаю данные по группе " + nameGroup);
                dtPeopleGroup = LoadGroupMembersFromDbToDataTable(nameGroup);

                logger.Trace("LoadRecords, DT - " + dtPeopleGroup.TableName + " , всего записей - " + dtPeopleGroup.Rows.Count);
                foreach (DataRow row in dtPeopleGroup.Rows)
                {
                    if (row[Names.FIO]?.ToString()?.Length > 0 && (row[Names.GROUP]?.ToString() == nameGroup || ("@" + row[Names.DEPARTMENT_ID]?.ToString()) == nameGroup))
                    {
                        person = new EmployeeFull()
                        {
                            fio = row[Names.FIO].ToString(),
                            Code = row[Names.CODE]?.ToString(),
                            GroupPerson = nameGroup,
                            Department = row[Names.DEPARTMENT]?.ToString(),
                            City = row[Names.PLACE_EMPLOYEE]?.ToString(),
                            DepartmentBossCode = row[Names.CHIEF_ID]?.ToString(),
                            PositionInDepartment = row[Names.EMPLOYEE_POSITION]?.ToString(),
                            DepartmentId = row[Names.DEPARTMENT_ID]?.ToString(),
                            ControlInHHMM = row[Names.DESIRED_TIME_IN]?.ToString() ?? "9",
                            ControlOutHHMM = row[Names.DESIRED_TIME_OUT]?.ToString() ?? "18",
                            Comment = row[Names.EMPLOYEE_SHIFT_COMMENT]?.ToString(),
                            Shift = row[Names.EMPLOYEE_SHIFT]?.ToString()
                        };
                        person.ControlInSeconds = ((string)person?.ControlInHHMM).ConvertTimeAsStringToSeconds();
                        person.ControlOutSeconds = ((string)person?.ControlOutHHMM).ConvertTimeAsStringToSeconds();

                        GetPersonRegistrationFromServer(ref dtPersonRegistrationsFullList, person, startDate, endDate);     //Search Registration at checkpoints of the selected person
                    }
                }
                nameOfLastTable = "PeopleGroup";
                SetStatusLabelText(StatusLabel2, $"Данные по группе '{nameGroup}' получены");
            }
            else
            {
                person = new EmployeeFull
                {
                    Code = ReturnTextOfControl(textBoxNav),
                    fio = ReturnTextOfControl(textBoxFIO),
                    GroupPerson = "One User",
                    Department = "",
                    DepartmentId = "",
                    City = "",
                    DepartmentBossCode = "",
                    PositionInDepartment = "Сотрудник",
                    Shift = "",
                    Comment = "",
                    ControlInHHMM = ConvertStringsTimeToStringHHMM(ReturnNumUpDown(numUpDownHourStart).ToString(), ReturnNumUpDown(numUpDownMinuteStart).ToString()),
                    ControlOutHHMM = ConvertStringsTimeToStringHHMM(ReturnNumUpDown(numUpDownHourEnd).ToString(), ReturnNumUpDown(numUpDownMinuteEnd).ToString())
                };
                
                SetStatusLabelText(StatusLabel2, $"Получаю данные по {person.fio?.ConvertFullNameToShortForm()}'");

                logger.Trace($"LoadRecords,One User: {person.fio}");

                GetPersonRegistrationFromServer(ref dtPersonRegistrationsFullList, person, startDate, endDate);

                SetStatusLabelText(StatusLabel2, $"Данные с СКД по '{ReturnTextOfControl(textBoxFIO)?.ConvertFullNameToShortForm()}' получены!");
            }
        }

        private void GetPersonRegistrationFromServer(ref DataTable dtTarget, EmployeeFull person, string startDay, string endDay)
        {
            SideOfPassagePoint sideOfPassagePoint;

            logger.Trace($"-= {nameof(GetPersonRegistrationFromServer)}=- , person {person.Code}");

            SeekAnualDays(ref dtTarget, ref person, false,
                startDay.ConvertDateAsStringToIntArray(),
                endDay.ConvertDateAsStringToIntArray(),
                ref myBoldedDates, ref workSelectedDays);
            DataRow rowPerson;
            string query = "";
            HashSet<string> personWorkedDays = new HashSet<string>();

            string[] cellData;
            string namePoint = "";
            string direction = "";

            string stringDataNew = "";
            string fullPointName = "";
            int seconds = 0;

            //look for a NAV and an idCard
            try
            {
                //list of users id and their id cards
                query = "Select id, tabnum FROM OBJ_PERSON where tabnum like '" + person.Code + "';";
                logger.Trace(query);

                //is looking for the idCard of the person's NAV
                using (SqlDbReader sqlDbTableReader = new SqlDbReader(sqlServerConnectionString))
                {
                    System.Data.SqlClient.SqlDataReader sqlData = sqlDbTableReader.GetData(query);
                    foreach (DbDataRecord record in sqlData)
                    {
                        if (record?["tabnum"]?.ToString()?.Trim() == person.Code)
                        {
                            person.idCard = Convert.ToInt32(record["id"].ToString().Trim());
                            break;
                        }
                        ProgressWork1Step();
                    }
                }

                // Passes By Points
                query = "SELECT p.param0 as param0, p.param1 as param1, p.objid as objid, p.objtype, p.action as action, " +
                       " pe.tabnum as nav, pe.facility_code as fac, pe.card as card, " +
                       " CONVERT(varchar, p.date, 120) AS date, CONVERT(varchar, p.time, 114) AS time" +
                       " FROM protocol p " +
                       " LEFT JOIN OBJ_PERSON pe ON  p.param1=pe.id " +
                       " where p.objtype like 'ABC_ARC_READER' " +
                       " AND p.param0 like '%%' " +
                       " AND p.param1 like '" + person.idCard + "' " +
                       " AND date > '" + startDay + "' AND date <= '" + endDay + "' " +
                       " ORDER BY date DESC, time DESC;"; //sorting

                logger.Trace(query);
                using (SqlDbReader sqlDbTableReader = new SqlDbReader(sqlServerConnectionString))
                {
                    System.Data.SqlClient.SqlDataReader sqlData = sqlDbTableReader.GetData(query);
                    foreach (DbDataRecord record in sqlData)
                    {
                        try
                        {
                            if (record["param0"]?.ToString()?.Trim()?.Length > 0 && record["param1"]?.ToString()?.Length > 0)
                            {
                                //  logger.Trace(person.NAV);
                                stringDataNew = record["date"]?.ToString()?.Trim()?.Split(' ')[0];
                                person.idCard = Convert.ToInt32(record["param1"].ToString().Trim());
                                seconds = ((string)record["time"]?.ToString()?.Trim()).ConvertTimeAsStringToSeconds();

                                fullPointName = record["objid"]?.ToString()?.Trim();
                                sideOfPassagePoint = collectionOfPassagePoints.GetPoint(fullPointName);
                                namePoint = sideOfPassagePoint._namePoint;
                                direction = sideOfPassagePoint._direction;

                                personWorkedDays.Add(stringDataNew);

                                rowPerson = dtTarget.NewRow();
                                rowPerson[Names.FIO] = record["param0"].ToString().Trim();
                                rowPerson[Names.CODE] = person.Code;
                                rowPerson[Names.N_ID] = person.idCard;
                                rowPerson[Names.GROUP] = person.GroupPerson;
                                rowPerson[Names.DEPARTMENT] = person.Department;
                                rowPerson[Names.EMPLOYEE_POSITION] = person.PositionInDepartment;
                                rowPerson[Names.PLACE_EMPLOYEE] = person.City;
                                rowPerson[Names.EMPLOYEE_SHIFT] = person.Shift;

                                //day of registration. real data
                                rowPerson[Names.DATE_REGISTRATION] = stringDataNew;
                                rowPerson[Names.TIME_REGISTRATION] = seconds;

                                rowPerson[Names.SERVER_SKD] = sServer1;
                                rowPerson[Names.CHECKPOINT_NAME] = namePoint;
                                rowPerson[Names.CHECKPOINT_DIRECTION] = direction;
                                rowPerson[Names.DESIRED_TIME_IN] = person.ControlInHHMM;
                                rowPerson[Names.DESIRED_TIME_OUT] = person.ControlOutHHMM;
                                rowPerson[Names.REAL_TIME_IN] = seconds.ConvertSecondsToStringHHMMSS();

                                dtTarget.Rows.Add(rowPerson);

                                logger.Trace(rowPerson[Names.FIO] + " " + stringDataNew + " " + seconds + " " + namePoint + " " + direction);
                                ProgressWork1Step();
                            }
                        }
                        catch (DbException dbexpt)
                        { logger.Warn(@"Ошибка доступа к данным БД: " + dbexpt.ToString()); }
                        catch (Exception err) { logger.Warn("GetPersonRegistrationFromServer: " + err.ToString()); }
                    }
                }

                stringDataNew = null; query = null;
                ProgressWork1Step();
            }
            catch (DbException dbexpt)
            { logger.Warn(@"Ошибка доступа к данным БД: " + dbexpt.ToString()); }
            catch (Exception err)
            { logger.Warn(@"GetPersonRegistrationFromServer: " + err.ToString()); }

            // рабочие дни в которые отсутствовал данная персона
            foreach (string day in workSelectedDays.Except(personWorkedDays))
            {
                rowPerson = dtTarget.NewRow();
                rowPerson[Names.FIO] = person.fio;
                rowPerson[Names.CODE] = person.Code;
                rowPerson[Names.GROUP] = person.GroupPerson;
                rowPerson[Names.N_ID] = person.idCard;

                rowPerson[Names.DEPARTMENT] = person.Department;
                rowPerson[Names.EMPLOYEE_POSITION] = person.PositionInDepartment;
                rowPerson[Names.PLACE_EMPLOYEE] = person.City;

                rowPerson[Names.EMPLOYEE_SHIFT] = person.Shift;

                rowPerson[Names.TIME_REGISTRATION] = "0"; //must be "0"!!!!
                rowPerson[Names.DATE_REGISTRATION] = day;
                rowPerson[Names.DAY_OF_WEEK] = DayOfWeekRussian((DateTime.Parse(day)).DayOfWeek.ToString());
                rowPerson[Names.DESIRED_TIME_IN] = person.ControlInHHMM;
                rowPerson[Names.DESIRED_TIME_OUT] = person.ControlOutHHMM;
                rowPerson[Names.EMPLOYEE_ABSENCE] = "1";

                dtTarget.Rows.Add(rowPerson);//добавляем рабочий день в который  сотрудник не выходил на работу
                ProgressWork1Step();
            }
            dtTarget.AcceptChanges();

            foreach (DataRow row in dtTarget.Rows)
            {
                if (row[Names.CODE].ToString() == person.Code)
                {
                    row[Names.FIO] = person.fio;
                    row[Names.GROUP] = person.GroupPerson;
                    row[Names.N_ID] = person.idCard;
                    row[Names.DEPARTMENT] = person.Department;
                    row[Names.PLACE_EMPLOYEE] = person.City;
                    row[Names.EMPLOYEE_POSITION] = person.PositionInDepartment;
                    row[Names.EMPLOYEE_SHIFT] = person.Shift;
                }
            }
            dtTarget.AcceptChanges();

            foreach (DataRow row in dtTarget.Rows) // search whole table
            {
                foreach (string day in workSelectedDays)
                {
                    if (row[Names.DATE_REGISTRATION]?.ToString() == day && row[Names.CODE]?.ToString() == person.Code)
                    {
                        try
                        {
                            row[Names.EMPLOYEE_SHIFT_COMMENT] = outPerson.Find((x) => x.date == day && x.code == person.Code).id;
                            logger.Trace("GetPersonRegistrationFromServer, outPerson " + person.Code + ", outReason - " + row[Names.EMPLOYEE_SHIFT_COMMENT].ToString());
                        }
                        catch { }
                        break;
                    }
                }
            }
            dtTarget.AcceptChanges();

            ProgressWork1Step();

            rowPerson = null;
            namePoint = null; direction = null;
            cellData = new string[1];
        }

        private string DayOfWeekRussian(string dayEnglish) //return a day of week as the same short name in Russian
        {
            string result ;
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
        private DataTable LoadGroupMembersFromDbToDataTable(string namePointedGroup) // dtPeopleGroup //"Select * FROM PeopleGroup where GroupPerson like '" + _textBoxReturnText(textBoxGroup) + "';"
        {
            logger.Trace($"-= {nameof(LoadGroupMembersFromDbToDataTable)} =-");

            DataTable peopleGroup = dtPeople.Clone();
            DataRow dataRow;

            string query = "Select FIO, NAV, GroupPerson, ControllingHHMM, ControllingOUTHHMM, Shift, Comment, Department, PositionInDepartment, DepartmentId, City, Boss FROM PeopleGroup ";
            if (namePointedGroup.StartsWith(@"@"))
            { query += "where DepartmentId like '" + namePointedGroup.Remove(0, 1) + "'; "; }
            else if (namePointedGroup.Length > 0)
            { query += "where GroupPerson like '" + namePointedGroup + "'; "; }
            else { query += ";"; }

            using (var sqlConnection = new SQLiteConnection(sqLiteLocalConnectionString))
            {
                sqlConnection.Open();
                using (var sqlCommand = new SQLiteCommand(query, sqlConnection))
                {
                    using (var sqlReader = sqlCommand.ExecuteReader())
                    {
                        foreach (DbDataRecord record in sqlReader)
                        {
                            if (record[@"FIO"]?.ToString()?.Length > 0 && record[@"NAV"]?.ToString()?.Length > 0)
                            {
                                dataRow = peopleGroup.NewRow();

                                dataRow[Names.FIO] = record[@"FIO"].ToString();
                                dataRow[Names.CODE] = record[@"NAV"].ToString();

                                dataRow[Names.GROUP] = record[@"GroupPerson"]?.ToString();
                                dataRow[Names.DEPARTMENT] = record[@"Department"]?.ToString();
                                dataRow[Names.DEPARTMENT_ID] = record[@"DepartmentId"]?.ToString();
                                dataRow[Names.EMPLOYEE_POSITION] = record[@"PositionInDepartment"]?.ToString();
                                dataRow[Names.PLACE_EMPLOYEE] = record[@"City"]?.ToString();
                                dataRow[Names.EMPLOYEE_SHIFT_COMMENT] = record[@"Comment"]?.ToString();
                                dataRow[Names.EMPLOYEE_SHIFT] = record[@"Shift"]?.ToString();

                                dataRow[Names.DESIRED_TIME_IN] = record[@"ControllingHHMM"]?.ToString();
                                dataRow[Names.DESIRED_TIME_OUT] = record[@"ControllingOUTHHMM"]?.ToString();

                                peopleGroup.Rows.Add(dataRow);
                            }
                        }
                    }
                }
            }
            query = null; dataRow = null;
            logger.Trace("LoadGroupMembersFromDbToDataTable, всего записей - " + peopleGroup.Rows.Count + ", группа - " + namePointedGroup);

            return peopleGroup;
        }

        private void infoItem_Click(object sender, EventArgs e)
        {
            ShowDataTableDbQuery(dbApplication, "TechnicalInfo",
                "SELECT PCName AS 'Версия Windows', POName AS 'Путь к ПО', POVersion AS 'Версия ПО', " +
                "LastDateStarted AS 'Дата использования', CurrentUser, FreeRam, GuidAppication ",
                "ORDER BY LastDateStarted DESC");
        }

        private void EnterEditAnualItem_Click(object sender, EventArgs e) //Select - EnterEditAnual() or ExitEditAnual()
        {
            if (EditAnualDaysItem.Text.Contains(Names.DAY_OFF_OR_WORK))
            {
                AddAnualDateItem.Font = new Font(this.Font, FontStyle.Bold);
                EditAnualDaysItem.Font = new Font(this.Font, FontStyle.Bold);
                EnterEditAnual();
            }
            else if (EditAnualDaysItem.Text.Contains(Names.END_EDIT))
            {
                ExitEditAnual();
                EditAnualDaysItem.Font = new Font(this.Font, FontStyle.Regular);
                AddAnualDateItem.Font = new Font(this.Font, FontStyle.Regular);
            }
        }

        private void EnterEditAnual()
        {
            logger.Trace($"-= {nameof(EnterEditAnual)} =-");

            ShowDataTableDbQuery(dbApplication, "BoldedDates",
                "SELECT DayBolded AS '" + Names.DAYOFF_DATE + "', DayType AS '" + Names.DAYOFF_TYPE + "', " +
                "NAV AS '" + Names.DAYOFF_USED_BY + "', DayDescription AS 'Описание', DateCreated AS '" + Names.DAYOFF_ADDED + "'",
                " ORDER BY DayBolded desc, NAV asc; ");

            EnableMenuItem(FunctionMenuItem, false);
            EnableMenuItem(GroupsMenuItem, false);
            EnableMenuItem(AddAnualDateItem, true);
            SetCheckBoxesAllFilters_Visible(false);

            comboBoxFio.Items.Add("");
            comboBoxFio.SelectedIndex = 0;

            toolTip1.SetToolTip(textBoxGroup, "Тип дня - 'Выходной' или 'Рабочий'");
            toolTip1.SetToolTip(textBoxGroupDescription, "Описание дня");
            labelGroup.Text = "Тип";
            textBoxGroup.Text = "";

            StatusLabel2.ForeColor = Color.Crimson;
            EditAnualDaysItem.Text = Names.END_EDIT;

            EditAnualDaysItem.ToolTipText = @"Выйти из режима редактирования рабочих и выходных дней";
            SetStatusLabelText(StatusLabel2, @"Режим редактирования рабочих и выходных дней");
        }

        private void ExitEditAnual()
        {
            logger.Trace($"-= {nameof(ExitEditAnual)} =-");

            comboBoxFio.Items?.RemoveAt(comboBoxFio.FindStringExact(""));
            if (comboBoxFio.Items.Count > 0)
            { comboBoxFio.SelectedIndex = 0; }

            EnableMenuItem(FunctionMenuItem, true);
            EnableMenuItem(GroupsMenuItem, true);
            EnableMenuItem(AddAnualDateItem, false);

            SetCheckBoxesAllFilters_Visible(true);

            EditAnualDaysItem.Text = Names.DAY_OFF_OR_WORK;
            EditAnualDaysItem.ToolTipText = Names.DAY_OFF_OR_WORK_EDIT;

            toolTip1.SetToolTip(textBoxGroup, "Создать или добавить в группу");
            toolTip1.SetToolTip(textBoxGroupDescription, "Изменить описание группы");
            labelGroup.Text = Names.GROUP;
            textBoxGroup.Text = "";

            SetStatusLabelForeColor(StatusLabel2, Color.Black);
            SetStatusLabelText(StatusLabel2, @"Завершен 'Режим редактирования дат праздников и выходных в локальной БД'");

            nameOfLastTable = "ListFIO";
            SeekAndShowMembersOfGroup("");
        }

        private void AddAnualDateItem_Click(object sender, EventArgs e) //AddAnualDate()
        {
            AddAnualDate();
            ShowDataTableDbQuery(dbApplication, "BoldedDates", "SELECT DayBolded AS '" + Names.DAYOFF_DATE + "', DayType AS '" + Names.DAYOFF_TYPE + "', " +
            "NAV AS '" + Names.DAYOFF_USED_BY + "', DayDescription AS 'Описание', DateCreated AS '" + Names.DAYOFF_ADDED + "'",
            " ORDER BY DayBolded desc, NAV asc; ");
        }

        private void AddAnualDate()
        {
            logger.Trace($"-= {nameof(AddAnualDate)} =-");

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

            using (var sqlConnection = new SQLiteConnection(sqLiteLocalConnectionString))
            {
                sqlConnection.Open();

                using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'BoldedDates' (DayBolded, NAV, DayType, DayDescription, DateCreated) " +
                    " VALUES (@BoldedDate, @NAV, @DayType, @DayDescription, @DateCreated)", sqlConnection))
                {
                    sqlCommand.Parameters.Add("@BoldedDate", DbType.String).Value = monthCalendar.SelectionStart.ToYYYYMMDD();
                    sqlCommand.Parameters.Add("@NAV", DbType.String).Value = nav;
                    sqlCommand.Parameters.Add("@DayType", DbType.String).Value = dayType;
                    sqlCommand.Parameters.Add("@DayDescription", DbType.String).Value = textBoxGroupDescription.Text.Trim();
                    sqlCommand.Parameters.Add("@DateCreated", DbType.String).Value = DateTime.Now.ToYYYYMMDD();
                    try { sqlCommand.ExecuteNonQuery(); } catch (Exception err) { MessageBox.Show(err.ToString()); }
                }
            }
        }

        private void DeleteAnualDateItem_Click(object sender, EventArgs e) //DeleteAnualDay()
        { DeleteAnualDay(); }

        private void DeleteAnualDay()
        {
            logger.Trace($"-= {nameof(DeleteAnualDay)} =-");

            string[] cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                 Names.DAYOFF_DATE, Names.DAYOFF_TYPE, Names.DAYOFF_USED_BY, Names.DAYOFF_ADDED });

            DeleteDataTableQueryParameters(dbApplication, "BoldedDates",
                "DayBolded", cellValue[0], "DayType", cellValue[1],
                "NAV", cellValue[2], "DateCreated", cellValue[3]).GetAwaiter().GetResult();

            ShowDataTableDbQuery(dbApplication, "BoldedDates", "SELECT DayBolded AS '" + Names.DAYOFF_DATE + "', DayType AS '" + Names.DAYOFF_TYPE + "', " +
            "NAV AS '" + Names.DAYOFF_USED_BY + "', DayDescription AS 'Описание', DateCreated AS '" + Names.DAYOFF_ADDED + "'",
            " ORDER BY DayBolded desc, NAV asc; ");
        }

        private void CheckBoxesFiltersAll_SetState(bool state)
        {
            SetCheckBoxCheckedStatus(checkBoxTimeViolations, state);
            SetCheckBoxCheckedStatus(checkBoxReEnter, state);
            SetCheckBoxCheckedStatus(checkBoxCelebrate, state);
            SetCheckBoxCheckedStatus(checkBoxWeekend, state);
        }

        private void CheckBoxesFiltersAll_Enable(bool state)
        {
            EnableControl(checkBoxTimeViolations, state);
            EnableControl(checkBoxReEnter, state);
            EnableControl(checkBoxCelebrate, state);
            EnableControl(checkBoxWeekend, state);
        }

        private void SetCheckBoxesAllFilters_Visible(bool state)
        {
            VisibleControl(checkBoxTimeViolations, state);
            VisibleControl(checkBoxReEnter, state);
            VisibleControl(checkBoxCelebrate, state);
            VisibleControl(checkBoxWeekend, state);
        }

        private async void checkBox_CheckStateChanged(object sender, EventArgs e)
        { await Task.Run(() => checkBoxCheckStateChanged()); }

        private void checkBoxCheckStateChanged()
        {
            logger.Trace($"-= {nameof(checkBoxCheckStateChanged)} =-");

            DataTable dtEmpty = new DataTable();
            EmployeeFull emptyPerson = new EmployeeFull();
            SeekAnualDays(ref dtEmpty, ref emptyPerson, false,
                ReturnDateTimePickerArray(dateTimePickerStart),
                ReturnDateTimePickerArray(dateTimePickerEnd),
                ref myBoldedDates, ref workSelectedDays);

            dtEmpty?.Dispose();

            CheckBoxesFiltersAll_Enable(false);
            VisibleControl(dataGridView1, false);

            string nameGroup = ReturnTextOfControl(textBoxGroup);

            //todo dubble
            // check need - DataTable dtTempIntermediate
            DataTable dtTempIntermediate = dtPeople.Clone();
            dtPersonTempAllColumns = dtPeople.Clone();
            EmployeeFull person = new EmployeeFull()
            {
                fio = ReturnTextOfControl(textBoxFIO),
                Code = ReturnTextOfControl(textBoxNav),
                GroupPerson = nameGroup,
                Department = nameGroup,
                ControlInSeconds = (int)(60 * 60 * numUpHourStart + 60 * numUpMinuteStart),
                ControlOutSeconds = (int)(60 * 60 * numUpHourEnd + 60 * numUpMinuteEnd),
                ControlInHHMM = ConvertDecimalTimeToStringHHMM(numUpHourStart, numUpMinuteStart),
                ControlOutHHMM = ConvertDecimalTimeToStringHHMM(numUpHourEnd, numUpMinuteEnd)
            };

            dtPersonTemp?.Clear();

            if ((nameOfLastTable == "PeopleGroupDescription" || nameOfLastTable == "PeopleGroup") && nameGroup.Length > 0)
            {
                dtPeopleGroup = LoadGroupMembersFromDbToDataTable(nameGroup);

                if (ReturnCheckboxCheckedStatus(checkBoxReEnter))
                {
                    foreach (DataRow row in dtPeopleGroup.Rows)
                    {
                        if (row[Names.FIO]?.ToString()?.Length > 0 && (row[Names.GROUP]?.ToString() == nameGroup || (@"@" + row[Names.DEPARTMENT_ID]?.ToString()) == nameGroup))
                        {
                            person = new EmployeeFull
                            {
                                fio = row[Names.FIO].ToString(),
                                Code = row[Names.CODE].ToString(),
                                GroupPerson = row[Names.GROUP].ToString(),
                                Department = row[Names.DEPARTMENT].ToString(),
                                PositionInDepartment = row[Names.EMPLOYEE_POSITION].ToString(),
                                City = row[Names.PLACE_EMPLOYEE]?.ToString(),
                                DepartmentId = row[Names.DEPARTMENT_ID].ToString(),
                                ControlInSeconds = row[Names.DESIRED_TIME_IN].ToString().ConvertTimeAsStringToSeconds(),
                                ControlOutSeconds = row[Names.DESIRED_TIME_OUT].ToString().ConvertTimeAsStringToSeconds(),
                                ControlInHHMM = row[Names.DESIRED_TIME_IN].ToString(),
                                ControlOutHHMM = row[Names.DESIRED_TIME_OUT].ToString(),
                                Comment = row[Names.EMPLOYEE_SHIFT_COMMENT].ToString(),
                                Shift = row[Names.EMPLOYEE_SHIFT].ToString()
                            };

                            FilterRegistrationsOfPerson(ref person, dtPersonRegistrationsFullList, ref dtTempIntermediate);
                        }
                    }
                }
                else
                { dtTempIntermediate = dtPersonRegistrationsFullList.Select("[Группа] = '" + nameGroup + "'").Distinct().CopyToDataTable(); }
            }
            else
            {
                if (!ReturnCheckboxCheckedStatus(checkBoxReEnter))
                { dtTempIntermediate = dtPersonRegistrationsFullList.Copy(); }
                else
                { FilterRegistrationsOfPerson(ref person, dtPersonRegistrationsFullList, ref dtTempIntermediate); }
            }

            //Table with all columns
            dtPersonTempAllColumns = dtTempIntermediate.Copy();
            dtPersonTemp = LeaveAndOrderColumnsOfDataTable(dtTempIntermediate, Names.orderColumnsRegistrations);

            //show selected data  within the selected collumns
            ShowDatatableOnDatagridview(dtPersonTemp, "PeopleGroup");

            //change enabling of checkboxes
            if (ReturnCheckboxCheckedStatus(checkBoxReEnter))// if (checkBoxReEnter.Checked)
            {
                EnableControl(checkBoxTimeViolations, true);
                EnableControl(checkBoxWeekend, true);
                EnableControl(checkBoxCelebrate, true);

                if (ReturnCheckboxCheckedStatus(checkBoxTimeViolations))  // if (checkBoxStartWorkInTime.Checked)
                { SetMenuItemBackColor(LoadDataItem, SystemColors.Control); }
            }
            else if (!ReturnCheckboxCheckedStatus(checkBoxReEnter))
            {
                SetCheckBoxCheckedStatus(checkBoxTimeViolations, false);
                SetCheckBoxCheckedStatus(checkBoxWeekend, false);
                SetCheckBoxCheckedStatus(checkBoxCelebrate, false);
                EnableControl(checkBoxTimeViolations, false);
                EnableControl(checkBoxWeekend, false);
                EnableControl(checkBoxCelebrate, false);
            }

            panelViewResize(numberPeopleInLoading);
            VisibleControl(dataGridView1, true);
            EnableControl(checkBoxReEnter, true);
        }

        private DataTable LeaveAndOrderColumnsOfDataTable(DataTable dt, string[] columns)
        {
            DataTable dtUniqRecords = dt.DefaultView.ToTable(true, columns);

            dtUniqRecords.SetColumnsOrder(columns);

            return dtUniqRecords;
        }

        private void FilterRegistrationsOfPerson(ref EmployeeFull person, DataTable dataTableSource, ref DataTable dataTableForStoring, string typeReport = "Полный")
        {
            logger.Trace($"-= {nameof(FilterRegistrationsOfPerson)} =-");
            logger.Trace("code: " + person.Code + "| dataTableSource: " + dataTableSource.Rows.Count, "| typeReport: " + typeReport);

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
                var allWorkedDaysPerson = dataTableSource.Select("[NAV-код] = '" + person.Code + "'");

                if (ReturnCheckboxCheckedStatus(checkBoxReEnter) || currentAction == "sendEmail") //checkBoxReEnter.Checked
                {
                    foreach (DataRow dataRowDate in allWorkedDaysPerson) //make the list of worked days
                    { hsDays.Add(dataRowDate[Names.DATE_REGISTRATION]?.ToString()); }

                    foreach (var workedDay in hsDays.ToArray())
                    {
                        //Select only one row with selected NAV for the selected workedDay
                        dtAllRegistrationsInSelectedDay = allWorkedDaysPerson.Distinct().CopyToDataTable().Select("[Дата регистрации] = '" + workedDay + "'").CopyToDataTable();

                        rowDtStoring = dtAllRegistrationsInSelectedDay.Select("[Дата регистрации] = '" + workedDay + "'").First();
                        rowDtStoring[Names.DAY_OF_WEEK] = DayOfWeekRussian(DateTime.Parse(workedDay).DayOfWeek.ToString());

                        //find first registration within the during selected workedDay
                        firstRegistrationInDay = Convert.ToInt32(dtAllRegistrationsInSelectedDay.Compute("MIN([Время регистрации])", string.Empty));

                        //find last registration within the during selected workedDay
                        lastRegistrationInDay = Convert.ToInt32(dtAllRegistrationsInSelectedDay.Compute("MAX([Время регистрации])", string.Empty));

                        //take and convert a real time coming into a string timearray
                        rowDtStoring[Names.TIME_REGISTRATION] = firstRegistrationInDay;              //("Время регистрации", typeof(decimal)), //15
                        rowDtStoring[Names.REAL_TIME_IN] = firstRegistrationInDay.ConvertSecondsToStringHHMMSS();  //("Фактич. время прихода ЧЧ:ММ:СС", typeof(string)),//24

                        // rowDtStoring[@"Реальное время ухода"] = lastRegistrationInDay;                 //("Реальное время ухода", typeof(decimal)), //18
                        rowDtStoring[Names.REAL_TIME_OUT] = lastRegistrationInDay.ConvertSecondsToStringHHMMSS();     //("Фактич. время ухода ЧЧ:ММ", typeof(string)), //25

                        //worked out times
                        workedSeconds = lastRegistrationInDay - firstRegistrationInDay;
                        rowDtStoring[Names.EMPLOYEE_TIME_SPENT] = workedSeconds;                                  // ("Реальное отработанное время", typeof(decimal)), //26
                        //  rowDtStoring[@"Отработанное время ЧЧ:ММ"] = ConvertSecondsToStringHHMMSS(workedSeconds);  //("Отработанное время ЧЧ:ММ", typeof(string)), //27
                        logger.Trace("FilterRegistrationsOfPerson: " + person.Code + "| " + rowDtStoring[Names.DATE_REGISTRATION]?.ToString() + " " + firstRegistrationInDay + " - " + lastRegistrationInDay);

                        //todo
                        //will calculate if day of week different
                        if (firstRegistrationInDay > (person.ControlInSeconds + offsetTimeIn) && firstRegistrationInDay != 0) // "Опоздание ЧЧ:ММ", typeof(bool)),           //28
                        {
                            if (typeReport == "Полный")
                            { rowDtStoring[Names.EMPLOYEE_BEING_LATE] = (firstRegistrationInDay - person.ControlInSeconds).ConvertSecondsToStringHHMMSS(); }
                            else if (typeReport == "Упрощенный")
                            { rowDtStoring[Names.EMPLOYEE_BEING_LATE] = "1"; }
                        }

                        if (lastRegistrationInDay < (person.ControlOutSeconds - offsetTimeOut) && lastRegistrationInDay != 0)  // "Ранний уход ЧЧ:ММ", typeof(bool)),                 //29
                        {
                            if (typeReport == "Полный")
                            { rowDtStoring[Names.EMPLOYEE_EARLY_DEPARTURE] = (person.ControlOutSeconds - lastRegistrationInDay).ConvertSecondsToStringHHMMSS(); }
                            else if (typeReport == "Упрощенный")
                            { rowDtStoring[Names.EMPLOYEE_EARLY_DEPARTURE] = "1"; }
                        }

                        if (rowDtStoring[Names.EMPLOYEE_ABSENCE]?.ToString() == "1" && typeReport == "Полный")  // "Ранний уход ЧЧ:ММ", typeof(bool)),                 //29
                        {
                            rowDtStoring[Names.EMPLOYEE_ABSENCE] = "Да";
                        }

                        exceptReason = rowDtStoring[Names.EMPLOYEE_SHIFT_COMMENT]?.ToString();

                        rowDtStoring[Names.EMPLOYEE_SHIFT_COMMENT] = outResons.Find((x) => x.id == exceptReason)?.nameReason;

                        switch (exceptReason)
                        {
                            case "2": //Отпуск
                            case "10": //Отпуск по беременности и родам
                            case "11": //Отпуск по уходу за ребёнком
                                rowDtStoring[Names.EMPLOYEE_EARLY_DEPARTURE] = "";
                                rowDtStoring[Names.EMPLOYEE_BEING_LATE] = "";
                                if (typeReport == "Полный")
                                { rowDtStoring[@"Отпуск"] = outResons.Find((x) => x.id == exceptReason)?.nameReason; }
                                else if (typeReport == "Упрощенный")
                                { rowDtStoring[@"Отпуск"] = "1"; }
                                break;

                            case "3": //Больничный
                            case "21": //Больничный 0,5 (менее < 5 часов)
                                rowDtStoring[Names.EMPLOYEE_EARLY_DEPARTURE] = "";
                                rowDtStoring[Names.EMPLOYEE_BEING_LATE] = "";
                                if (typeReport == "Полный")
                                { rowDtStoring[Names.EMPLOYEE_SICK_LEAVE] = outResons.Find((x) => x.id == exceptReason)?.nameReason; }
                                else if (typeReport == "Упрощенный")
                                { rowDtStoring[Names.EMPLOYEE_SICK_LEAVE] = "1"; }
                                break;

                            case "1": //Отпуск (за свой счёт)
                            case "9": //Прогул
                            case "12": //Отпуск по уходу за ребёнком возрастом до 3-х лет
                            case "20": //Отпуск за свой счет 0,5 (менее < 5 часов)
                                rowDtStoring[Names.EMPLOYEE_EARLY_DEPARTURE] = "";
                                rowDtStoring[Names.EMPLOYEE_BEING_LATE] = "";
                                if (typeReport == "Полный")
                                { rowDtStoring[Names.EMPLOYEE_HOOKY] = outResons.Find((x) => x.id == exceptReason)?.nameReason; }
                                else if (typeReport == "Упрощенный")
                                { rowDtStoring[Names.EMPLOYEE_HOOKY] = "1"; }
                                break;

                            case "4": //Командировка
                            case "5": //Выходной
                            case "6": //На выезде (по работе)
                            case "7": //Забыл пропуск
                            case "13": //Отгул (отпросился)
                            case "14": //Индивидуальный график
                            case "15": //Индивидуальный график
                            case "18": //Согласованное отсутствие (менее < 3 часов)
                            case "19": //Согласованное отсутствие (менее < 5 часов)

                                rowDtStoring[Names.EMPLOYEE_SHIFT_COMMENT] = outResons.Find((x) => x.id == exceptReason)?.nameReason;
                                rowDtStoring[Names.EMPLOYEE_EARLY_DEPARTURE] = "";
                                rowDtStoring[Names.EMPLOYEE_BEING_LATE] = "";
                                break;

                            default:
                                break;
                        }

                        dtTemp.ImportRow(rowDtStoring);
                    }
                }
                else if (!ReturnCheckboxCheckedStatus(checkBoxReEnter))
                {
                    foreach (DataRow dr in allWorkedDaysPerson)
                    { dtTemp.ImportRow(dr); }
                }

                if (ReturnCheckboxCheckedStatus(checkBoxWeekend) || currentAction == "sendEmail")//checkBoxWeekend Checking
                {
                    SeekAnualDays(ref dtTemp, ref person, true,
                        reportStartDay.ConvertDateAsStringToIntArray(),
                        reportLastDay.ConvertDateAsStringToIntArray(),
                        ref myBoldedDates, ref workSelectedDays);
                }

                if (ReturnCheckboxCheckedStatus(checkBoxTimeViolations)) //checkBoxStartWorkInTime Checking
                { QueryDeleteDataFromDataTable(ref dtTemp, "[Опоздание ЧЧ:ММ]='' AND [Ранний уход ЧЧ:ММ]=''", person.Code); }

                foreach (DataRow dr in dtTemp.AsEnumerable())
                { dataTableForStoring.ImportRow(dr); }

                allWorkedDaysPerson = null;
            }
            catch (Exception err) { MessageBox.Show(err.ToString()); }

            hsDays = null;
            rowDtStoring = null;
            dtTemp?.Dispose();
            exceptReason = null;
            dtAllRegistrationsInSelectedDay?.Dispose();
        }

        private void SeekAnualDays(ref DataTable dt, ref EmployeeFull person, bool delRow, int[] startOfPeriod, int[] endOfPeriod, ref string[] boldedDays, ref string[] workDays)//   //Exclude Anual Days from the table "PersonTemp" DB
        {
            logger.Trace($"-= {nameof(SeekAnualDays)} =-");

            if (person == null)
            { person = new EmployeeFull(); }
            if (person.Code == null)
            { person.Code = "0"; }

            List<string> daysBolded = new List<string>();
            List<string> daysListBolded = new List<string>();
            List<string> daysListWorkedInDB = new List<string>();

            var oneDay = TimeSpan.FromDays(1);
            var twoDays = TimeSpan.FromDays(2);

            var mySelectedStartDay = new DateTime(startOfPeriod[0], startOfPeriod[1], startOfPeriod[2]);
            var mySelectedEndDay = new DateTime(endOfPeriod[0], endOfPeriod[1], endOfPeriod[2]);
            var myMonthCalendar = new MonthCalendar
            {
                MaxSelectionCount = 60,
                SelectionRange = new SelectionRange(mySelectedStartDay, mySelectedEndDay)
            };

            //whole range of days in the selection period
            List<string> wholeSelectedDays = new List<string>();
            for (var myDate = myMonthCalendar.SelectionStart; myDate <= myMonthCalendar.SelectionEnd; myDate += oneDay)
            {
                wholeSelectedDays.Add(myDate.ToYYYYMMDD());
            }
            wholeSelectedDays.Sort();

            logger.Trace("SeekAnualDays,start-end: " + person.Code + " - " +
                startOfPeriod[0] + " " + startOfPeriod[1] + " " + startOfPeriod[2] + " - " +
                endOfPeriod[0] + " " + endOfPeriod[1] + " " + endOfPeriod[2]);

            logger.Trace("SeekAnualDays, wholeSelectedDays amount:" + wholeSelectedDays.Count);

            for (int year = -1; year < 1; year++)
            {
                daysListBolded.Add(new DateTime(startOfPeriod[0] + year, 1, 1).ToYYYYMMDD());
                daysListBolded.Add(new DateTime(startOfPeriod[0] + year, 1, 2).ToYYYYMMDD());
                daysListBolded.Add(new DateTime(startOfPeriod[0] + year, 1, 7).ToYYYYMMDD());
                daysListBolded.Add(new DateTime(startOfPeriod[0] + year, 3, 8).ToYYYYMMDD());
                daysListBolded.Add(new DateTime(startOfPeriod[0] + year, 5, 1).ToYYYYMMDD());
                daysListBolded.Add(new DateTime(startOfPeriod[0] + year, 5, 2).ToYYYYMMDD());
                daysListBolded.Add(new DateTime(startOfPeriod[0] + year, 5, 9).ToYYYYMMDD());
                daysListBolded.Add(new DateTime(startOfPeriod[0] + year, 6, 28).ToYYYYMMDD());
                daysListBolded.Add(new DateTime(startOfPeriod[0] + year, 8, 24).ToYYYYMMDD());
                daysListBolded.Add(new DateTime(startOfPeriod[0] + year, 10, 16).ToYYYYMMDD());
            }

            //Easter - Православная Пасха
            int Y = startOfPeriod[0];
            int a = Y % 19;
            int b = Y % 4;
            int c = Y % 7;
            int d = (19 * a + 15) % 30;
            int e = (2 * b + 4 * c + 6 * d + 6) % 7;
            int f = d + e;
            int monthWhenWillBeEaster = 4;
            int dayWhenWillBeEaster = 23;

            if (d == 29 && e == 6)
            {
                monthWhenWillBeEaster = 4;
                dayWhenWillBeEaster = 19;
            }
            else if (d == 28 && e == 6)
            {
                monthWhenWillBeEaster = 4;
                dayWhenWillBeEaster = 18;
            }
            else if (f <= 9)
            {
                monthWhenWillBeEaster = 3;
                dayWhenWillBeEaster = 22 + f;
            }
            else
            {
                monthWhenWillBeEaster = 4;
                dayWhenWillBeEaster = f - 9;
            }
            DateTime dayOff = new DateTime(startOfPeriod[0], monthWhenWillBeEaster, dayWhenWillBeEaster).AddDays(13);
            logger.Trace("SeekAnualDays, AddDayOff Easter ady: " + dayOff.ToYYYYMMDD());

            switch ((int)dayOff.DayOfWeek)
            {
                case (int)Day.Sunday:
                    daysListBolded.Add(dayOff.AddDays(1).ToYYYYMMDD());
                    logger.Trace("SeekAnualDays,AddDayOff Next day after Easter day: " + dayOff.AddDays(1).ToYYYYMMDD());
                    break;

                case (int)Day.Saturday:
                    daysListBolded.Add(dayOff.AddDays(2).ToYYYYMMDD());
                    logger.Trace("SeekAnualDays,AddDayOff Next 2nd day afer Easter day: " + dayOff.AddDays(2).ToYYYYMMDD());
                    break;

                default:
                    break;
            }

            DateTime dayBolded;

            List<string> days = ReturnBoldedDaysFromDB(person.Code, @"Выходной");
            if (days?.Count > 0)
            {
                foreach (string myDate in days) //days of looking for is 'Выходной' // or - Рабочий
                {
                    daysListBolded.Add(DateTime.Parse(myDate).ToYYYYMMDD());
                    logger.Trace("SeekAnualDays,AddBoldedDate from DB: " + myDate);
                }
            }

            //Independence day
            dayBolded = new DateTime(startOfPeriod[0], 8, 24);
            daysListBolded.Add(dayBolded.ToYYYYMMDD());
            logger.Trace("SeekAnualDays,AddBoldedDate Independence day: " + dayBolded.ToYYYYMMDD());

            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Sunday:
                    daysListBolded.Add(dayBolded.AddDays(1).ToYYYYMMDD());
                    logger.Trace("SeekAnualDays,AddBoldedDate Independence day: " + dayBolded.AddDays(1).ToYYYYMMDD());
                    break;

                case (int)Day.Saturday:
                    daysListBolded.Add(dayBolded.AddDays(2).ToYYYYMMDD());
                    logger.Trace("SeekAnualDays,AddBoldedDate Independence day: " + dayBolded.AddDays(2).ToYYYYMMDD());
                    break;

                default:
                    break;
            }

            //day of Ukraine Force
            dayBolded = new DateTime(startOfPeriod[0], 10, 16);
            daysListBolded.Add(dayBolded.ToYYYYMMDD());
            logger.Trace("SeekAnualDays,AddBoldedDate day of Ukraine Force: " + dayBolded.ToYYYYMMDD());

            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Sunday:
                    daysListBolded.Add(dayBolded.AddDays(1).ToYYYYMMDD());
                    logger.Trace("SeekAnualDays,AddBoldedDate day of Ukraine Force: " + dayBolded.AddDays(1).ToYYYYMMDD());
                    break;

                case (int)Day.Saturday:
                    daysListBolded.Add(dayBolded.AddDays(2).ToYYYYMMDD());
                    logger.Trace("SeekAnualDays,AddBoldedDate day of Ukraine Force: " + dayBolded.AddDays(2).ToYYYYMMDD());
                    break;

                default:
                    break;
            }

            //add all weekends to bolded days
            for (var myDate = myMonthCalendar.SelectionStart; myDate <= myMonthCalendar.SelectionEnd; myDate += oneDay)
            {
                if (myDate.DayOfWeek == DayOfWeek.Saturday || myDate.DayOfWeek == DayOfWeek.Sunday)
                {
                    daysListBolded.Add(myDate.ToYYYYMMDD());
                    logger.Trace("SeekAnualDays,AddBoldedDate weekends: " + myDate.ToYYYYMMDD());
                }
            }
            daysListBolded.Sort();
            logger.Trace("SeekAnualDays, daysListBolded:" + daysListBolded.ToArray().Length);

            //Add works days in List 'daysListWorkedInDB'
            days = ReturnBoldedDaysFromDB(person.Code, @"Рабочий");
            if (days?.Count > 0)
            {
                foreach (string myDate in days) //days of looking for is 'Рабочий'
                {
                    daysListWorkedInDB.Add(myDate);
                    logger.Trace("SeekAnualDays, Removed worked day from bolded: " + myDate);
                }
            }
            daysListWorkedInDB.Sort();
            logger.Trace("SeekAnualDays, daysListWorkedInDB:" + daysListWorkedInDB.ToArray().Length);

            List<string> tmpWholeSelectedDays = wholeSelectedDays;
            List<string> tmpDaysListBolded = daysListBolded;

            //except from list the  worked days of the local db
            tmpDaysListBolded.Except(daysListWorkedInDB);

            //get only worked days within the selected period
            List<string> tmp = tmpWholeSelectedDays.Except(tmpDaysListBolded).ToList();

            workDays = tmp.ToArray();

            logger.Trace("SeekAnualDays, amount worked days:" + workDays.Length);
            foreach (string str in workDays)
            { logger.Trace("SeekAnualDays, worked day: " + str); }

            boldedDays = wholeSelectedDays.Except(tmp).ToArray();
            logger.Trace("SeekAnualDays, amount bolded days:" + boldedDays.Length);
            foreach (string str in boldedDays)
            { logger.Trace("SeekAnualDays, bolded day: " + str); }

            foreach (var day in boldedDays)
            {
                if (delRow && dt != null)
                {
                    logger.Trace("SeekAnualDays, QueryDeleteDataFromDataTable: " + day);
                    QueryDeleteDataFromDataTable(ref dt, "[Дата регистрации]='" + day + "'", person.Code); // ("Дата регистрации",typeof(string)),//12
                }
            }

            if (dt != null)
            { dt.AcceptChanges(); }

            daysBolded.Sort();

            if (person == null || person.Code == "0")
            {
                monthCalendar.RemoveAllBoldedDates();
                foreach (string day in boldedDays)
                {
                    monthCalendar.AddBoldedDate(DateTime.Parse(day));
                    logger.Trace("SeekAnualDays, Added BoldedDate: " + day);
                }
            }

            foreach (string day in boldedDays)
            { logger.Trace("SeekAnualDays, Result bolded day: " + day); }

            foreach (string day in workDays)
            { logger.Trace("SeekAnualDays, Result work day: " + day); }

            myMonthCalendar.Dispose();
        }

        private List<string> ReturnBoldedDaysFromDB(string nav, string dayType)
        {
            List<string> boldedDays = new List<string>();
            string query ;
            if (nav.Length == 6)
            {
                query = $"SELECT DayBolded FROM BoldedDates WHERE (NAV LIKE '{nav}' OR  NAV LIKE '0') AND DayType LIKE '{dayType}';";
            }
            else
            {
                query = $"SELECT DayBolded FROM BoldedDates WHERE (NAV LIKE '0') AND DayType LIKE '{dayType}';";
            }

            using (SqLiteDbWrapper dbReader = new SqLiteDbWrapper(sqLiteLocalConnectionString, dbApplication))
            {
                System.Data.SQLite.SQLiteDataReader data = null;
                try
                {
                    data = dbReader?.GetData(query);
                }
                catch { logger.Info("ReturnBoldedDaysFromDB: no any info"); }

                if (data != null)
                {
                    foreach (DbDataRecord record in data)
                    {
                        if (record["DayBolded"]?.ToString()?.Length > 0)
                        {
                            boldedDays.Add(record["DayBolded"].ToString());
                        }
                    }
                }
                logger.Trace($"ReturnBoldedDaysFromDB: query: {query}\n{boldedDays.Count} rows loaded");
            }

            return boldedDays;
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
            }
            catch (Exception err)
            { logger.Warn(err.ToString()); }

            dt.AcceptChanges();
        }

        /*
        private List<string> SeekEasterDays( int[] startOfPeriod, int[] endOfPeriod)
        {
            List<string> daysListBolded = new List<string>();

            // Алгоритм для вычисления католической Пасхи http://snippets.dzone.com/posts/show/765
            int Y = startOfPeriod[0];
            int a = Y % 19;
            int aPr= Y % 19;
            int b = Y / 100;
            int bPr = Y % 4;
            int c = Y % 100;
            int cPr = Y % 7;
            int d = b / 4;
            int dPr = (19 * aPr + 15) % 30;
            int e = b % 4;
            int ePr = (2 * bPr + 4 * cPr + 6 * dPr + 6) % 7;
            int f = (b + 8) / 25;
            int fPr = dPr + ePr;
            int g = (b - f + 1) / 3;
            int h = (19 * a + b - d - g + 15) % 30;
            int i = c / 4;
            int k = c % 4;
            int L = (32 + 2 * e + 2 * i - h - k) % 7;
            int m = (a + 11 * h + 22 * L) / 451;
            int monthEaster = (h + L - 7 * m + 114) / 31;
            int dayEaster = ((h + L - 7 * m + 114) % 31) + 1;

            int monthEasterPr = 1;
            int dayEasterPr = 1;

            if (dPr == 29 && ePr == 6)
            {
                dayEasterPr = 19;
                 monthEasterPr= 4;
            }
           else if (dPr == 28 && ePr == 6)
            {
                dayEasterPr = 18;
                 monthEasterPr= 4;
            }
           else if (fPr <= 9)
            {
               dayEasterPr  = 22+fPr;
                 monthEasterPr= 3;
            }
            else
            {
               dayEasterPr  = fPr-9;
                 monthEasterPr= 4;
            }
            //Easter - Paskha
            DateTime dayBolded = new DateTime(startOfPeriod[0], monthEasterPr, dayEasterPr).AddDays(13);
            logger.Trace("SeekAnualDays,AddBoldedDate Easter: " + dayBolded.ToYYYYMMDD());

            switch ((int)dayBolded.DayOfWeek)
            {
                case (int)Day.Sunday:
                    daysListBolded.Add(dayBolded.AddDays(1).ToYYYYMMDD());

                    logger.Trace("SeekAnualDays,AddBoldedDate EasterNext day: " + dayBolded.AddDays(1).ToYYYYMMDD());
                    break;

                case (int)Day.Saturday:
                    daysListBolded.Add(dayBolded.AddDays(2).ToYYYYMMDD());

                    logger.Trace("SeekAnualDays,AddBoldedDate EasterNextNext day: " + dayBolded.AddDays(2).ToYYYYMMDD());
                    break;

                default:
                    break;
            }

            return daysListBolded;
        }
        */

        private void MakeZip(string[] files, string fullNameZip)
        {
            logger.Trace($"-= {nameof(MakeZip)} =-");

            foreach (string dirPath in files)
            {
                if (dirPath.Contains(@"\"))
                {
                    try
                    {
                        System.IO.Directory.CreateDirectory(dirPath.Replace(dirPath, appFolderTempPath + @"\" + dirPath.Remove(dirPath.IndexOf('\\'))));
                    }
                    catch (Exception err)
                    {
                        logger.Trace($"{dirPath} - {err.ToString()}");
                    }
                }
            }
            foreach (var file in files)
            {
                try
                {
                    System.IO.File.Copy(file, System.IO.Path.Combine(appFolderTempPath, file), true);
                }
                catch (Exception err) { logger.Trace($"MakeZip,System.IO.File.Copy: {file} - {err.ToString()}"); }
            }
            System.IO.Compression.ZipFile.CreateFromDirectory(appFolderTempPath, localAppFolderPath + @"\" + fullNameZip, System.IO.Compression.CompressionLevel.Optimal, false);
            logger.Info($"Архив создан: {localAppFolderPath}\\{fullNameZip}");
        }

        private void MakeZip(string filePath, string fullNameZip)
        {
            logger.Trace($"-= {nameof(MakeZip)} =-");

            if (filePath.Contains(@"\"))
            {
                try { System.IO.Directory.CreateDirectory(filePath.Replace(filePath, appFolderTempPath + @"\" + filePath.Remove(filePath.IndexOf('\\')))); }
                catch (Exception err) { logger.Trace(filePath + " - " + err.ToString()); }
            }
            try
            {
                System.IO.File.Copy(filePath, appFolderTempPath + @"\" + filePath, true);
            }
            catch (Exception err) { logger.Trace($"{filePath} - {err.ToString()}"); }

            System.IO.Compression.ZipFile.CreateFromDirectory(appFolderTempPath, localAppFolderPath + @"\" + fullNameZip, System.IO.Compression.CompressionLevel.Optimal, false);
            logger.Info($"Made archive: {localAppFolderPath}\\{fullNameZip}");
        }

        //----- Clearing. Start ---------//
        private void ClearRegistryItem_Click(object sender, EventArgs e) //ClearRegistryData()
        { ClearData(); }

        private void ClearData()
        {
            try
            {
                using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(appRegistryKey))
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
            }
            catch (Exception err)
            {
                SetStatusLabelText(StatusLabel2,
                    "Ошибки с доступом у реестру на запись. Данные не удалены.",
                    true);
                AddLoggerTraceText(err.ToString());
            }

            GC.Collect();
            VacuumDB(sqLiteLocalConnectionString);

            ClearItemsInFolder(@"*.xlsx");
            ClearItemsInFolder(@"*.log");
        }

        private void ClearItemsInFolder(string maskFiles)
        {
            if (System.IO.File.Exists(maskFiles))
            {
                try
                {
                    logger.Trace($"maskFiles: {maskFiles}");
                    System.IO.File.Delete(maskFiles);
                    logger.Trace($"Удален файл: {maskFiles}");
                }
                catch (Exception e) { logger.Warn($"Файл '{maskFiles}'не удален из-за ошибки: {e.Message}"); }
            }
            else if (System.IO.Directory.Exists(maskFiles))
            {
                var dir = new System.IO.DirectoryInfo(maskFiles);

                try
                {
                    logger.Trace($"dir.FullName: {dir.FullName}");
                    dir.Delete(true);
                    logger.Trace($"Папка удалена: {maskFiles}");
                }
                catch (Exception e) { logger.Warn($"Папка не удалена: {maskFiles} {e.Message}"); }

                try
                {
                    System.Threading.Thread.Sleep(200); //for prevent error of creating directory immediately after delete of one
                    logger.Trace($"dir.FullName: {dir.FullName}");
                    dir.Create();
                    logger.Trace($"Папка создана: {maskFiles}");
                }
                catch (Exception e) { logger.Warn($"Папка не создана: {maskFiles} {e.Message}"); }
            }
            else
            {
                var dir = System.IO.Path.GetDirectoryName(maskFiles);//get path of dir

                System.IO.FileInfo[] filesPath = new System.IO.DirectoryInfo(dir).GetFiles(
maskFiles.Remove(0, maskFiles.LastIndexOf(@"\") + 1),
System.IO.SearchOption.AllDirectories); //get files from dir
                if (filesPath?.Length > 0)
                {
                    foreach (var file in filesPath)
                    {
                        try
                        {
                            logger.Trace($"file: {file.FullName}");
                            file.Delete();
                            logger.Trace($"Удален файл: {file.FullName}");
                        }
                        catch (Exception e) { logger.Warn($"Файл '{file.Name}'не удален из-за ошибки: {e.Message}"); }
                    }
                }
            }
        }

        private void VacuumDB(string dbPath)
        {
            SQLiteConnectionStringBuilder builder =                new SQLiteConnectionStringBuilder
                {
                    DataSource = dbPath,
                    PageSize = 4096,
                    UseUTF16Encoding = true
                };

            using (SQLiteConnection conn = new SQLiteConnection(builder.ConnectionString))
            {
                conn.Open();

                using (SQLiteCommand vacuum = new SQLiteCommand(@"VACUUM", conn))
                { vacuum.ExecuteNonQuery(); }
            }
        }

        //----- Clearing. End ---------//

        //gathering a person's features from textboxes and other controls
        private void SelectPersonFromControls(ref EmployeeFull personSelected)
        {
            personSelected.fio = ReturnTextOfControl(textBoxFIO);
            personSelected.Code = ReturnTextOfControl(textBoxNav);
            personSelected.GroupPerson = ReturnTextOfControl(textBoxGroup);

            personSelected.ControlInHHMM = ConvertDecimalTimeToStringHHMM(ReturnNumUpDown(numUpDownHourStart), ReturnNumUpDown(numUpDownMinuteStart));
            personSelected.ControlOutHHMM = ConvertDecimalTimeToStringHHMM(ReturnNumUpDown(numUpDownHourEnd), ReturnNumUpDown(numUpDownMinuteEnd));
        }

        //---- Start. Drawing ---//
        private void VisualItem_Click(object sender, EventArgs e) //FindWorkDaysInSelected() , DrawFullWorkedPeriodRegistration()
        {
            logger.Trace($"-= {nameof(VisualItem_Click)} =-");

            EmployeeFull personVisual = new EmployeeFull();

            try
            {
                string[] cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                    Names.FIO,
                    Names.CODE,
                    Names.GROUP,
                    Names.DESIRED_TIME_IN,
                    Names.DESIRED_TIME_OUT,
                    Names.DEPARTMENT,
                    Names.EMPLOYEE_POSITION,
                    Names.EMPLOYEE_SHIFT,
                    Names.DEPARTMENT_ID,
                    Names.EMPLOYEE_SHIFT_COMMENT
                });

                if (DataGridViewOperations.RowsCount(dataGridView1) > -1)
                {
                    personVisual.fio = cellValue[0];
                    personVisual.Code = cellValue[1]; //Take the name of selected group
                    personVisual.ControlInHHMM = cellValue[3]; //Take the name of selected group

                    decimal[] timeIn = ConvertStringTimeHHMMToDecimalArray(personVisual.ControlInHHMM);
                    personVisual.ControlInSeconds = (int)timeIn[4];

                    personVisual.ControlOutHHMM = cellValue[4]; //Take the name of selected group
                    decimal[] timeOut = ConvertStringTimeHHMMToDecimalArray(personVisual.ControlOutHHMM);

                    personVisual.ControlOutSeconds = (int)timeOut[4];
                    SetNumUpDown(numUpDownHourStart, timeIn[0]);
                    SetNumUpDown(numUpDownMinuteStart, timeIn[1]);

                    SetNumUpDown(numUpDownHourEnd, timeOut[0]);
                    SetNumUpDown(numUpDownMinuteEnd, timeOut[1]);

                    personVisual.Department = cellValue[5];
                    personVisual.PositionInDepartment = cellValue[6];
                    personVisual.Shift = cellValue[7];
                    personVisual.DepartmentId = cellValue[8];
                    personVisual.Comment = cellValue[9];

                    if (nameOfLastTable == "PeopleGroup" || nameOfLastTable == "ListFIO")
                    {
                        personVisual.GroupPerson = cellValue[2]; //Take the name of selected group
                        StatusLabel2.Text = $"Выбрана группа: {personVisual.GroupPerson} | Курсор на: {personVisual.fio}";
                        groupBoxPeriod.BackColor = Color.PaleGreen;
                    }
                    else if (nameOfLastTable == "PersonRegistrationsList")
                    {
                        personVisual.GroupPerson = ReturnTextOfControl(textBoxGroup);
                        StatusLabel2.Text = $"Выбран: {personVisual.fio}";
                    }
                }
            }
            catch (Exception err) { logger.Warn($"VisualItem_Click: {err.ToString()}"); }

            if (personVisual.fio.Length == 0)
            {
                SelectPersonFromControls(ref personVisual);
            }

            VisibleControl(dataGridView1, false);

            CheckBoxesFiltersAll_Enable(false);

            if (ReturnCheckboxCheckedStatus(checkBoxReEnter))
            {
                logger.Trace("DrawFullWorkedPeriodRegistration: ");
                DrawFullWorkedPeriodRegistration(ref personVisual);
            }
            else
            {
                logger.Trace("DrawRegistration: ");
                DrawRegistration(ref personVisual);
            }

            VisibleMenuItem(TableModeItem, true);
            VisibleMenuItem(VisualModeItem, false);
            VisibleMenuItem(ChangeColorMenuItem, true);
            VisibleMenuItem(TableExportToExcelItem, false);

            VisibleControl(panelView, true);
            VisibleControl(pictureBox1, true);
        }

        private void DrawRegistration(ref EmployeeFull personDraw)  // Visualisation of registration
        {
            logger.Trace($"-= {nameof(DrawRegistration)} =-");

            int pointDrawYfor_rects = 44; //начальное смещение линии рабочего графика
            int pointDrawYfor_rectsReal = 39; // начальное смещение линии отработанного графика

            int iWidthRects = 2; // ширина прямоугольников = время нахождение в рабочей зоне(минимальное)

            int iLenghtRect = 0; //количество  входов-выходов в рабочие дни для всех отобранных людей для  анализа регистраций входа-выхода

            int minutesIn = 0;     // время входа в минутах планируемое
            int minutesInFact = 0;     // время выхода в минутах фактическое
            int minutesOut = 0;    // время входа в минутах планируемое
            int minutesOutFact = 0;    // время выхода в минутах фактическое

            //constant for a person
            string fio = personDraw.fio;
            string nav = personDraw.Code;
            string group = personDraw.GroupPerson;
            string dayRegistration = "";
            string directionPass = ""; //string pointName = "";

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
            { hsNAV.Add(row[Names.CODE].ToString()); }
            string[] arrayNAVs = hsNAV.ToArray(); //unique NAV-codes
            int countNAVs = arrayNAVs.Count(); //the number of selected people
            numberPeopleInLoading = countNAVs;
            logger.Trace("DrawRegistration,countNAVs: " + countNAVs);

            //count max number of events in-out all of selected people (the group or a single person)
            //It needs to prevent the error "index of scope"
            DataTable dtEmpty = new DataTable();
            SeekAnualDays(ref dtEmpty, ref personDraw, false,
                ReturnDateTimePickerArray(dateTimePickerStart),
                ReturnDateTimePickerArray(dateTimePickerEnd),
                ref myBoldedDates, ref workSelectedDays);
            dtEmpty.Dispose();

            foreach (DataRow row in rowsPersonRegistrationsForDraw.Rows)
            {
                for (int k = 0; k < workSelectedDays.Length; k++)
                {
                    if (workSelectedDays[k].Length == 10 && row[Names.DATE_REGISTRATION].ToString().Contains(workSelectedDays[k]))  //если приход пришел в рабочий день - считаем
                    { iLenghtRect++; }
                }
            }
            logger.Trace("DrawRegistration,iLenghtRect: " + iLenghtRect);

            panelView.Height = iShiftHeightAll + iOffsetBetweenHorizontalLines * workSelectedDays.Count() * countNAVs;
            // panelView.AutoScroll = false;
            // panelView.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            // panelView.Anchor = AnchorStyles.Bottom;
            // panelView.Anchor = AnchorStyles.Left;
            // panelView.Dock = DockStyle.None;
            ResumePpanel(panelView);

            pictureBox1?.Dispose();
            if (panelView.Controls.Count > 1)
            { panelView.Controls.RemoveAt(1); }

            pictureBox1 = new PictureBox
            {
                //    Location = new Point(0, 0),
                SizeMode = PictureBoxSizeMode.AutoSize,
                Size = new Size(
                    iShiftStart + iNumbersOfHoursInDay * iOffsetBetweenVerticalLines,
                    iShiftHeightAll + iOffsetBetweenHorizontalLines * workSelectedDays.Count() * countNAVs
                // (iShiftStart + iNumbersOfHoursInDay * iOffsetBetweenVerticalLines + 2) / 2 // 1740 на 870 - 24 часа и 43 строчки
                // 2 * (iShiftStart + iNumbersOfHoursInDay * iOffsetBetweenVerticalLines + 2) / 5  //1740 на 696 - 24 часа и 34 строчки
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
                using (var myBrushWorkHour = new SolidBrush(Color.Gray))
                {
                    using (var myBrushRealWorkHour = new SolidBrush(clrRealRegistration))
                    {
                        using (var myBrushAxis = new SolidBrush(Color.Black))
                        {
                            var pointForN_A = new PointF(0, iOffsetBetweenHorizontalLines);
                            var pointForN_B = new PointF(200, iOffsetBetweenHorizontalLines);

                            var axis = new Pen(Color.Black);

                            Rectangle[] rectsReal = new Rectangle[iLenghtRect]; //количество пересечений
                            Rectangle[] rectsRealMark = new Rectangle[iLenghtRect];
                            Rectangle[] rects = new Rectangle[workSelectedDays.Length * countNAVs];

                            int irectsTempReal = 0;

                            int numberRectangle_rectsRealMark = 0;
                            int numberRectangle_rects = 0;

                            foreach (string singleNav in arrayNAVs)
                            {
                                logger.Trace("DrawRegistration,draw: " + singleNav);

                                foreach (string workDay in workSelectedDays)
                                {
                                    foreach (DataRow row in rowsPersonRegistrationsForDraw.Rows)
                                    {
                                        nav = row[Names.CODE].ToString();
                                        dayRegistration = row[Names.DATE_REGISTRATION].ToString();

                                        if (singleNav.Contains(nav) && dayRegistration.Contains(workDay))
                                        {
                                            fio = row[Names.FIO].ToString();
                                            minutesIn = (int)ConvertStringTimeHHMMToDecimalArray(row[Names.DESIRED_TIME_IN].ToString())[3];
                                            minutesInFact = (int)ConvertStringTimeHHMMToDecimalArray(row[Names.REAL_TIME_IN].ToString())[3];
                                            minutesOut = (int)ConvertStringTimeHHMMToDecimalArray(row[Names.DESIRED_TIME_OUT].ToString())[3];
                                            minutesOutFact = (int)ConvertStringTimeHHMMToDecimalArray(row[Names.REAL_TIME_OUT].ToString())[3];
                                            directionPass = row[Names.CHECKPOINT_DIRECTION].ToString().ToLower();

                                            //pass by a point
                                            rectsRealMark[numberRectangle_rectsRealMark] = new Rectangle(
                                            iShiftStart + minutesInFact,             /* X */
                                            pointDrawYfor_rectsReal,                 /* Y */
                                            iWidthRects,                             /* width */
                                            iHeightLineRealWork                      /* height */
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
                                       iHeightLineWork                              /* height */
                                       );

                                    pointDrawYfor_rectsReal += iOffsetBetweenHorizontalLines;
                                    pointDrawYfor_rects += iOffsetBetweenHorizontalLines;
                                    numberRectangle_rects++;
                                }

                                //place the current FIO and days in visualisation
                                foreach (string workDay in workSelectedDays)
                                {
                                    pointForN_A.Y += iOffsetBetweenHorizontalLines;
                                    gr.DrawString(
                                        workDay + " (" + fio?.ConvertFullNameToShortForm() + ")",
                                        font,
                                        myBrushAxis,
                                        pointForN_A); //Paint workdays and people' FIO
                                }
                            }

                            //Fill with rectangles RealWork
                            gr.FillRectangles(myBrushRealWorkHour, rectsReal);

                            //Fill All Mark at Passthrow Points
                            gr.FillRectangles(myBrushRealWorkHour, rectsRealMark); //draw the real first come of the person

                            // Fill rectangles WorkTime shit
                            gr.FillRectangles(myBrushWorkHour, rects);

                            //Draw axes for days
                            for (int k = 0; k < workSelectedDays.Length * countNAVs; k++)
                            {
                                pointForN_B.Y += iOffsetBetweenHorizontalLines;
                                gr.DrawLine(
                                    axis,
                                    new Point(0, iShiftHeightAll + k * iOffsetBetweenHorizontalLines),
                                    new Point(pictureBox1.Width, iShiftHeightAll + k * iOffsetBetweenHorizontalLines));
                            }

                            //Draw other axes
                            gr.DrawString(
                                "Время, часы:",
                                font,
                                SystemBrushes.WindowText,
                                new Point(iShiftStart - 110, iOffsetBetweenHorizontalLines / 4));
                            gr.DrawString("Дата (ФИО)",
                                font,
                                SystemBrushes.WindowText,
                                new Point(10, iOffsetBetweenHorizontalLines));
                            gr.DrawLine(
                                axis, new Point(0, 0),
                                new Point(iShiftStart, iShiftHeightAll));
                            gr.DrawLine(
                                axis,
                                new Point(iShiftStart, 0),
                                new Point(iShiftStart, iShiftHeightAll));

                            for (int k = 0; k < iNumbersOfHoursInDay; k++)
                            {
                                gr.DrawLine(
                                    axis,
                                    new Point(iShiftStart + k * iOffsetBetweenVerticalLines, iShiftHeightAll),
                                    new Point(iShiftStart + k * iOffsetBetweenVerticalLines, Convert.ToInt32(pictureBox1.Height)));
                                gr.DrawString(
                                    Convert.ToString(k),
                                    font,
                                    SystemBrushes.WindowText,
                                    new Point(320 + k * iOffsetBetweenVerticalLines, iOffsetBetweenHorizontalLines));
                            }
                            gr.DrawLine(
                                axis,
                                new Point(iShiftStart + iNumbersOfHoursInDay * iOffsetBetweenVerticalLines, iShiftHeightAll),
                                new Point(iShiftStart + iNumbersOfHoursInDay * iOffsetBetweenVerticalLines, Convert.ToInt32(pictureBox1.Height)));

                            axis.Dispose();
                            rectsRealMark = null;
                            rectsReal = null;
                            rects = null;
                        }
                    }
                }
                sLastSelectedElement = "DrawRegistration";
            }
            //---------------------------------------------------------------//
            // End of the Block with Drawing //

            pictureBox1.Image = bmp;
            pictureBox1.Refresh();
            panelView.Controls.Add(pictureBox1);
            RefreshPictureBox(pictureBox1, bmp);
            panelViewResize(numberPeopleInLoading);

            font.Dispose();
        }

        private void DrawFullWorkedPeriodRegistration(ref EmployeeFull personDraw)  // Draw the whole period registration
        {
            logger.Trace($"-= {nameof(DrawFullWorkedPeriodRegistration)} =-");

            int pointDrawYfor_rects = 44; //начальное смещение линии рабочего графика
            int pointDrawYfor_rectsReal = 39; // начальное смещение линии отработанного графика
            int iLenghtRect = 0; //количество  входов-выходов в рабочие дни для всех отобранных людей для  анализа регистраций входа-выхода

            int minutesIn = 0;     // время входа в минутах планируемое
            int minutesInFact = 0;     // время выхода в минутах фактическое
            int minutesOut = 0;    // время входа в минутах планируемое
            int minutesOutFact = 0;    // время выхода в минутах фактическое

            //constant for a person
            string fio = personDraw.fio;
            string nav = personDraw.Code;
            string group = personDraw.GroupPerson;
            string dayRegistration = "";
            string directionPass = ""; //string pointName = "";

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
            { hsNAV.Add(row[Names.CODE].ToString()); }
            string[] arrayNAVs = hsNAV.ToArray(); //unique NAV-codes
            int countNAVs = arrayNAVs.Count(); //the number of selected people
            numberPeopleInLoading = countNAVs;

            //count max number of events in-out all of selected people (the group or a single person)
            //It needs to prevent the error "index of scope"
            DataTable dtEmpty = new DataTable();
            SeekAnualDays(ref dtEmpty, ref personDraw, false,
                ReturnDateTimePickerArray(dateTimePickerStart),
                ReturnDateTimePickerArray(dateTimePickerEnd),
                ref myBoldedDates, ref workSelectedDays);
            dtEmpty.Dispose();

            foreach (DataRow row in rowsPersonRegistrationsForDraw)
            {
                for (int k = 0; k < workSelectedDays.Length; k++)
                {
                    if (workSelectedDays[k].Length == 10 && row[Names.DATE_REGISTRATION].ToString().Contains(workSelectedDays[k]))  //если приход пришел в рабочий день - считаем
                    { iLenghtRect++; }
                }
            }

            panelView.Height = iShiftHeightAll + iOffsetBetweenHorizontalLines * workSelectedDays.Count() * countNAVs;
            // panelView.AutoScroll = false;
            // panelView.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            // panelView.Anchor = AnchorStyles.Bottom;
            // panelView.Anchor = AnchorStyles.Left;
            // panelView.Dock = DockStyle.None;
            ResumePpanel(panelView);

            pictureBox1?.Dispose();
            if (panelView.Controls.Count > 1)
            { panelView.Controls.RemoveAt(1); }

            pictureBox1 = new PictureBox
            {
                //Location = new Point(0, 0),
                SizeMode = PictureBoxSizeMode.AutoSize,
                Size = new Size(
                    iShiftStart + iNumbersOfHoursInDay * iOffsetBetweenVerticalLines,
                    iShiftHeightAll + iOffsetBetweenHorizontalLines * workSelectedDays.Count() * countNAVs
                // (iShiftStart + iNumbersOfHoursInDay * iOffsetBetweenVerticalLines + 2) / 2 // 1740 на 870 - 24 часа и 43 строчки
                // 2 * (iShiftStart + iNumbersOfHoursInDay * iOffsetBetweenVerticalLines + 2) / 5  //1740 на 696 - 24 часа и 34 строчки
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
                using (var myBrushWorkHour = new SolidBrush(Color.Gray))
                {
                    using (var myBrushRealWorkHour = new SolidBrush(clrRealRegistration))
                    {
                        using (var myBrushAxis = new SolidBrush(Color.Black))
                        {
                            var pointForN_A = new PointF(0, iOffsetBetweenHorizontalLines);
                            var pointForN_B = new PointF(200, iOffsetBetweenHorizontalLines);

                            var axis = new Pen(Color.Black);

                            Rectangle[] rectsRealMark = new Rectangle[iLenghtRect];
                            Rectangle[] rects = new Rectangle[workSelectedDays.Length * countNAVs];

                            int numberRectangle_rectsRealMark = 0;
                            int numberRectangle_rects = 0;

                            foreach (string singleNav in arrayNAVs)
                            {
                                foreach (string workDay in workSelectedDays)
                                {
                                    foreach (DataRow row in rowsPersonRegistrationsForDraw)
                                    {
                                        nav = row[Names.CODE].ToString();
                                        dayRegistration = row[Names.DATE_REGISTRATION].ToString();

                                        if (singleNav.Contains(nav) && dayRegistration.Contains(workDay))
                                        {
                                            fio = row[Names.FIO].ToString();
                                            minutesIn = (int)ConvertStringTimeHHMMToDecimalArray(row[Names.DESIRED_TIME_IN].ToString())[3];
                                            minutesInFact = (int)ConvertStringTimeHHMMToDecimalArray(row[Names.REAL_TIME_IN].ToString())[3];
                                            minutesOut = (int)ConvertStringTimeHHMMToDecimalArray(row[Names.DESIRED_TIME_OUT].ToString())[3];
                                            minutesOutFact = (int)ConvertStringTimeHHMMToDecimalArray(row[Names.REAL_TIME_OUT].ToString())[3];
                                            directionPass = row[Names.CHECKPOINT_DIRECTION].ToString().ToLower();

                                            //pass by a point
                                            rectsRealMark[numberRectangle_rectsRealMark] = new Rectangle(
                                            iShiftStart + minutesInFact,             /* X */
                                            pointDrawYfor_rectsReal,                 /* Y */
                                            minutesOutFact - minutesInFact,          /* width */
                                            iHeightLineRealWork                      /* height */
                                            );

                                            numberRectangle_rectsRealMark++;
                                        }
                                    }

                                    //work shift
                                    rects[numberRectangle_rects] = new Rectangle(
                                       iShiftStart + minutesIn,                     /* X */
                                       pointDrawYfor_rects,                         /* Y */
                                       minutesOut - minutesIn,                      /* width */
                                       iHeightLineWork                              /* height */
                                       );

                                    pointDrawYfor_rectsReal += iOffsetBetweenHorizontalLines;
                                    pointDrawYfor_rects += iOffsetBetweenHorizontalLines;
                                    numberRectangle_rects++;
                                }

                                //place the current FIO and days in visualisation
                                foreach (string workDay in workSelectedDays)
                                {
                                    pointForN_A.Y += iOffsetBetweenHorizontalLines;
                                    gr.DrawString(
                                        workDay + " (" + fio?.ConvertFullNameToShortForm() + ")",
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
                                pointForN_B.Y += iOffsetBetweenHorizontalLines;
                                gr.DrawLine(
                                    axis,
                                    new Point(0, iShiftHeightAll + k * iOffsetBetweenHorizontalLines),
                                    new Point(pictureBox1.Width, iShiftHeightAll + k * iOffsetBetweenHorizontalLines));
                            }

                            //Draw other axes
                            gr.DrawString(
                                "Время, часы:",
                                font,
                                SystemBrushes.WindowText,
                                new Point(iShiftStart - 110, iOffsetBetweenHorizontalLines / 4));
                            gr.DrawString("Дата (ФИО)",
                                font,
                                SystemBrushes.WindowText,
                                new Point(10, iOffsetBetweenHorizontalLines));
                            gr.DrawLine(
                                axis, new Point(0, 0),
                                new Point(iShiftStart, iShiftHeightAll));
                            gr.DrawLine(
                                axis,
                                new Point(iShiftStart, 0),
                                new Point(iShiftStart, iShiftHeightAll));

                            for (int k = 0; k < iNumbersOfHoursInDay; k++)
                            {
                                gr.DrawLine(
                                    axis,
                                    new Point(iShiftStart + k * iOffsetBetweenVerticalLines, iShiftHeightAll),
                                    new Point(iShiftStart + k * iOffsetBetweenVerticalLines, Convert.ToInt32(pictureBox1.Height)));
                                gr.DrawString(
                                    Convert.ToString(k),
                                    font,
                                    SystemBrushes.WindowText,
                                    new Point(320 + k * iOffsetBetweenVerticalLines, iOffsetBetweenHorizontalLines));
                            }
                            gr.DrawLine(
                                axis,
                                new Point(iShiftStart + iNumbersOfHoursInDay * iOffsetBetweenVerticalLines, iShiftHeightAll),
                                new Point(iShiftStart + iNumbersOfHoursInDay * iOffsetBetweenVerticalLines, Convert.ToInt32(pictureBox1.Height)));

                            axis.Dispose();
                            rects = null;
                            rectsRealMark = null;
                        }
                    }
                }
                sLastSelectedElement = "DrawFullWorkedPeriodRegistration";
            }
            //---------------------------------------------------------------//
            // End of the Block with Drawing //

            pictureBox1.Image = bmp;
            pictureBox1.Refresh();
            panelView.Controls.Add(pictureBox1);
            RefreshPictureBox(pictureBox1, bmp);
            panelViewResize(numberPeopleInLoading);

            font.Dispose();
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
            VisibleControl(pictureBox1, false);

            try
            {
                if (panelView?.Controls?.Count > 1) panelView.Controls.RemoveAt(1);
                bmp?.Dispose();
                pictureBox1?.Dispose();
            }
            catch { }

            VisibleControl(dataGridView1, true);

            sLastSelectedElement = "dataGridView";
            panelViewResize(numberPeopleInLoading);

            CheckBoxesFiltersAll_Enable(true);
            VisibleMenuItem(TableExportToExcelItem, true);
            VisibleMenuItem(TableModeItem, false);
            VisibleMenuItem(VisualModeItem, true);
            VisibleMenuItem(ChangeColorMenuItem, false);
        }

        private void panelView_SizeChanged(object sender, EventArgs e)
        { panelViewResize(numberPeopleInLoading); }

        private void panelViewResize(int numberPeople) //Change PanelView
        {
            int iOffsetBetweenHorizontalLines = 19;
            int iShiftHeightAll = 36;
            switch (sLastSelectedElement)
            {
                case "DrawFullWorkedPeriodRegistration":
                    SetPanelHeight(panelView, iShiftHeightAll + iOffsetBetweenHorizontalLines * workSelectedDays.Length * numberPeople); //Fixed size of Picture. If need autosize - disable this row
                    break;

                case "DrawRegistration":
                    SetPanelHeight(panelView, iShiftHeightAll + iOffsetBetweenHorizontalLines * workSelectedDays.Length * numberPeople); //Fixed size of Picture. If need autosize - disable this row
                    break;

                default:
                    SetPanelAnchor(panelView, (AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top));
                    SetPanelHeight(panelView, ReturnPanelParentHeight(panelView) - 120);
                    SetPanelAutoScroll(panelView, true);
                    SetPanelAutoSizeMode(panelView, AutoSizeMode.GrowAndShrink);
                    ResumePpanel(panelView);
                    break;
            }

            if (ReturnPanelControlsCount(panelView) > 1)
            {
                RefreshPictureBox(pictureBox1, bmp);
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

            using (var sqlConnection = new SQLiteConnection(sqLiteLocalConnectionString))
            {
                sqlConnection.Open();
                using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'ConfigDB' (ParameterName, Value, DateCreated) " +
                    " VALUES (@ParameterName, @Value, @DateCreated)", sqlConnection))
                {
                    sqlCommand.Parameters.Add("@ParameterName", DbType.String).Value = "clrRealRegistration";
                    sqlCommand.Parameters.Add("@Value", DbType.String).Value = clrRealRegistration.Name;
                    sqlCommand.Parameters.Add("@DateCreated", DbType.String).Value = DateTime.Now.ToYYYYMMDDHHMM();
                    try { sqlCommand.ExecuteNonQuery(); } catch (Exception err) { MessageBox.Show(err.ToString()); }
                }
            }

            logger.Trace("ColorizeDraw:clrRealRegistration.Name: " + clrRealRegistration.Name);
        }

        private void ColorizeDraw(Color color)
        {
            EmployeeFull personSelected = new EmployeeFull();
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
            MessageBox.Show(
               @"Перед получением информации необходимо в Настройках:\n\n" +
                @"1.1. Добавить URI адрес СКД-сервера с ПО 'Интеллект' (SERVER.DOMAIN.SUBDOMAIN),\nа также логин и пароль пользователя для доступа к SQL-серверу СКД\n" +
                @"1.2. Добавить URI адрес сервера с базой сотрудников корпоративного сайта вэб-сервера (SERVER.DOMAIN),\nnа также логин и пароль пользователя для доступа к MySQL-серверу\n" +
                @"1.3. Добавить URI адрес сервера с контроллером домена (SERVER.DOMAIN.SUBDOMAIN),\nа также логин, пароль и домен(DOMAIN.SUBDOMAIN) логина пользователя для доступа к данным\n" +
                @"1.4. Добавить URI адрес почтового сервера (SERVER.DOMAIN.SUBDOMAIN),\nа также email и (не обязательно) пароль пользователя для отправки рассылок с отчетами\n" +
                @"2. Сохранить введенные параметры\n" +
                @"2.1. В случае ввода некорректных данных получение данных и/или отправка рассылок будут заблокированы\n" +
                @"2.2. Если данные введены корректно (будут отсутствовать ошибки о вводе некорректных данных), необходимо перезапустить программу\n" +
                @"3. После этого можно:\n" +
                @"3.1. Получать списки сотрудников\n" +
                @"3.2. Создавать и использовать ранее сохраненные локально группы пользователей\n" +
                @"3.3. Добавлять или корректировать праздничные дни, отгулы, отпуски - персонального для каждого или для всех\n" +
                @"3.4. Загружать регистрации пропусков по группам или отдельным сотрудников, генерировать отчеты, " +
                @"визуализировать полученные данные, отправлять отчеты по спискам рассылок, сгенерированным автоматически или вручную\n" +
                @"3.5. Создавать собственные группы генерации отчетов, собственные рассылки отчетов\n" +
                @"3.6. Проводить анализ полученных данных как в табличном виде так и визуально, экспортировать данные в Excel файл." +
                @"3.7. Загружать все попытки регистрации пропусков, включая запрещенные попытки прохода, за текущий день или выбранный период." +
                @"\nФильтровать эти данные по пользователям или попыткам прохода" +
                @"4. ПО способно самостоятельно или принудительно проверять наличие обновления на сервере." +
                @"4.1. Для использования данного функционала заполните в конфигурации в параметр 'serverUpdateURL' URI адрес папки сервера с обновлениями (SERVER.DOMAIN.SUBDOMAIN\FOLDER_WITH_UPDATES)." +
                @"\n\nДата и время локального ПК: " + ReturnDateTimePicker(dateTimePickerEnd).ToYYYYMMDDHHMMSS(),

               "Информация о программе",
               MessageBoxButtons.OK,
               MessageBoxIcon.Information,
               MessageBoxDefaultButton.Button1);
        }

        private void PrepareForMakingFormMailing(object sender, EventArgs e) //MailingItem()
        {
            nameOfLastTable = "Mailing";
            EnableMenuItem(SettingsMenuItem, false);
            EnableMenuItem(FunctionMenuItem, false);
            EnableMenuItem(GroupsMenuItem, false);
            CheckBoxesFiltersAll_Enable(false);
            VisibleControl(panelView, false);

            btnPropertiesSave.Text = "Сохранить рассылку";
            ClearButtonClickEvent(btnPropertiesSave);
            btnPropertiesSave.Click += new EventHandler(ButtonPropertiesSave_MailingSave);

            MakeFormMailing();
        }

        private void MakeFormMailing()
        {
            List<string> periodComboParameters = new List<string>
            {
                "Предыдущий месяц",
                "Текущий месяц"
            };

            List<string> listComboParameters9 = new List<string>
            {
                "Активная",
                "Неактивная"
            };

            List<string> listComboParameters15 = new List<string>
            {
                "Полный",
                "Упрощенный"
            };

            List<string> listComboParameters = new List<string>
            {
                "Все"
            };

            //get list groups from DB and add to listComboParameters
            using (var sqlConnection = new SQLiteConnection(sqLiteLocalConnectionString))
            {
                sqlConnection.Open();

                using (var sqlCommand = new SQLiteCommand(
                    "SELECT GroupPerson FROM PeopleGroupDescription;", sqlConnection))
                {
                    using (var sqlReader = sqlCommand.ExecuteReader())
                    {
                        foreach (DbDataRecord record in sqlReader)
                        {
                            string group = record["GroupPerson"]?.ToString();
                            if (group?.Length > 0)
                            {
                                listComboParameters.Add(group);
                            }
                        }
                    }
                }
                sqlConnection.Close();
            }

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
                    "День подготовки отчета", "", "День, в который выполнять подготовку и отправку данного отчета./nНачало, Средина, Конец"
                    );
        }

        private void SaveMailing(string recipientEmail, string groupsReport, string nameReport, string descriptionReport,
            string periodPreparing, string status, string dateCreatingMailing, string SendingDate, string typeReport, string daySendingReport)
        {
            //  method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace(nameof(SaveMailing));

            //     bool recipientValid = false;
            //    bool senderValid = false;

            //     if (recipientEmail?.Length > 0 && recipientEmail.Contains('.') && recipientEmail.Contains('@') && recipientEmail?.Split('.').Count() > 1)
            //    { recipientValid = true; }

            //    if (senderEmail?.Length > 0 && senderEmail.Contains('.') && senderEmail.Contains('@') && senderEmail?.Split('.').Count() > 1)
            //   { senderValid = true; }

            using (SQLiteConnection sqlConnection = new SQLiteConnection(sqLiteLocalConnectionString))
            {
                sqlConnection.Open();

                SQLiteCommand sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                sqlCommand1.ExecuteNonQuery();

                using (SQLiteCommand sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'Mailing' (RecipientEmail, GroupsReport, NameReport, Description, Period, Status, DateCreated, SendingLastDate, TypeReport, DayReport)" +
                           " VALUES (@RecipientEmail, @GroupsReport, @NameReport, @Description, @Period, @Status, @DateCreated, @SendingLastDate, @TypeReport, @DayReport)", sqlConnection))
                {
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
            SetStatusLabelText(StatusLabel2, $"Добавлена рассылка: {nameReport}| Всего рассылок: {DataGridViewOperations.RowsCount(dataGridView1)}");
        }

        private void ConfigurationItem_Click(object sender, EventArgs e)
        {
            ShowDataTableDbQuery(dbApplication, "ConfigDB", "SELECT ParameterName AS 'Имя параметра', " +
            "Value AS 'Данные', Description AS 'Описание', DateCreated AS 'Дата создания/модификации'",
            " ORDER BY ParameterName asc, DateCreated desc; ");
        }

        private void SettingsProgrammItem_Click(object sender, EventArgs e)
        { MakeFormSettings(); }

        private void MakeFormSettings()
        {
            EnableMainMenuItems(false);

            VisibleControl(panelView, false);

            btnPropertiesSave.Text = "Сохранить настройки";
            ClearButtonClickEvent(btnPropertiesSave);
            btnPropertiesSave.Click += new EventHandler(buttonPropertiesSave_Click);
            ViewFormSettings(
                "Сервер СКД", sServer1, "Имя сервера \"Server\" с базой Intellect в виде - NameOfServer.Domain.Subdomain",
                "Имя пользователя", sServer1UserName, "Имя администратора \"sa\" SQL-сервера",
                "Пароль", sServer1UserPassword, "Пароль администратора \"sa\" SQL-сервера \"Server\"",
                "Почтовый сервер", mailServer, "Имя почтового сервера \"Mail Server\" в виде - NameOfServer.Domain.Subdomain",
                "e-mail пользователя", mailSenderAddress, "E-mail отправителя рассылок виде - User.Name@MailServer.Domain.Subdomain",
                "Пароль", mailsOfSenderOfPassword, "Пароль E-mail отправителя почты",
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
            //   method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace($"-= {nameof(ViewFormSettings)} =-");

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

                listboxPeriod = new ListBox
                {
                    Location = new Point(300, 121),
                    Size = new Size(120, 20),
                    Parent = groupBoxProperties
                };
                toolTip1.SetToolTip(listboxPeriod, tooltip8);

                listboxPeriod.DataSource = periodStrings8;
                listboxPeriod.KeyPress += new KeyPressEventHandler(PeriodCombo_KeyPress);
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

                textBoxSettings16.KeyPress += new KeyPressEventHandler(textboxDate_KeyPress);
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
            listboxPeriod?.BringToFront();

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
            logger.Trace($"-= {nameof(buttonPropertiesCancel_Click)} =-");

            string btnName = btnPropertiesSave.Text;

            DisposeTemporaryControls();

            if (btnName == @"Сохранить рассылку")
            {
                ShowDataTableDbQuery(dbApplication, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
                "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
                "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
                " ORDER BY RecipientEmail asc, DateCreated desc; ");
            }
            EnableMainMenuItems(true);
            VisibleControl(panelView, true);
        }

        private void DisposeTemporaryControls()
        {
            VisibleControl(groupBoxProperties, false);
            DisposeControl(labelServer1);
            DisposeControl(labelServer1UserName);
            DisposeControl(labelServer1UserPassword);
            DisposeControl(labelMailServerName);
            DisposeControl(labelMailServerUserName);
            DisposeControl(labelMailServerUserPassword);
            DisposeControl(labelmysqlServer);
            DisposeControl(labelmysqlServerUserName);
            DisposeControl(labelmysqlServerUserPassword);

            DisposeControl(textBoxServer1);
            DisposeControl(textBoxServer1UserName);
            DisposeControl(textBoxServer1UserPassword);
            DisposeControl(textBoxMailServerName);
            DisposeControl(textBoxMailServerUserName);
            DisposeControl(textBoxMailServerUserPassword);
            DisposeControl(textBoxmysqlServer);
            DisposeControl(textBoxmysqlServerUserName);
            DisposeControl(textBoxmysqlServerUserPassword);

            DisposeControl(listComboLabel);
            DisposeControl(periodComboLabel);
            DisposeControl(labelSettings9);

            DisposeControl(listCombo);
            DisposeControl(listboxPeriod);
            DisposeControl(comboSettings9);

            DisposeControl(labelSettings15);
            DisposeControl(comboSettings15);

            DisposeControl(labelSettings16);
            DisposeControl(textBoxSettings16);
            DisposeControl(checkBox1);
        }

        private void buttonPropertiesSave_Click(object sender, EventArgs e) //SaveProperties()
        {
            SaveProperties(); //btnPropertiesSave

            DisposeTemporaryControls();
            EnableMainMenuItems(true);
            VisibleControl(panelView, true);
        }

        private void ButtonPropertiesSave_MailingSave(object sender, EventArgs e)
        {
            logger.Trace($"-= {nameof(ButtonPropertiesSave_MailingSave)} =-");

            string recipientEmail = ReturnTextOfControl(textBoxServer1UserName);
            string senderEmail = mailSenderAddress;
            if (mailSenderAddress.Length == 0)
            { senderEmail = ReturnTextOfControl(textBoxServer1); }
            string nameReport = ReturnTextOfControl(textBoxMailServerName);
            string description = ReturnTextOfControl(textBoxMailServerUserName);
            string report = ReturnComboBoxSelectedItem(listCombo);
            string period = ReturnListBoxSelectedItem(listboxPeriod);
            string status = ReturnComboBoxSelectedItem(comboSettings9);
            string typeReport = ReturnComboBoxSelectedItem(comboSettings15);
            string dayReport = ReturnTextOfControl(textBoxSettings16);

            if (recipientEmail.Length > 5 && nameReport.Length > 0)
            {
                SaveMailing(recipientEmail, report, nameReport, description, period, status,
                    DateTime.Now.ToYYYYMMDDHHMM(), "", typeReport, dayReport);
            }

            ShowDataTableDbQuery(dbApplication, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
            "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
            "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
            " ORDER BY RecipientEmail asc, DateCreated desc; ");

            DisposeTemporaryControls();
            EnableMainMenuItems(true);
            VisibleControl(panelView, true);
        }

        private void SaveProperties() //Save Parameters into Registry and variables
        {
            logger.Trace($"-= {nameof(SaveProperties)} =-");

            string server = ReturnTextOfControl(textBoxServer1);
            string user = ReturnTextOfControl(textBoxServer1UserName);
            string password = ReturnTextOfControl(textBoxServer1UserPassword);

            string sMailServer = ReturnTextOfControl(textBoxMailServerName);
            string sMailUser = ReturnTextOfControl(textBoxMailServerUserName);
            string sMailUserPassword = ReturnTextOfControl(textBoxMailServerUserPassword);

            string sMySqlServer = ReturnTextOfControl(textBoxmysqlServer);
            string sMySqlServerUser = ReturnTextOfControl(textBoxmysqlServerUserName);
            string sMySqlServerUserPassword = ReturnTextOfControl(textBoxmysqlServerUserPassword);

            string checkInputedData = $"Data Source={server}\\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID={user};Password={password};Connect Timeout=5";

            CheckAliveIntellectServer(sServer1, sqlServerConnectionString, queryCheckSystemDdSCA);

            if (bServer1Exist)
            {
                VisibleControl(groupBoxProperties, false);
                EnableMenuItem(GetFioItem, true);

                sServer1 = server;
                sServer1UserName = user;
                sServer1UserPassword = password;
                sqlServerConnectionString = checkInputedData;
                sqlServerConnectionStringEF = new System.Data.Entity.Core.EntityClient.EntityConnectionStringBuilder()
                {
                    Metadata = @"res://*/DBFirstCatalogModel.csdl|res://*/DBFirstCatalogModel.ssdl|res://*/DBFirstCatalogModel.msl",
                    Provider = @"System.Data.SqlClient",
                    ProviderConnectionString = $"{checkInputedData};multipleactiveresultsets=True;App=EntityFramework"
                };

                mailServer = sMailServer;
                mailSenderAddress = sMailUser;
                mailsOfSenderOfPassword = sMailUserPassword;

                mysqlServer = sMySqlServer;
                mysqlServerUserName = sMySqlServerUser;
                mysqlServerUserPassword = sMySqlServerUserPassword;
                mysqlServerConnectionString = $"server={mysqlServer};User={mysqlServerUserName};Password={mysqlServerUserPassword};database=wwwais;convert zero datetime=True;Connect Timeout=60";

                // Save data in Registry
                //try
                //{
                //    using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(appRegistryKey))
                //    {
                //        try { EvUserKey.SetValue("SKDServer", sServer1, Microsoft.Win32.RegistryValueKind.String); } catch { }
                //        try { EvUserKey.SetValue("SKDUser", EncryptionDecryptionCriticalData.EncryptStringToBase64Text(sServer1UserName, keyEncryption, keyDencryption), Microsoft.Win32.RegistryValueKind.String); } catch { }
                //        try { EvUserKey.SetValue("SKDUserPassword", EncryptionDecryptionCriticalData.EncryptStringToBase64Text(sServer1UserPassword, keyEncryption, keyDencryption), Microsoft.Win32.RegistryValueKind.String); } catch { }

                //        try { EvUserKey.SetValue("MySQLServer", mysqlServer, Microsoft.Win32.RegistryValueKind.String); } catch { }
                //        try { EvUserKey.SetValue("MySQLUser", mysqlServerUserName, Microsoft.Win32.RegistryValueKind.String); } catch { }
                //        try { EvUserKey.SetValue("MySQLUserPassword", EncryptionDecryptionCriticalData.EncryptStringToBase64Text(mysqlServerUserPassword, keyEncryption, keyDencryption), Microsoft.Win32.RegistryValueKind.String); } catch { }

                //        logger.Info("CreateSubKey: Данные в реестре сохранены");
                //    }
                //}
                //catch (Exception err) { logger.Error("CreateSubKey: Ошибки с доступом на запись в реестр. Данные сохранены не корректно. " + err.ToString()); }

                string resultSaving = "";
                resultSaving += SaveParameterInConfigASTA(new ParameterConfig()
                {
                    name = "SKDServer",
                    value = sServer1,
                    description = "URI SKD-сервера",
                    isSecret = false,
                    isExample = "no"
                }) + "\n";

                resultSaving += SaveParameterInConfigASTA(new ParameterConfig()
                {
                    name = "SKDUser",
                    value = sServer1UserName,
                    description = "SKD MSSQL User's Name",
                    isSecret = true,
                    isExample = "no"
                }) + "\n";

                resultSaving += SaveParameterInConfigASTA(new ParameterConfig()
                {
                    name = "SKDUserPassword",
                    value = sServer1UserPassword,
                    description = "SKD MSSQL User's Password",
                    isSecret = true,
                    isExample = "no"
                }) + "\n";

                resultSaving += SaveParameterInConfigASTA(new ParameterConfig()
                {
                    name = "MailServer",
                    value = mailServer,
                    description = "URI Mail-серверa",
                    isSecret = false,
                    isExample = "no"
                }) + "\n";

                resultSaving += SaveParameterInConfigASTA(new ParameterConfig()
                {
                    name = "MailUser",
                    value = mailSenderAddress,
                    description = "Sender E-Mail's",
                    isSecret = false,
                    isExample = "no"
                }) + "\n";

                resultSaving += SaveParameterInConfigASTA(new ParameterConfig()
                {
                    name = "MailUserPassword",
                    value = mailsOfSenderOfPassword,
                    description = "Password of sender of e-mails",
                    isSecret = true,
                    isExample = "no"
                }) + "\n";

                resultSaving += SaveParameterInConfigASTA(new ParameterConfig()
                {
                    name = "MySQLServer",
                    value = mysqlServer,
                    description = "URL MySQL серверa (www)",
                    isSecret = false,
                    isExample = "no"
                }) + "\n";

                resultSaving += SaveParameterInConfigASTA(new ParameterConfig()
                {
                    name = "MySQLUser",
                    value = mysqlServerUserName,
                    description = "MySQL User login",
                    isSecret = false,
                    isExample = "no"
                }) + "\n";

                resultSaving += SaveParameterInConfigASTA(new ParameterConfig()
                {
                    name = "MySQLUserPassword",
                    value = mysqlServerUserPassword,
                    description = "Password of MySQL User",
                    isSecret = true,
                    isExample = "no"
                });

                DisposeTemporaryControls();
                VisibleControl(panelView, true);

                if (mailServer?.Length > 0 && mailServerSMTPPort > 0)
                { _mailServer = new MailServer(mailServer, mailServerSMTPPort); }
               
                if (mailSenderAddress != null && mailSenderAddress.Contains('@'))
                { _mailUser = new MailUser(NAME_OF_SENDER_REPORTS, mailSenderAddress); }

                ShowDataTableDbQuery(dbApplication, "ConfigDB", "SELECT ParameterName AS 'Имя параметра', " +
               "Value AS 'Данные', Description AS 'Описание', DateCreated AS 'Дата создания/модификации'",
               " ORDER BY ParameterName asc, DateCreated desc; ");

                Task.Run(() => MessageBox.Show(resultSaving));
                SetStatusLabelText(StatusLabel2, "Введенны корректные данные для подключения к серверу СКД");
                SetStatusLabelBackColor(StatusLabel2, Color.PaleGreen);
            }
            else
            {
                SetStatusLabelText(StatusLabel2, "Данные некорректные или же сервер СКД не доступен");
                SetStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                GetInfoSetup();
            }
        }
        //--- End. Features of programm ---//

        //--- Start. Behaviour Controls ---//
        private void PeriodCombo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)//если нажата Enter
            {
                listboxPeriod.Items.Add(listboxPeriod.Text);
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
                SetTextBoxFIOAndTextBoxNAVFromSelectedComboboxFio();
            }
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
            if (PersonOrGroupItem.Text == Names.WORK_WITH_A_PERSON)
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
            if (PersonOrGroupItem.Text == Names.WORK_WITH_A_PERSON)
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

        private void SetGroupNameInStatusLabel()
        {
            if (ReturnTextOfControl(textBoxGroupDescription)?.Length > 0)
            {
                SetStatusLabelText(
                    StatusLabel2,
                    $"Создать или добавить в группу: {ReturnTextOfControl(textBoxGroup)}({ReturnTextOfControl(textBoxGroupDescription)})");
            }
            else
            {
                SetStatusLabelText(
                    StatusLabel2,
                    $"Создать или добавить в группу: {ReturnTextOfControl(textBoxGroup)}");
            }
        }

        private void textBoxGroupDescription_TextChanged(object sender, EventArgs e)
        { SetGroupNameInStatusLabel(); }

        private void textBoxGroup_TextChanged(object sender, EventArgs e)
        {
            if (textBoxGroup.Text.Trim().Length > 0)
            {
                AddPersonToGroupItem.Enabled = true;
                CreateGroupItem.Enabled = true;

                SetGroupNameInStatusLabel();
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
            numUpHourStart = ReturnNumUpDown(numUpDownHourStart);
            numUpMinuteStart = ReturnNumUpDown(numUpDownMinuteStart);
            numUpHourEnd = ReturnNumUpDown(numUpDownHourEnd);
            numUpMinuteEnd = ReturnNumUpDown(numUpDownMinuteEnd);
        }

        private void dateTimePickerStart_CloseUp(object sender, EventArgs e)
        {
            LoadDataItem.Enabled = true;
            LoadDataItem.BackColor = Color.PaleGreen;

            SetMenuItemsTextAfterClosingDateTimePicker();
        }

        private void dateTimePickerEnd_CloseUp(object sender, EventArgs e)
        { SetMenuItemsTextAfterClosingDateTimePicker(); }

        private void SetMenuItemsTextAfterClosingDateTimePicker()
        {
            string dayStart = ReturnDateTimePicker(dateTimePickerStart).ToYYYYMMDD();
            string dayEnd = ReturnDateTimePicker(dateTimePickerEnd).ToYYYYMMDD();

            SetMenuItemText(LoadLastInputsOutputsItem, $"Загрузить все события СКД за сегодня ({ DateTime.Now.ToYYYYMMDD()})");

            if (dayStart != dayEnd)
                SetMenuItemText(LoadInputsOutputsItem, $"Загрузить все события СКД с {dayStart} по {dayEnd}");
            else
                SetMenuItemText(LoadInputsOutputsItem, $"Загрузить все события СКД за {dayStart}");
        }

        private void PersonOrGroupItem_Click(object sender, EventArgs e) //PersonOrGroup()
        { PersonOrGroup(); }

        private void PersonOrGroup()
        {
            string menu = ReturnMenuItemText(PersonOrGroupItem);
            switch (menu)
            {
                case (Names.WORK_WITH_A_GROUP):
                    SetMenuItemText(PersonOrGroupItem, Names.WORK_WITH_A_PERSON);
                    EnableControl(comboBoxFio, false);
                    nameOfLastTable = "PersonRegistrationsList";
                    break;

                case (Names.WORK_WITH_A_PERSON):
                    SetMenuItemText(PersonOrGroupItem, Names.WORK_WITH_A_GROUP);
                    EnableControl(comboBoxFio, true);
                    nameOfLastTable = "PeopleGroup";
                    break;

                default:
                    SetMenuItemText(PersonOrGroupItem, Names.WORK_WITH_A_GROUP);
                    EnableControl(comboBoxFio, true);
                    nameOfLastTable = "PeopleGroup";
                    break;
            }
        }

        //--- End. Behaviour Controls ---//

        //---  Start.  DatagridView functions ---//
        private void dataGridView1_DoubleClick(object sender, EventArgs e) //SearchMembersSelectedGroup()
        {
            if (nameOfLastTable == "PeopleGroup" ||
                nameOfLastTable == "PeopleGroupDescription" ||
                nameOfLastTable == "Mailing")
            { SearchMembersSelectedGroup(); }
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e) //DataGridView1CellEndEdit()
        { DataGridView1CellEndEdit(); }

        private void DataGridView1CellEndEdit()
        {
            logger.Trace($"-= {nameof(DataGridView1CellEndEdit)} =-");

            string fio = "";
            string nav = "";
            string group = "";
            int currRow = DataGridViewOperations.RowsCount(dataGridView1);

            if (currRow > -1)
            {
                try
                {
                    string currColumn = DataGridViewOperations.ColumnName(dataGridView1, DataGridViewOperations.CurrentColumnIndex(dataGridView1));
                    string currCellValue = dgvo.CurrentCellValue(dataGridView1);
                    string editedCell = "";

                    switch (nameOfLastTable)
                    {
                        case "BoldedDates":
                            {
                                string[] cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            Names.DAYOFF_DATE,
                            Names.DAYOFF_USED_BY,
                            Names.DAYOFF_TYPE });

                                string dayType = "";
                                if (textBoxGroup?.Text?.Trim()?.Length == 0 || textBoxGroup?.Text?.ToLower()?.Trim() == "выходной")
                                { dayType = "Выходной"; }
                                else { dayType = "Рабочий"; }

                                if (textBoxNav?.Text?.Trim()?.Length != 6)
                                { nav = "для всех"; }
                                else { nav = textBoxNav.Text.Trim(); }

                                string navD = "";
                                if (cellValue[1]?.Length != 6)
                                { navD = "всех"; }
                                else { navD = cellValue[1]; }

                                using (var sqlConnection = new SQLiteConnection(sqLiteLocalConnectionString))
                                {
                                    sqlConnection.Open();

                                    using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'BoldedDates' (DayBolded, NAV, DayType, DayDescription, DateCreated) " +
                                        " VALUES (@BoldedDate, @NAV, @DayType, @DayDescription, @DateCreated)", sqlConnection))
                                    {
                                        sqlCommand.Parameters.Add("@BoldedDate", DbType.String).Value = monthCalendar.SelectionStart.ToYYYYMMDD();
                                        sqlCommand.Parameters.Add("@NAV", DbType.String).Value = nav;
                                        sqlCommand.Parameters.Add("@DayType", DbType.String).Value = dayType;
                                        sqlCommand.Parameters.Add("@DayDescription", DbType.String).Value = textBoxGroupDescription.Text.Trim();
                                        sqlCommand.Parameters.Add("@DateCreated", DbType.String).Value = DateTime.Now.ToYYYYMMDD();
                                        try { sqlCommand.ExecuteNonQuery(); } catch { }
                                    }
                                }
                                break;
                            }
                        case "PeopleGroup":
                        case "ListFIO":
                            {
                                string[] cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            Names.FIO,
                            Names.CODE,
                            Names.GROUP,
                            Names.DESIRED_TIME_IN,
                            Names.DESIRED_TIME_OUT,
                            Names.DEPARTMENT,
                            Names.EMPLOYEE_POSITION,
                            Names.EMPLOYEE_SHIFT,
                            Names.EMPLOYEE_SHIFT_COMMENT,
                            Names.DEPARTMENT_ID
                            });

                                fio = cellValue[0];
                                textBoxFIO.Text = fio;

                                nav = cellValue[1];
                                textBoxNav.Text = nav;

                                group = cellValue[2];
                                textBoxGroup.Text = group;

                                //todo
                                //Change to UPDATE, but not Replace!!!!
                                /*
                                 myCommand1.CommandText = "UPDATE dept SET dname = :dname, loc = :loc WHERE deptno = @deptno";
                                myCommand1.Parameters.Add("@deptno", 20);
                                myCommand1.Parameters.Add("dname", "SALES");
                                myCommand1.Parameters.Add("loc", "NEW YORK");
                                */
                                using (var connection = new SQLiteConnection(sqLiteLocalConnectionString))
                                {
                                    connection.Open();
                                    using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'PeopleGroup' (FIO, NAV, GroupPerson, ControllingHHMM, ControllingOUTHHMM, Shift, Comment, Department, PositionInDepartment, DepartmentId, City) " +
                                                                                                " VALUES (@FIO, @NAV, @GroupPerson, @ControllingHHMM, @ControllingOUTHHMM, @Shift, @Comment, @Department, @PositionInDepartment, @DepartmentId, @City)", connection))
                                    {
                                        sqlCommand.Parameters.Add("@FIO", DbType.String).Value = fio;
                                        sqlCommand.Parameters.Add("@NAV", DbType.String).Value = nav;

                                        sqlCommand.Parameters.Add("@GroupPerson", DbType.String).Value = group;
                                        sqlCommand.Parameters.Add("@Department", DbType.String).Value = cellValue[5];
                                        sqlCommand.Parameters.Add("@PositionInDepartment", DbType.String).Value = cellValue[6];
                                        sqlCommand.Parameters.Add("@DepartmentId", DbType.String).Value = cellValue[9];

                                        sqlCommand.Parameters.Add("@ControllingHHMM", DbType.String).Value = cellValue[3];
                                        sqlCommand.Parameters.Add("@ControllingOUTHHMM", DbType.String).Value = cellValue[4];

                                        sqlCommand.Parameters.Add("@Shift", DbType.String).Value = cellValue[7];
                                        sqlCommand.Parameters.Add("@Comment", DbType.String).Value = cellValue[8];

                                        try { sqlCommand.ExecuteNonQuery(); } catch { }
                                    }
                                }
                                //  nameOfLastTable = "PeopleGroup";
                                SeekAndShowMembersOfGroup(group);
                                SetStatusLabelText(StatusLabel2, $"Обновлено время прихода {fio?.ConvertFullNameToShortForm()} в группе: {group}");
                                break;
                            }
                        case "PeopleGroupDescription":
                            {
                                string[] cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            Names.GROUP,
                            Names.GROUP_DECRIPTION });

                                textBoxGroup.Text = cellValue[0]; //Take the name of selected group
                                textBoxGroupDescription.Text = cellValue[1]; //Take the name of selected group
                                groupBoxPeriod.BackColor = Color.PaleGreen;
                                groupBoxFilterReport.BackColor = SystemColors.Control;
                                SetStatusLabelText(StatusLabel2, $"Выбрана группа: {cellValue[0]}");
                                break;
                            }
                        case "Mailing":
                            {
                                string[] cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            @"Получатель", @"Отчет по группам", @"Наименование", @"Описание",
                            @"Период", @"Статус", @"Тип отчета", @"День отправки отчета" });

                                switch (currColumn)
                                {
                                    case "День отправки отчета":
                                        editedCell = ReturnStrongNameDayOfSendingReports(cellValue[7]);
                                        ExecuteSqlAsync("UPDATE 'Mailing' SET DayReport='" + editedCell + "' WHERE RecipientEmail='" + cellValue[0]
                                          + "' AND NameReport='" + cellValue[2] + "' AND GroupsReport ='" + cellValue[1]
                                          + $"' AND Period='{cellValue[4]}' AND TypeReport ='{cellValue[6]}' AND Status ='{cellValue[5]}' AND Description ='{cellValue[3]}';");
                                        break;

                                    case "Тип отчета":
                                        if (currCellValue == "Полный") { editedCell = "Полный"; }
                                        else { editedCell = "Упрощенный"; }

                                        ExecuteSqlAsync($"UPDATE 'Mailing' SET TypeReport='{editedCell}' WHERE RecipientEmail='" + cellValue[0]
                                          + $"' AND NameReport='{cellValue[2]}' AND GroupsReport ='" + cellValue[1]
                                          + $"' AND Period='{cellValue[4]}' AND DayReport='{cellValue[7]}' AND Status ='{cellValue[5]}' AND Description ='{cellValue[3]}';");
                                        break;

                                    case "Статус":
                                        if (currCellValue.ToLower().Contains("неактивн")) { editedCell = "Неактивная"; }
                                        else { editedCell = "Активная"; }

                                        ExecuteSqlAsync($"UPDATE 'Mailing' SET Status='{editedCell}' WHERE RecipientEmail='" + cellValue[0]
                                          + $"' AND NameReport='{cellValue[2]}' AND GroupsReport ='" + cellValue[1]
                                          + $"' AND Period='{cellValue[4]}' AND DayReport='{cellValue[7]}' AND TypeReport ='{cellValue[6]}' AND Description ='{cellValue[3]}';");
                                        break;

                                    case "Период":
                                        if (currCellValue.ToLower().Contains("текущ")) { editedCell = "Текущий месяц"; }
                                        else if (currCellValue.ToLower().Contains("предыд")) { editedCell = "Предыдущий месяц"; }
                                        else { editedCell = currCellValue; }

                                        ExecuteSqlAsync($"UPDATE 'Mailing' SET Period='{editedCell}' WHERE RecipientEmail='" + cellValue[0]
                                           + $"' AND NameReport='{cellValue[2]}' AND GroupsReport ='" + cellValue[1]
                                           + $"' AND TypeReport ='{cellValue[6]}' AND DayReport='{cellValue[7]}' AND Status ='{cellValue[5]}' AND Description ='{cellValue[3]}';");
                                        break;

                                    case "Описание":
                                        editedCell = currCellValue;

                                        ExecuteSqlAsync("UPDATE 'Mailing' SET Description='" + editedCell + "' WHERE RecipientEmail='" + cellValue[0]
                                          + "' AND NameReport='" + cellValue[2] + "' AND GroupsReport ='" + cellValue[1]
                                          + "' AND TypeReport ='" + cellValue[6] + "' AND DayReport='" + cellValue[7]
                                          + "' AND Status ='" + cellValue[5] + "' AND Period='" + cellValue[4] + "';");
                                        break;

                                    case "Отчет по группам":
                                        editedCell = currCellValue;

                                        ExecuteSqlAsync("UPDATE 'Mailing' SET GroupsReport ='" + editedCell + "' WHERE RecipientEmail='" + cellValue[0]
                                          + "' AND NameReport='" + cellValue[2] + "' AND Description ='" + cellValue[3]
                                          + "' AND Status ='" + cellValue[5] + "' AND Period='" + cellValue[4]
                                          + "' AND TypeReport ='" + cellValue[6] + "' AND DayReport='" + cellValue[7] + "';");
                                        break;

                                    case "Получатель":
                                        if (currCellValue.Contains('@') && currCellValue.Contains('.'))
                                        { editedCell = currCellValue; }
                                        else
                                        { editedCell = mailJobReportsOfNameOfReceiver; }

                                        ExecuteSqlAsync("UPDATE 'Mailing' SET RecipientEmail ='" + editedCell + "' WHERE TypeReport ='" + cellValue[6]
                                              + "' AND NameReport='" + cellValue[2] + "' AND GroupsReport ='" + cellValue[1]
                                              + "' AND DayReport='" + cellValue[7] + "' AND Period='" + cellValue[4]
                                              + "' AND Status ='" + cellValue[5] + "' AND Description ='" + cellValue[3] + "';");

                                        break;

                                    case "Наименование":
                                        editedCell = currCellValue;

                                        ExecuteSqlAsync("UPDATE 'Mailing' SET NameReport ='" + editedCell + "' WHERE RecipientEmail='" + cellValue[0]
                                          + "' AND Description='" + cellValue[3] + "' AND GroupsReport ='" + cellValue[1]
                                          + "' AND DayReport='" + cellValue[7] + "' AND TypeReport ='" + cellValue[6]
                                          + "' AND Period ='" + cellValue[4] + "' AND Status ='" + cellValue[5] + "';");
                                        break;

                                    default:
                                        break;
                                }

                                ShowDataTableDbQuery(dbApplication, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
                                "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
                                "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
                                " ORDER BY RecipientEmail asc, DateCreated desc; ");
                                break;
                            }
                        case "MailingException":
                            {
                                string[] cellValue =
                                    dgvo.FindValuesInCurrentRow(dataGridView1, new string[] { @"Получатель", @"Описание" });

                                switch (currColumn)
                                {
                                    case "Получатель":
                                        ExecuteSqlAsync("UPDATE 'MailingException' SET RecipientEmail='" + currCellValue +
                                            "' WHERE Description='" + cellValue[1] + "';");
                                        break;

                                    case "Описание":
                                        ExecuteSqlAsync("UPDATE 'MailingException' SET Description='" + currCellValue +
                                            "' WHERE RecipientEmail='" + cellValue[0] + "';");
                                        break;

                                    default:
                                        break;
                                }

                                ShowDataTableDbQuery(dbApplication, "MailingException", "SELECT RecipientEmail AS 'Получатель', " +
                                "NameReport AS 'Наименование', Description AS 'Описание', DateCreated AS 'Дата создания/модификации', " +
                                " DayReport AS 'День отправки отчета'", " ORDER BY RecipientEmail asc, DateCreated desc; ");
                                break;
                            }
                        case "SelectedCityToLoadFromWeb":
                            {
                                string[] cellValue =
                                    dgvo.FindValuesInCurrentRow(dataGridView1, new string[] { Names.PLACE_EMPLOYEE, @"Дата создания" });

                                switch (currColumn)
                                {
                                    case "Местонахождение сотрудника":
                                        ExecuteSqlAsync("UPDATE 'SelectedCityToLoadFromWeb' SET City='" + cellValue[0] +
                                                            "' WHERE DateCreated='" + cellValue[1] + "';");
                                        break;

                                    default:
                                        break;
                                }

                                ShowDataTableDbQuery(dbApplication, "SelectedCityToLoadFromWeb", "SELECT City AS 'Местонахождение сотрудника', DateCreated AS 'Дата создания'",
                                " ORDER BY City asc, DateCreated desc; ");
                                break;
                            }
                        default:
                            break;
                    }
                }
                catch (Exception expt) { logger.Trace($"DataGridView1CellEndEdit: {expt.ToString()}"); }
            }
        }

        //void or async Task
        private void ExecuteSqlAsync(string query) //Prepare DB and execute of SQL Query
        {
            if (dbApplication.Exists)
            {
                using (SqLiteDbWrapper dbWriter = new SqLiteDbWrapper(sqLiteLocalConnectionString, dbApplication))
                {
                    logger.Trace($"query: {query}");

                    dbWriter.Status += AddLoggerTraceText;
                    dbWriter.Execute(query);
                    dbWriter.Status -= AddLoggerTraceText;
                }
            }
        }

        
        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (DataGridViewOperations.ColumnsCount(dataGridView1) > 0 && DataGridViewOperations.RowsCount(dataGridView1) > 0)
            {
                if (
                    nameOfLastTable == @"PeopleGroup" ||
                    nameOfLastTable == @"Mailing" ||
                    nameOfLastTable == @"MailingException"
                    )
                {
                    // e.ColumnIndex == this.dataGridView1.Columns["NAV-код"].Index
                    DataGridViewCell cell = this.dataGridView1.Rows[e.RowIndex]?.Cells[e.ColumnIndex];
                    cell.ToolTipText = "Для установки нового значения нажмите F2,\nвнесите новое значение,\nа затем нажмите Enter";
                }
            }
        }

        //right click of mouse on the datagridview
        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {
            logger.Trace($"-= {nameof(dataGridView1_MouseClick)} =-");
            string[] cellValue;
            int currentMouseOverRow = dataGridView1.HitTest(e.X, e.Y).RowIndex;

            if (-1 < currentMouseOverRow)
            {
                if (e.Button == MouseButtons.Right)
                {
                    string txtboxGroup = ReturnTextOfControl(textBoxGroup);
                    string txtboxGroupDescription = ReturnTextOfControl(textBoxGroupDescription);

                    mRightClick = new ContextMenu();

                    switch (nameOfLastTable)
                    {
                        case "PeopleGroupDescription":
                            {
                                string recepient = "";
                                cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] { Names.GROUP, Names.GROUP_DECRIPTION, Names.RECEPIENTS_OF_REPORTS });

                                if (cellValue[2]?.Length > 0)
                                { recepient = cellValue[2]; }
                                else if (mailSenderAddress?.Length > 0)
                                { recepient = mailSenderAddress; }

                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: $"Загрузить регистрации пропусков группы: '{cellValue[1]}' за {ReturnDateTimePicker(dateTimePickerStart).ToMonthNameAndYear()}",
                                    onClick: GetDataOfGroup_Click));

                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: $"Загрузить регистрации пропусков группы: '{cellValue[1]}' за {ReturnDateTimePicker(dateTimePickerStart).ToMonthNameAndYear()} и подготовить отчет",
                                    onClick: DoReportByRightClick));

                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: $"Загрузить регистрации пропусков группы: '{cellValue[1]}' за {ReturnDateTimePicker(dateTimePickerStart).ToMonthNameAndYear()} и отправить: {recepient}",
                                    onClick: DoReportAndEmailByRightClick));

                                mRightClick.MenuItems.Add("-");
                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: $"Удалить группу: '{cellValue[0]}'({cellValue[1]})",
                                    onClick: DeleteCurrentRow));
                                mRightClick.Show(dataGridView1, new Point(e.X, e.Y));

                                break;
                            }
                        case "LastIputsOutputs":
                            {
                                cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] { Names.FIO, Names.N_ID_STRING, Names.CHECKPOINT_ACTION });

                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: "Обновить данные о регистрации входов-выходов сотрудников",
                                   onClick: LoadLastIputsOutputs_Update_Click));
                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: $"Подсветить все входы-выходы '{cellValue[0]}'",
                                   onClick: PaintRowsFioItem_Click));
                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: "Сбросить фильтр",
                                   onClick: ResetFilterLoadLastIputsOutput_Click));
                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: $"Подсветить все состояния '{cellValue[2]}'",
                                   onClick: PaintRowsActionItem_Click));
                                mRightClick.MenuItems.Add("-");
                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: $"Загрузить данные регистраций входов-выходов '{cellValue[0]}' за { ReturnDateTimePicker(dateTimePickerStart).ToMonthNameAndYear()}",
                                    onClick: GetDataOfPerson_Click));
                                mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                                break;
                            }
                        case "Mailing":
                            {
                                cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                                @"Наименование", @"Описание", @"День отправки отчета", @"Период", @"Тип отчета", @"Получатель"});

                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: $"Выполнить активные рассылки по всем у кого: тип отчета - {cellValue[4]} за {cellValue[3]} на {cellValue[2]}",
                                    onClick: SendAllReportsInSelectedPeriod));
                                mRightClick.MenuItems.Add("-");

                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: $"Выполнить рассылку: {cellValue[0]}({cellValue[1]}) для {cellValue[5]}",
                                    onClick: DoMainAction));
                                mRightClick.MenuItems.Add("-");

                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: $"Выполнить рассылку: {cellValue[0]}({cellValue[1]}) за {ReturnDateTimePicker(dateTimePickerStart).ToMonthNameAndYear()} для {cellValue[5]}",
                                    onClick: DoMainAction));
                                mRightClick.MenuItems.Add("-");

                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: "Создать новую рассылку",
                                    onClick: PrepareForMakingFormMailing));
                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: $"Клонировать рассылку: {cellValue[0]} ({cellValue[1]})",
                                    onClick: MakeCloneMailing));
                                mRightClick.MenuItems.Add("-");

                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: $"Состав рассылки: {cellValue[0]} ({cellValue[1]})",
                                    onClick: MembersGroupItem_Click));
                                mRightClick.MenuItems.Add("-");

                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: $"Удалить рассылку: {cellValue[0]} ({cellValue[1]})",
                                    onClick: DeleteCurrentRow));
                                mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                                break;
                            }
                        case "MailingException":
                            {
                                cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] { @"Получатель" });

                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: "Добавить новый адрес для исключения из рассылок отчетов",
                                    onClick: MakeNewRecepientExcept));
                                mRightClick.MenuItems.Add("-");
                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: $"Удалить адрес, ранее внесенный как 'исключеный из рассылок': {cellValue[0]}",
                                    onClick: DeleteCurrentRow));
                                mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                                break;
                            }
                        case "PeopleGroup":
                        case "ListFIO":
                            {
                                cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                                    Names.FIO,
                                    Names.CODE,
                                    Names.DEPARTMENT,
                                    Names.EMPLOYEE_POSITION,
                                    Names.CHIEF_ID,
                                    Names.EMPLOYEE_SHIFT,
                                    Names.DEPARTMENT_ID,
                                    Names.PLACE_EMPLOYEE,
                                    Names.GROUP
                                        });

                                if (string.Compare(cellValue[8], txtboxGroup) != 0 && txtboxGroup?.Length > 0) //добавить пункт меню если в текстбоксе группа другая
                                {
                                    mRightClick.MenuItems.Add(new MenuItem(
                                        text: $"Добавить '{cellValue[0]}' в группу '{txtboxGroup}'",
                                        onClick: AddPersonToGroupItem_Click));
                                    mRightClick.MenuItems.Add("-");
                                }
                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: $"Загрузить регистрации пропусков группы: '{cellValue[8]}' за {ReturnDateTimePicker(dateTimePickerStart).ToMonthNameAndYear()}",
                                    onClick: GetDataOfGroup_Click));
                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: $"Загрузить регистрации пропусков: '{cellValue[0]}' за {ReturnDateTimePicker(dateTimePickerStart).ToMonthNameAndYear()}",
                                    onClick: GetDataOfPerson_Click));
                                mRightClick.MenuItems.Add("-");
                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: $"Удалить '{cellValue[0]}' из группы '{txtboxGroup}'",
                                    onClick: DeleteCurrentRow));
                                mRightClick.Show(dataGridView1, new Point(e.X, e.Y)); break;
                            }
                        case "BoldedDates":
                            {
                                cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] { Names.DAYOFF_DATE, Names.DAYOFF_USED_BY, Names.DAYOFF_TYPE });

                                string dayType = "";
                                if (txtboxGroup?.Length == 0 || txtboxGroup?.ToLower() == "выходной")
                                { dayType = "Выходной"; }
                                else { dayType = "Рабочий"; }

                                string nav = "";
                                if (textBoxNav?.Text?.Trim()?.Length != 6)
                                { nav = "для всех"; }
                                else { nav = textBoxNav.Text.Trim(); }

                                string navD = "";
                                if (cellValue[1]?.Length != 6)
                                { navD = "всех"; }
                                else { navD = cellValue[1]; }

                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: $"Сохранить для {nav} как '{dayType} '{monthCalendar.SelectionStart.ToYYYYMMDD()}",
                                    onClick: AddAnualDateItem_Click));
                                mRightClick.MenuItems.Add("-");
                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: $"Удалить из сохранненых '{cellValue[2]}'  '{cellValue[0]}' для {navD}",
                                    onClick: DeleteAnualDateItem_Click));
                                mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                                break;
                            }
                        case "SelectedCityToLoadFromWeb":
                            {
                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: "Добавить новый город",
                                    onClick: AddNewCityToLoadByRightClick));
                                mRightClick.MenuItems.Add("-");
                                mRightClick.MenuItems.Add(new MenuItem(
                                    text: "Удалить выбранный город",
                                    onClick: DeleteCityToLoadByRightClick));
                                mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                                break;
                            }
                        default:
                            StatusLabel2.Text = "";
                            break;
                    }
                }
                else if (e.Button == MouseButtons.Left)
                {
                    if (0 < DataGridViewOperations.RowsCount(dataGridView1) && currentMouseOverRow < DataGridViewOperations.RowsCount(dataGridView1))
                    {
                        try
                        {
                            logger.Trace(nameOfLastTable);

                            switch (nameOfLastTable)
                            {
                                case "PeopleGroupDescription":
                                    {
                                        cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                                        Names.GROUP,
                                        Names.GROUP_DECRIPTION,
                                        Names.DEPARTMENT
                                        });

                                        textBoxGroup.Text = cellValue[0]; //Take the name of selected group
                                        textBoxGroupDescription.Text = cellValue[1]; //Take the name of selected group
                                        groupBoxPeriod.BackColor = Color.PaleGreen;
                                        groupBoxFilterReport.BackColor = SystemColors.Control;

                                        string descr = cellValue[2] ?? cellValue[1] ?? cellValue[0];

                                        StatusLabel2.Text = descr?.Length > 0 ?
                                            $"Выбрано: {descr}"
                                            : "";
                                        if (textBoxFIO?.TextLength > 3)
                                        {
                                            comboBoxFio.SelectedIndex = comboBoxFio.FindString(textBoxFIO.Text);
                                        }
                                        break;
                                    }

                                case "ListFIO":
                                case "PeopleGroup":
                                case "PersonRegistrationsList":
                                    {
                                        cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                                        Names.GROUP,
                                        Names.FIO,
                                        Names.CODE,
                                        Names.DEPARTMENT,
                                        Names.GROUP_DECRIPTION
                                        });

                                        textBoxGroup.Text = cellValue[0];
                                        textBoxFIO.Text = cellValue[1];
                                        textBoxNav.Text = cellValue[2];
                                        string descr = cellValue[3] ?? cellValue[4] ?? cellValue[2];
                                        logger.Trace($"{cellValue[0]} {cellValue[1]} {cellValue[2]} {cellValue[3]} {cellValue[4]} ");

                                        StatusLabel2.Text = descr?.Length > 0 || cellValue[1]?.Length > 0 ?
                                            $"Выбрано: {descr} |Курсор на: {cellValue[1]?.ConvertFullNameToShortForm()}"
                                            : "";

                                        groupBoxPeriod.BackColor = Color.PaleGreen;
                                        groupBoxTimeStart.BackColor = Color.PaleGreen;
                                        groupBoxTimeEnd.BackColor = Color.PaleGreen;
                                        groupBoxFilterReport.BackColor = SystemColors.Control;

                                        break;
                                    }

                                case "LastIputsOutputs":
                                    {
                                        cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                                        Names.N_ID_STRING,
                                        Names.FIO,
                                        Names.CHECKPOINT_ACTION
                                        });

                                        textBoxFIO.Text = cellValue[1];
                                        textBoxNav.Text = "";

                                        StatusLabel2.Text = cellValue[1]?.Length > 0 ?
                                            $"Выбрано : {cellValue[1]?.ConvertFullNameToShortForm()}"
                                            : "";
                                    }
                                    break;

                                case "Mailing":
                                    {
                                        cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                                        @"Наименование", @"Описание", @"День отправки отчета", @"Период", @"Тип отчета", @"Получатель"});

                                        StatusLabel2.Text = $"Выбран получатель {cellValue[5]}, день отправки {cellValue[2]}";
                                    }
                                    break;

                                default:
                                    StatusLabel2.Text = "";
                                    break;
                            }
                        }
                        catch (Exception err)
                        {
                            logger.Warn($"dataGridView1CellClick, {nameOfLastTable}: {err.ToString()}");
                        }
                    }
                }
            }
        }

        private void SelectedToLoadCityItem_Click(object sender, EventArgs e) //SelectedToLoadCity()
        { SelectedToLoadCity(); }

        private void SelectedToLoadCity()
        {
            ShowDataTableDbQuery(dbApplication, "SelectedCityToLoadFromWeb",
                "SELECT City AS 'Местонахождение сотрудника', DateCreated AS 'Дата создания'",
                " ORDER BY DateCreated desc; ");
        }

        private void AddNewCityToLoadByRightClick(object sender, EventArgs e)
        {
            AddNewCityToLoad();
            SelectedToLoadCity();
        }

        private void AddNewCityToLoad()
        {
            logger.Trace($"-= {nameof(AddNewCityToLoad)} =-");

            using (var sqlConnection = new SQLiteConnection(sqLiteLocalConnectionString))
            {
                sqlConnection.Open();
                using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'SelectedCityToLoadFromWeb' (City, DateCreated) " +
                    " VALUES (@City, @DateCreated)", sqlConnection))
                {
                    sqlCommand.Parameters.Add("@City", DbType.String).Value = "City";
                    sqlCommand.Parameters.Add("@DateCreated", DbType.String).Value = DateTime.Now.ToYYYYMMDDHHMM();
                    try { sqlCommand.ExecuteNonQuery(); } catch (Exception err) { MessageBox.Show(err.ToString()); }
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
            string[] cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                Names.PLACE_EMPLOYEE, "Дата создания"});

            DeleteDataTableQueryParameters(dbApplication, "SelectedCityToLoadFromWeb", "City", cellValue[0]).GetAwaiter();
        }

        private async void DoReportAndEmailByRightClick(object sender, EventArgs e)
        {
            await Task.Run(() => DoReportAndEmailByRightClick());
        }

        private void DoReportAndEmailByRightClick()
        {
            logger.Trace($"-= {nameof(DoReportAndEmailByRightClick)} =-");

            string[] cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                Names.GROUP,
                Names.GROUP_DECRIPTION,
                Names.RECEPIENTS_OF_REPORTS
            });
            resultOfSendingReports = new List<Mailing>();
            logger.Trace("DoReportAndEmailByRightClick");

            SetStatusLabelText(StatusLabel2, $"Готовлю отчет по группе {cellValue[0]}");
            nameOfLastTable = "Mailing";
            currentAction = "sendEmail";

            if (cellValue[2]?.Length > 0)
            {
                MailingAction(
                    "sendEmail",
                    cellValue[2],
                    mailSenderAddress,
                    cellValue[0],
                    cellValue[0],
                    cellValue[1],
                   ReturnDateTimePicker(dateTimePickerStart).ToMonthNameAndYear(),
                    "Активная",
                    "Упрощенный",
                    "END_OF_MONTH");
            }
            else if (mailSenderAddress?.Length > 0)
            {
                MailingAction("sendEmail", mailSenderAddress, mailSenderAddress,
             cellValue[0], cellValue[0], cellValue[1], SelectedDatetimePickersPeriodMonth(), "Активная", "Упрощенный", DateTime.Now.ToYYYYMMDDHHMM());
            }
            else
            {
                SetStatusLabelText(
                    StatusLabel2,
                    $"Попытка отправить отчет {cellValue[0]} не существующему получателю",
                    true);
            }

            SetStatusLabelForeColor(StatusLabel2, Color.Black);

            if (resultOfSendingReports.Count > 0)
            { SendReport(); }

            ProgressBar1Stop();
            nameOfLastTable = "PeopleGroupDescription";
        }

        private void DoReportByRightClick(object sender, EventArgs e)
        { DoReportByRightClick(); }

        private void DoReportByRightClick()
        {
            logger.Trace($"-= {nameof(DoReportByRightClick)} =-");

            string[] cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                Names.GROUP,
                Names.GROUP_DECRIPTION,
                Names.RECEPIENTS_OF_REPORTS
            });

            SetStatusLabelText(StatusLabel2, $"Готовлю отчет по группе {cellValue[0]}");
            logger.Trace("DoReportByRightClick: " + cellValue[0]);

            resultOfSendingReports = new List<Mailing>();

            GetRegistrationAndSendReport(
                cellValue[0],
                cellValue[0],
                cellValue[1],
                ReturnDateTimePicker(dateTimePickerStart).ToMonthNameAndYear(),
                "Активная",
                "Упрощенный",
                DateTime.Now.ToYYYYMMDDHHMM(),
                false, "", "");

            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("explorer.exe", " /select, " + filePathExcelReport + @".xlsx")); // //System.Reflection.Assembly.GetExecutingAssembly().Location)

            SetStatusLabelForeColor(StatusLabel2, Color.Black);

            if (resultOfSendingReports.Count > 0)
            { SendReport(); }

            ProgressBar1Stop();
            nameOfLastTable = "PeopleGroupDescription";
        }

        private void MakeCloneMailing(object sender, EventArgs e) //MakeCloneMailing()
        { MakeCloneMailing(); }

        private void MakeCloneMailing()
        {
            logger.Trace($"-= {nameof(MakeCloneMailing)} =-");

            string[] cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                @"Получатель", @"Отчет по группам", @"Наименование", @"Описание", @"Период"});

            SaveMailing(
               cellValue[0], cellValue[1], cellValue[2] + "_1",
               cellValue[3] + "_1", cellValue[4], "Неактивная", DateTime.Now.ToYYYYMMDDHHMM(), "", "Копия", DEFAULT_DAY_OF_SENDING_REPORT);

            ShowDataTableDbQuery(dbApplication, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
            "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
            "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
            " ORDER BY RecipientEmail asc, DateCreated desc; ");
        }

        private void MakeNewRecepientExcept(object sender, EventArgs e) //MakeNewRecepientExcept(), ShowDataTableDbQuery()
        {
            MakeNewRecepientExcept();
            ShowDataTableDbQuery(dbApplication, "MailingException", "SELECT RecipientEmail AS 'Получатель', " +
            "NameReport AS 'Наименование', Description AS 'Описание', DateCreated AS 'Дата создания/модификации', " +
            " DayReport AS 'День отправки отчета'", " ORDER BY RecipientEmail asc, DateCreated desc; ");
        }

        private void MakeNewRecepientExcept()
        {
            logger.Trace($"-= {nameof(MakeNewRecepientExcept)} =-");

            using (var sqlConnection = new SQLiteConnection(sqLiteLocalConnectionString))
            {
                sqlConnection.Open();
                using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'MailingException' (RecipientEmail, GroupsReport, Description, DateCreated) " +
                    " VALUES (@RecipientEmail, @GroupsReport, @Description, @DateCreated)", sqlConnection))
                {
                    sqlCommand.Parameters.Add("@RecipientEmail", DbType.String).Value = "example@mail.com";
                    sqlCommand.Parameters.Add("@GroupsReport", DbType.String).Value = "";
                    sqlCommand.Parameters.Add("@Description", DbType.String).Value = "example";
                    sqlCommand.Parameters.Add("@DateCreated", DbType.String).Value = DateTime.Now.ToYYYYMMDDHHMM();
                    try { sqlCommand.ExecuteNonQuery(); } catch (Exception err) { MessageBox.Show(err.ToString()); }
                }
            }
        }

        private void MailingsExceptItem_Click(object sender, EventArgs e)
        {
            EnableControl(comboBoxFio, false);
            dataGridView1.Select();

            ShowDataTableDbQuery(dbApplication, "MailingException", "SELECT RecipientEmail AS 'Получатель', " +
            "NameReport AS 'Наименование', Description AS 'Описание', DateCreated AS 'Дата создания/модификации', " +
            " DayReport AS 'День отправки отчета'", " ORDER BY RecipientEmail asc, DateCreated desc; ");
        }

        private void MailingsShowItem_Click(object sender, EventArgs e)
        {
            EnableControl(comboBoxFio, false);
            dataGridView1.Select();

            ShowDataTableDbQuery(dbApplication, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
            "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
            "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
            " ORDER BY RecipientEmail asc, DateCreated desc; ");
        }

        private async void DoMainAction(object sender, EventArgs e) //DoMainAction()
        {
            await Task.Run(() => DoMainAction());
        }

        private void DoMainAction()
        {
            logger.Trace($"-= {nameof(DoMainAction)} =-");
            ProgressBar1Start();

            switch (nameOfLastTable)
            {
                case "PeopleGroupDescription":
                case "PeopleGroup":
                    { break; }
                case "Mailing": //send report by e-mail
                    {
                        //текущий режим работы приложения
                        currentAction = "sendEmail";
                        resultOfSendingReports = new List<Mailing>();

                        string[] cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            @"Получатель", @"Отчет по группам", @"Наименование", @"Описание",
                            @"Период", @"Статус", @"Тип отчета", @"День отправки отчета" });
                        SetStatusLabelText(StatusLabel2, $"Готовлю отчет {cellValue[2]}");

                        ExecuteSqlAsync("UPDATE 'Mailing' SET SendingLastDate='" + DateTime.Now.ToYYYYMMDDHHMM()
                            + "' WHERE RecipientEmail='" + cellValue[0] + "' AND GroupsReport ='" + cellValue[1]
                            + "' AND NameReport='" + cellValue[2] + "' AND Description ='" + cellValue[3]
                            + "' AND Period='" + cellValue[4] + "' AND Status='" + cellValue[5]
                            + "' AND TypeReport='" + cellValue[6] + "' AND DayReport ='" + cellValue[7]
                            + "';");

                        MailingAction("sendEmail", cellValue[0], mailSenderAddress,
                            cellValue[1], cellValue[2], cellValue[3], cellValue[4],
                            cellValue[5], cellValue[6], cellValue[7]);

                        logger.Info("DoMainAction, sendEmail: Получатель: " +
                            cellValue[0] + "|" + cellValue[1] + "|Наименование: " + cellValue[2] + "|" +
                            cellValue[3] + "|Период: " + cellValue[4] + "|" + cellValue[5] + "|" +
                            cellValue[6] + "|" + cellValue[7]);

                        ShowDataTableDbQuery(dbApplication, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
                        "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
                        "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
                        " ORDER BY RecipientEmail asc, DateCreated desc; ");

                        SetStatusLabelForeColor(StatusLabel2, Color.Black);

                        if (resultOfSendingReports.Count > 0)
                        { SendReport(); }

                        break;
                    }
                default:
                    break;
            }

            ProgressBar1Stop();
        }

        private async void SendAllReportsInSelectedPeriod(object sender, EventArgs e) //SendAllReportsInSelectedPeriod()
        {
           await Task.Run(() => SendAllReportsInSelectedPeriod());
        }

        private void SendAllReportsInSelectedPeriod()
        {
            logger.Trace($"-= {nameof(SendAllReportsInSelectedPeriod)} =-");
            ProgressBar1Start();

            resultOfSendingReports = new List<Mailing>();

            string[] cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            @"Получатель", @"Отчет по группам", @"Наименование", @"Описание",
                            @"Период", @"Статус", @"Тип отчета", @"День отправки отчета" });

            SetStatusLabelText(StatusLabel2, $"Готовлю все активные рассылки с отчетами {cellValue[6]} за {cellValue[4]} на {cellValue[7]}");

            currentAction = "sendEmail";
            DoListsFioGroupsMailings();

            string recipient = "";
            string gproupsReport = "";
            string nameReport = "";
            string descriptionReport = "";
            string period = "";
            string status = "";
            string typeReport = "";
            string dayReport = "";
            string str = "";

            DataTable dtEmpty = new DataTable();
            EmployeeFull personEmpty = new EmployeeFull();

            SeekAnualDays(ref dtEmpty, ref personEmpty, false,
                DateTime.Now.FirstDayOfMonth().ToIntYYYYMMDD(), DateTime.Now.LastDayOfMonth().ToIntYYYYMMDD(),
                ref myBoldedDates, ref workSelectedDays
                );
            dtEmpty = null;
            personEmpty = null;
            DaysWhenSendReports daysToSendReports = new DaysWhenSendReports(workSelectedDays, ShiftDaysBackOfSendingFromLastWorkDay, DateTime.Now.LastDayOfMonth().Day);
            DaysOfSendingMail daysOfSendingMail = daysToSendReports.GetDays();

            logger.Trace(
                $"SendAllReportsInSelectedPeriod: активные отчеты {cellValue[6]} за {cellValue[4]} {DateTime.Now.FirstDayOfMonth().ToYYYYMMDD()} - {DateTime.Now.LastDayOfMonth().ToYYYYMMDD()} на дату - {cellValue[7]}");

            HashSet<Mailing> mailingList = new HashSet<Mailing>();

            using (var sqlConnection = new SQLiteConnection(sqLiteLocalConnectionString))
            {
                sqlConnection.Open();

                using (var sqlCommand = new SQLiteCommand("SELECT * FROM Mailing;", sqlConnection))
                {
                    using (var reader = sqlCommand.ExecuteReader())
                    {
                        foreach (DbDataRecord record in reader)
                        {
                            if (
                                record["RecipientEmail"]?.ToString()?.Length > 0 &&
                                record["DayReport"]?.ToString()?.Trim()?.ToUpper() == cellValue[7]
                                )
                            {
                                recipient = record["RecipientEmail"].ToString();
                                gproupsReport = record["GroupsReport"].ToString();
                                nameReport = record["NameReport"].ToString();
                                descriptionReport = record["Description"].ToString();
                                period = record["Period"].ToString();
                                status = record["Status"].ToString();
                                typeReport = record["TypeReport"].ToString();

                                string dayReportInDB = record["DayReport"]?.ToString()?.Trim();
                                dayReport = ReturnStrongNameDayOfSendingReports(dayReportInDB);

                                int dayToSendReport = ReturnNumberStrongNameDayOfSendingReports(dayReport, daysOfSendingMail);

                                logger.Trace($"DayReport: {dayReportInDB} {dayReport} {dayToSendReport}");

                                if (
                                        status == "Активная" &&
                                        typeReport == cellValue[6] &&
                                        period == cellValue[4] &&
                                        dayReportInDB == cellValue[7]
                                        )
                                {
                                    mailingList.Add(new Mailing()
                                    {
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

            foreach (Mailing mailng in mailingList)
            {
                SetStatusLabelText(StatusLabel2, $"Готовлю отчет {mailng._nameReport}");

                str = "UPDATE 'Mailing' SET SendingLastDate='" + DateTime.Now.ToYYYYMMDDHHMM() +
                    "' WHERE RecipientEmail='" + mailng._recipient +
                    "' AND NameReport='" + mailng._nameReport +
                    "' AND Period='" + mailng._period +
                    "' AND Status='" + mailng._status +
                    "' AND TypeReport='" + mailng._typeReport +
                    "' AND GroupsReport ='" + mailng._groupsReport + "';";
                logger.Trace(str);
                ExecuteSqlAsync(str);
                GetRegistrationAndSendReport(
                    mailng._groupsReport, mailng._nameReport, mailng._descriptionReport, mailng._period, mailng._status,
                    mailng._typeReport, mailng._dayReport, true, mailng._recipient, mailSenderAddress);

                ProgressWork1Step();
            }

            logger.Info("Перечень задач по подготовке и отправке отчетов завершен...");

            ShowDataTableDbQuery(dbApplication, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
            "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
            "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
            " ORDER BY RecipientEmail asc, DateCreated desc; ");

            SetStatusLabelForeColor(StatusLabel2, Color.Black);

            if (resultOfSendingReports.Count > 0)
            {
                SendReport();
            }

            ProgressBar1Stop();
        }

        private void DeleteCurrentRow(object sender, EventArgs e) //DeleteCurrentRow()
        {
            if (DataGridViewOperations.RowsCount(dataGridView1) > -1)
            { DeleteCurrentRow(); }
        }

        private void DeleteCurrentRow()
        {
            logger.Trace($"-= {nameof(DeleteCurrentRow)} =-");
            string group = ReturnTextOfControl(textBoxGroup);

            switch (nameOfLastTable)
            {
                case "PeopleGroupDescription":
                    {
                        string[] cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            Names.GROUP
                        });

                        DeleteDataTableQueryParameters(dbApplication, "PeopleGroup", "GroupPerson", cellValue[0], "", "", "", "").GetAwaiter().GetResult();
                        DeleteDataTableQueryParameters(dbApplication, "PeopleGroupDescription", "GroupPerson", cellValue[0], "", "", "", "").GetAwaiter().GetResult();

                        UpdateAmountAndRecepientOfPeopleGroupDescription();
                        ShowDataTableDbQuery(dbApplication,
                            "PeopleGroupDescription",
                            "SELECT GroupPerson AS 'Группа', GroupPersonDescription AS 'Описание группы', AmountStaffInDepartment AS 'Колличество сотрудников в группе' ", " group by GroupPerson ORDER BY GroupPerson asc; ");
                        SetStatusLabelText(StatusLabel2, "Удалена группа: " + cellValue[0] + "| Всего групп: " + DataGridViewOperations.RowsCount(dataGridView1));
                        MembersGroupItem.Enabled = true;
                        break;
                    }
                case "PeopleGroup" when group?.Length > 0:
                    {
                        int indexCurrentRow = DataGridViewOperations.RowsCount(dataGridView1);

                        string[] cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            Names.CODE,
                            Names.GROUP });
                        DeleteDataTableQueryParameters(dbApplication, "PeopleGroup", "GroupPerson", cellValue[1], "NAV", cellValue[0], "", "").GetAwaiter().GetResult();

                        if (indexCurrentRow > 2)
                        { SeekAndShowMembersOfGroup(group); }
                        else
                        {
                            DeleteDataTableQueryParameters(dbApplication, "PeopleGroupDescription", "GroupPerson", cellValue[1], "", "", "", "").GetAwaiter().GetResult();

                            UpdateAmountAndRecepientOfPeopleGroupDescription();
                            ShowDataTableDbQuery(dbApplication, "PeopleGroupDescription", "SELECT GroupPerson AS 'Группа', GroupPersonDescription AS 'Описание группы', AmountStaffInDepartment AS 'Колличество сотрудников в группе' ", " group by GroupPerson ORDER BY GroupPerson asc; ");
                        }

                        textBoxGroup.BackColor = Color.White;
                        MembersGroupItem.Enabled = true;
                        break;
                    }
                case "Mailing":
                    {
                        string[] cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            @"Получатель", @"Наименование", @"Дата создания/модификации",
                            @"Отчет по группам", @"Период", @"Тип отчета" });
                        DeleteDataTableQueryParameters(dbApplication, "Mailing",
                            "RecipientEmail", cellValue[0],
                            "NameReport", cellValue[1],
                            "DateCreated", cellValue[2],
                            "GroupsReport", cellValue[3],
                            "TypeReport", cellValue[5],
                            "Period", cellValue[4]).GetAwaiter().GetResult();

                        ShowDataTableDbQuery(dbApplication, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
                        "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
                        "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
                        " ORDER BY RecipientEmail asc, DateCreated desc; ");
                        SetStatusLabelText(StatusLabel2, "Удалена рассылка отчета " + cellValue[1] + "| Всего рассылок: " + DataGridViewOperations.RowsCount(dataGridView1));
                        break;
                    }
                case "MailingException":
                    {
                        string[] cellValue = dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            @"Получатель"});
                        DeleteDataTableQueryParameters(dbApplication, "MailingException",
                            "RecipientEmail", cellValue[0]).GetAwaiter().GetResult();

                        ShowDataTableDbQuery(dbApplication, "MailingException", "SELECT RecipientEmail AS 'Получатель', " +
                        "NameReport AS 'Наименование', Description AS 'Описание', DateCreated AS 'Дата создания/модификации', " +
                        "DayReport AS 'День отправки отчета'", " ORDER BY RecipientEmail asc, DateCreated desc; ");
                        SetStatusLabelText(StatusLabel2, "Удален из исключений " + cellValue[0] + "| Всего исключений: " + DataGridViewOperations.RowsCount(dataGridView1));
                        break;
                    }
                default:
                    break;
            }
            group = null;
        }

        //---  End.  DatagridView functions ---//

        //---  Start. Schedule Functions ---//

        private void ModeAppItem_Click(object sender, EventArgs e)  //ModeApp()
        { SwitchAppMode(); }

        private void SwitchAppMode()       // ExecuteAutoMode()
        {
            logger.Trace($"-= {nameof(SwitchAppMode)} =-");

            if (currentModeAppManual)
            {
                SetMenuItemText(ModeItem, "Выключить режим e-mail рассылок");
                SetMenuItemTooltip(ModeItem, "Включен автоматический режим. Выполняются Активные рассылки из БД.");
                SetMenuItemBackColor(ModeItem, Color.DarkOrange);

                SetStatusLabelText(StatusLabel2, "Включен режим рассылки отчетов по почте");
                SetStatusLabelBackColor(StatusLabel2, Color.PaleGreen); //Color.DarkOrange

                try
                {
                    using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(appRegistryKey))
                    {
                        EvUserKey.SetValue("ModeApp", "0", Microsoft.Win32.RegistryValueKind.String);
                        logger.Info("Ключ ModeApp в Registry сохранен");
                    }

                    using (Microsoft.Win32.RegistryKey EvUserKey =
                         Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run\"))
                    {
                        EvUserKey.SetValue(appName, "\"" + Application.ExecutablePath + "\"");
                        logger.Info("Ключ AutoRun App в Registry сохранен");
                    }
                }
                catch (Exception err) { logger.Warn("Save ModeApp,AutoRun in Registry. Последний режим работы не сохранен. " + err); }
                ExecuteAutoMode(true);
            }
            else
            {
                SetMenuItemText(ModeItem, "Включить режим автоматических e-mail рассылок");
                SetMenuItemTooltip(ModeItem, "Включен интерактивный режим. Все рассылки остановлены.");
                SetMenuItemBackColor(ModeItem, SystemColors.Control);

                SetStatusLabelText(StatusLabel2, "Интерактивный режим");
                SetStatusLabelBackColor(StatusLabel2, SystemColors.Control);

                try
                {
                    using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(appRegistryKey))
                    {
                        EvUserKey.SetValue("ModeApp", "1", Microsoft.Win32.RegistryValueKind.String);
                        logger.Info("Ключ ModeApp в Registry сохранен");
                    }

                    using (Microsoft.Win32.RegistryKey EvUserKey =
                        Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", true))
                    {
                        EvUserKey.DeleteValue(appName);
                        logger.Info("Ключ AutoRun App из Registry удален.");
                    }
                }
                catch (Exception err) { logger.Warn("Ошибка удаления ключа AutoRun App из Registry. " + err.ToString()); }
                ExecuteAutoMode(false);
            }

            //trigger mode
            if (currentModeAppManual)
            { currentModeAppManual = false; }
            else
            { currentModeAppManual = true; }
        }

        private void ExecuteAutoMode(bool manualMode) //InitScheduleTask()
        { Task.Run(() => InitScheduleTask(manualMode)); }

        private void InitScheduleTask(bool manualMode) //ScheduleTask()
        {
            long interval = 60 * 1000; //60 seconds
            if (manualMode)
            {
                SetMenuItemText(ModeItem, "Выключить режим e-mail рассылок");
                SetMenuItemTooltip(ModeItem, "Включен автоматический режим. Выполняются Активные рассылки из БД.");
                SetMenuItemBackColor(ModeItem, Color.DarkOrange);
                SetStatusLabelText(StatusLabel2, "Включен режим авторассылки отчетов");
                SetStatusLabelBackColor(StatusLabel2, Color.PaleGreen);

                timer?.Dispose();
                currentAction = "sendEmail";

                timer = new System.Threading.Timer(new System.Threading.TimerCallback(ScheduleTask), null, 0, interval);
            }
            else
            {
                SetMenuItemText(ModeItem, "Включить режим автоматических e-mail рассылок");
                SetMenuItemTooltip(ModeItem, "Включен интерактивный режим. Все рассылки остановлены.");
                SetMenuItemBackColor(ModeItem, SystemColors.Control);

                SetStatusLabelText(StatusLabel2, "Интерактивный режим");
                SetStatusLabelBackColor(StatusLabel2, SystemColors.Control);

                timer?.Dispose();
            }
        }

        private void ScheduleTask(object obj) //SelectMailingDoAction()
        {
            logger.Trace($"-= {nameof(ScheduleTask)} =-");

            lock (synclock)
            {
                DateTime dd = DateTime.Now;
                if (dd.Hour == 4 && dd.Minute == 10 && sent == false) //do something at Hour 2 and 5 minute //dd.Day == 1 &&
                {
                    SetStatusLabelText(StatusLabel2, "Ведется работа по подготовке отчетов " + DateTime.Now.ToYYYYMMDDHHMM() + " ...");
                    SetStatusLabelBackColor(StatusLabel2, Color.LightPink);
                    CheckAliveIntellectServer(sServer1, sqlServerConnectionString, queryCheckSystemDdSCA);
                    SelectMailingDoAction();
                    sent = true;
                    SetStatusLabelText(StatusLabel2, "Все задачи по подготовке и отправке отчетов завершены.");
                    logger.Info("");
                    logger.Info("---/  " + DateTime.Now.ToYYYYMMDDHHMMSS() + "  /---");
                }
                else
                {
                    sent = false;
                }

                if (dd.Hour == 7 && dd.Minute == 1)
                {
                    SetStatusLabelText(StatusLabel2, "Режим почтовых рассылок. " + DateTime.Now.ToYYYYMMDDHHMM());
                    SetStatusLabelBackColor(StatusLabel2, Color.LightCyan);
                    ClearItemsInFolder(@"*.xlsx");
                }
            }
        }

        private void TestToSendAllMailingsItem_Click(object sender, EventArgs e) //SelectMailingDoAction()
        { TestToSendAllMailings(); }

        private async void TestToSendAllMailings()
        {
            logger.Trace($"-= {nameof(TestToSendAllMailings)} =-");

            await Task.Run(() => CheckAliveIntellectServer(sServer1, sqlServerConnectionString, queryCheckSystemDdSCA));
            await Task.Run(() => UpdateMailingInDB());
            await Task.Run(() => SelectMailingDoAction());
        }

        private string ReturnStrongNameDayOfSendingReports(string inputDate)
        {
            string result;

            if (inputDate.Equals("1") || inputDate.Equals("01") || inputDate.Contains("ПЕРВ") || inputDate.Contains("НАЧАЛ") || inputDate.Contains("START"))
            {
                result = "START_OF_MONTH";
            }
            else if (inputDate.Equals("16") || inputDate.Equals("15") || inputDate.Equals("14") || inputDate.Contains("СЕРЕД") || inputDate.Contains("СРЕД") || inputDate.Contains("MIDDLE"))
            {
                result = "MIDDLE_OF_MONTH";
            }
            else { result = "END_OF_MONTH"; }

            return result;
        }

        private int ReturnNumberStrongNameDayOfSendingReports(string inputDate, DaysOfSendingMail daysOfSendingMail)
        {
            int result;

            if (inputDate.Equals("START_OF_MONTH"))
            {
                result = daysOfSendingMail.START_OF_MONTH;
            }
            else if (inputDate.Equals("MIDDLE_OF_MONTH"))
            {
                result = daysOfSendingMail.MIDDLE_OF_MONTH;
            }
            else
            {
                result = daysOfSendingMail.LAST_WORK_DAY_OF_MONTH;
            }
            return result;
        }

        private void SelectMailingDoAction()
        {
            logger.Trace($"-= {nameof(SelectMailingDoAction)} =-");

            ProgressBar1Start();

            currentAction = "sendEmail";
            resultOfSendingReports = new List<Mailing>();
            HashSet<Mailing> mailingList = new HashSet<Mailing>();

            DoListsFioGroupsMailings();

            string recipient = "";
            string gproupsReport = "";
            string nameReport = "";
            string descriptionReport = "";
            string period = "";
            string status = "";
            string typeReport = "";
            string dayReport = "";
            string str = "";

            DataTable dtEmpty = new DataTable();
            EmployeeFull personEmpty = new EmployeeFull();

            SeekAnualDays(ref dtEmpty, ref personEmpty, false,
                DateTime.Now.FirstDayOfMonth().ToIntYYYYMMDD(), DateTime.Now.LastDayOfMonth().ToIntYYYYMMDD(),
                ref myBoldedDates, ref workSelectedDays
                );

            SeekAnualDays(ref dtEmpty, ref personEmpty, false,
                DateTime.Now.FirstDayOfMonth().ToIntYYYYMMDD(), DateTime.Now.LastDayOfMonth().ToIntYYYYMMDD(),
                ref myBoldedDates, ref workSelectedDays
                );
            dtEmpty = null;
            personEmpty = null;

            dateTimePickerStart.Value = DateTime.Now.FirstDayOfMonth();

            DaysWhenSendReports daysToSendReports =
                new DaysWhenSendReports(workSelectedDays, ShiftDaysBackOfSendingFromLastWorkDay, DateTime.Now.LastDayOfMonth().Day);
            DaysOfSendingMail daysOfSendingMail = daysToSendReports.GetDays();

            logger.Trace("SelectMailingDoAction: all of daysOfSendingMail within the selected period: " +
                DateTime.Now.FirstDayOfMonth().ToYYYYMMDD() + " - " + DateTime.Now.LastDayOfMonth().ToYYYYMMDD() + ": " +
                daysOfSendingMail.START_OF_MONTH + ", " + daysOfSendingMail.MIDDLE_OF_MONTH + ", " +
                daysOfSendingMail.LAST_WORK_DAY_OF_MONTH + ", " + daysOfSendingMail.END_OF_MONTH
                );

            using (var sqlConnection = new SQLiteConnection(sqLiteLocalConnectionString))
            {
                sqlConnection.Open();

                using (var sqlCommand = new SQLiteCommand("SELECT * FROM Mailing;", sqlConnection))
                {
                    using (var reader = sqlCommand.ExecuteReader())
                    {
                        foreach (DbDataRecord record in reader)
                        {
                            if (record["RecipientEmail"]?.ToString()?.Length > 0)
                            {
                                recipient = record["RecipientEmail"].ToString();
                                gproupsReport = record["GroupsReport"].ToString();
                                nameReport = record["NameReport"].ToString();
                                descriptionReport = record["Description"].ToString();
                                period = record["Period"].ToString();
                                status = record["Status"].ToString().ToLower();
                                typeReport = record["TypeReport"].ToString();

                                string dayReportInDB = record["DayReport"]?.ToString()?.Trim()?.ToUpper();
                                dayReport = ReturnStrongNameDayOfSendingReports(dayReportInDB);

                                int dayToSendReport = ReturnNumberStrongNameDayOfSendingReports(dayReport, daysOfSendingMail);

                                logger.Trace("DayReport: " + dayReportInDB + " " + dayReport + "| dayToSendReport: " + dayToSendReport + "| today.Day: " + DateTime.Now.Day);

                                if (status == "активная" && dayToSendReport == DateTime.Now.Day)
                                {
                                    mailingList.Add(new Mailing()
                                    {
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
                                else if (status == "активная")
                                {
                                    logger.Trace("SelectMailingDoAction: не добавлена рассылка: " +
                                        recipient + "| " +
                                        gproupsReport + "| " +
                                        nameReport + "| " +
                                        period + "| " +
                                        typeReport + "| " +
                                        dayReport + "| "
                                        );
                                }
                            }
                        }
                    }
                }
            }

            if (mailingList.Count > 0)
            {
                logger.Info("Выполняю сбор данных регистраций и рассылку отчетов за период: " +
                        DateTime.Now.FirstDayOfMonth().ToYYYYMMDD() + " - " + DateTime.Now.LastDayOfMonth().ToYYYYMMDD()
                        );

                foreach (Mailing mailng in mailingList)
                {
                    SetStatusLabelText(StatusLabel2, "Готовлю отчет " + mailng._nameReport);

                    str = "UPDATE 'Mailing' SET SendingLastDate='" + DateTime.Now.ToYYYYMMDDHHMM() +
                        "' WHERE RecipientEmail='" + mailng._recipient +
                        "' AND NameReport='" + mailng._nameReport +
                        "' AND Period='" + mailng._period +
                        "' AND Status='" + mailng._status +
                        "' AND TypeReport='" + mailng._typeReport +
                        "' AND GroupsReport ='" + mailng._groupsReport + "';";
                    logger.Trace(str);
                    ExecuteSqlAsync(str);
                    GetRegistrationAndSendReport(
                        mailng._groupsReport, mailng._nameReport, mailng._descriptionReport, mailng._period, mailng._status,
                        mailng._typeReport, mailng._dayReport, true, mailng._recipient, mailSenderAddress);

                    ProgressWork1Step();
                }
                logger.Info("SelectMailingDoAction: Перечень задач по подготовке и отправке отчетов завершен...");
            }
            ShowDataTableDbQuery(dbApplication, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
            "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
            "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
            " ORDER BY RecipientEmail asc, DateCreated desc; ");

            SetStatusLabelForeColor(StatusLabel2, Color.Black);

            if (resultOfSendingReports.Count > 0)
            { SendReport(); }

            ProgressBar1Stop();
        }

        private void UpdateMailingInDB()
        {
            logger.Trace($"-= {nameof(UpdateMailingInDB)} =-");

            ProgressBar1Start();

            string recipient;
            string gproupsReport;
            string nameReport;
            string descriptionReport;
            string period;
            string status;
            string typeReport;
            string dayReport;
            string str;

            HashSet<Mailing> mailingList = new HashSet<Mailing>();

            using (var sqlConnection = new SQLiteConnection(sqLiteLocalConnectionString))
            {
                sqlConnection.Open();

                using (var sqlCommand = new SQLiteCommand("SELECT * FROM Mailing;", sqlConnection))
                {
                    using (var reader = sqlCommand.ExecuteReader())
                    {
                        foreach (DbDataRecord record in reader)
                        {
                            if (record["RecipientEmail"]?.ToString()?.Length > 0)
                            {
                                recipient = record["RecipientEmail"].ToString();
                                gproupsReport = record["GroupsReport"].ToString();
                                nameReport = record["NameReport"].ToString();
                                descriptionReport = record["Description"].ToString();
                                period = record["Period"].ToString();
                                status = record["Status"].ToString().ToLower();
                                typeReport = record["TypeReport"].ToString();

                                string dayReportInDB = record["DayReport"]?.ToString()?.Trim()?.ToUpper();
                                logger.Trace("DayReport: " + dayReportInDB);

                                dayReport = ReturnStrongNameDayOfSendingReports(dayReportInDB);

                                mailingList.Add(new Mailing()
                                {
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

            foreach (Mailing mailng in mailingList)
            {
                str = "UPDATE 'Mailing' SET DayReport='" + mailng._dayReport +
                    "' WHERE RecipientEmail='" + mailng._recipient +
                    "' AND NameReport='" + mailng._nameReport +
                    "' AND Period='" + mailng._period +
                    "' AND Status='" + mailng._status +
                    "' AND TypeReport='" + mailng._typeReport +
                    "' AND GroupsReport ='" + mailng._groupsReport + "';";

                logger.Trace(str);
                ExecuteSqlAsync(str);

                ProgressWork1Step();
            }

            logger.Info("UpdateMailingInDB: Перечень задач по подготовке и отправке отчетов завершен...");

            ShowDataTableDbQuery(dbApplication, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
            "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
            "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
            " ORDER BY RecipientEmail asc, DateCreated desc; ");

            ProgressBar1Stop();
        }

        private string GetSafeFilename(string filename, string splitter = "_")
        {
            return string.Join(splitter, filename.Split(System.IO.Path.GetInvalidFileNameChars()));
        }

        private string SelectedDatetimePickersPeriodMonth() //format of result: "1971-01-01 00:00:00|1971-01-31 23:59:59" // 'yyyy-MM-dd HH:mm:SS'
        {
            return ReturnDateTimePicker(dateTimePickerStart).ToYYYYMMDD() + " 00:00:00" + "|" + ReturnDateTimePicker(dateTimePickerEnd).ToYYYYMMDD() + " 23:59:59";
        }

        private void MailingAction(string mainAction, string recipientEmail, string senderEmail, string groupsReport, string nameReport, string description, string period, string status, string typeReport, string dayReport)
        {
            logger.Trace($"-= {nameof(MailingAction)} =-");
            logger.Trace($"mainAction: { mainAction} |recipientEmail: { recipientEmail} |senderEmail: { senderEmail} |groupsReport: { groupsReport} |nameReport: { nameReport} |description: { description} |period: { period} |status: { status} |typeReport: { typeReport} |dayReport: { dayReport} ");

            SetStatusLabelBackColor(StatusLabel2, SystemColors.Control);

            switch (mainAction)
            {
                case "saveEmail":
                    {
                        SaveMailing(mailSenderAddress, groupsReport, nameReport, description, period, status, DateTime.Now.ToYYYYMMDDHHMM(), "", typeReport, dayReport);

                        ShowDataTableDbQuery(dbApplication, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
                        "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
                        "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
                        " ORDER BY RecipientEmail asc, DateCreated desc; ");
                        break;
                    }
                case "sendEmail":
                    {
                        CheckAliveIntellectServer(sServer1, sqlServerConnectionString, queryCheckSystemDdSCA);

                        if (bServer1Exist)
                        {
                            GetRegistrationAndSendReport(groupsReport, nameReport, description, period, status, typeReport, dayReport, true, recipientEmail, senderEmail);
                            logger.Info("MailingAction: Задача по подготовке и отправке отчета '" + nameReport + "' выполнена ");
                        }
                        break;
                    }
                default:
                    break;
            }
        }

        private void GetRegistrationAndSendReport(string groupsReport, string nameReport, string description, string period, string status, string typeReport, string dayReport, bool sendReport, string recipientEmail, string senderEmail)
        {
            logger.Trace($"-= {nameof(GetRegistrationAndSendReport)} =-");

            DataTable dtTempIntermediate = dtPeople.Clone();
            EmployeeFull person = new EmployeeFull();
            DateTime selectedDate = DateTime.Now;

            GetNamesOfPassagePoints();  //Get names of the Passage Points

            if (period.ToLower().Contains("текущ"))
            { }
            else if (period.ToLower().Contains("предыдущ"))
            { selectedDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddDays(-1); }
            else
            { selectedDate = ReturnDateTimePicker(dateTimePickerStart); }

            reportStartDay = selectedDate.FirstDayOfMonth().ToYYYYMMDD() + " 00:00:00";
            reportLastDay = selectedDate.LastDayOfMonth().ToYYYYMMDD() + " 23:59:59";

            /*  if (!period.ToLower().Contains("предыдущ") && !period.ToLower().Contains("текущ"))
              {
                  reportStartDay = SelectedDatetimePickersPeriodMonth().Split('|')[0];
                  reportLastDay = SelectedDatetimePickersPeriodMonth().Split('|')[1];
              }*/

            DateTime start = DateTime.Parse(reportStartDay);
            DateTime end = DateTime.Parse(reportLastDay);

            SeekAnualDays(ref dtTempIntermediate, ref person, false, start.ToIntYYYYMMDD(), end.ToIntYYYYMMDD(), ref myBoldedDates, ref workSelectedDays);
            logger.Trace(reportStartDay + " - " + reportLastDay);

            string nameGroup = "";

            string titleOfbodyMail = "";
            string[] groups = groupsReport.Split('+');

            foreach (string groupName in groups)
            {
                SetStatusLabelBackColor(StatusLabel2, SystemColors.Control);
                nameGroup = groupName.Trim();
                if (nameGroup.Length > 0)
                {
                    dtPersonRegistrationsFullList?.Clear();
                    LoadRecords(groupName, reportStartDay, reportLastDay, "sendEmail");//typeReport== only one group

                    dtTempIntermediate = dtPeople.Clone();
                    dtPersonTemp = dtPeople.Clone();

                    person = new EmployeeFull();

                    dtPeopleGroup = LoadGroupMembersFromDbToDataTable(nameGroup);

                    foreach (DataRow row in dtPeopleGroup.Rows)
                    {
                        if (row[Names.FIO]?.ToString()?.Length > 0 &&
                            (row[Names.GROUP]?.ToString() == nameGroup || (@"@" + row[Names.DEPARTMENT_ID]?.ToString()) == nameGroup))
                        {
                            person = new EmployeeFull()
                            {
                                fio = row[Names.FIO].ToString(),
                                Code = row[Names.CODE].ToString(),

                                GroupPerson = row[Names.GROUP].ToString(),
                                Department = row[Names.DEPARTMENT].ToString(),
                                PositionInDepartment = row[Names.EMPLOYEE_POSITION].ToString(),
                                DepartmentId = row[Names.DEPARTMENT_ID].ToString(),
                                City = row[Names.PLACE_EMPLOYEE].ToString(),

                                ControlInSeconds = row[Names.DESIRED_TIME_IN].ToString().ConvertTimeAsStringToSeconds(),
                                ControlOutSeconds = row[Names.DESIRED_TIME_OUT].ToString().ConvertTimeAsStringToSeconds(),
                                ControlInHHMM = row[Names.DESIRED_TIME_IN].ToString(),
                                ControlOutHHMM = row[Names.DESIRED_TIME_OUT].ToString(),

                                Comment = row[Names.EMPLOYEE_SHIFT_COMMENT].ToString(),
                                Shift = row[Names.EMPLOYEE_SHIFT].ToString()
                            };

                            FilterRegistrationsOfPerson(ref person, dtPersonRegistrationsFullList, ref dtTempIntermediate, typeReport);
                        }
                    }

                    dtPersonTemp = LeaveAndOrderColumnsOfDataTable(dtTempIntermediate, Names.orderColumnsFinacialReport);

                    logger.Trace("dtTempIntermediate: " + dtTempIntermediate.Rows.Count);
                    logger.Trace("dtPersonTemp: " + dtPersonTemp.Rows.Count);

                    if (dtPersonTemp.Rows.Count > 0)
                    {
                        string nameFile = nameReport + " " + reportStartDay.Split(' ')[0] + "-" + reportLastDay.Split(' ')[0] + " " + groupName + " от " + DateTime.Now.ToYYYYMMDDHHMM();
                        string illegal = GetSafeFilename(nameFile);
                        filePathExcelReport = System.IO.Path.Combine(localAppFolderPath, illegal);

                        logger.Trace("Подготавливаю отчет: " + filePathExcelReport + @".xlsx");
                        ExportDatatableSelectedColumnsToExcel(dtPersonTemp, nameReport, filePathExcelReport);

                        if (sendReport)
                        {
                            if (reportExcelReady)
                            {
                                titleOfbodyMail = "с " + reportStartDay.Split(' ')[0] + " по " + reportLastDay.Split(' ')[0];
                                SetStatusLabelText(StatusLabel2, "Выполняю отправку отчета адресату: " + recipientEmail);

                                foreach (var oneAddress in recipientEmail.Split(','))
                                {
                                    if (oneAddress.Contains('@'))
                                    {
                                        SendReport(oneAddress.Trim(), titleOfbodyMail, description, filePathExcelReport + @".xlsx", appName);
                                        logger.Trace($"SendEmail, From: {mailSenderAddress}| To: {oneAddress}| Subject: {titleOfbodyMail}| {description}| attached: {filePathExcelReport}.xlsx"
                                            );
                                    }
                                }

                                SetStatusLabelText(StatusLabel2, $"{DateTime.Now.ToYYYYMMDDHHMM()} Отчет '{nameReport}'({groupName}) отправлен {recipientEmail}");
                                SetStatusLabelBackColor(StatusLabel2, Color.PaleGreen);
                            }
                            else
                            {
                                SetStatusLabelText(
                                    StatusLabel2,
                                    $"{DateTime.Now.ToYYYYMMDDHHMM()} Ошибка экспорта в файл отчета: {nameReport}({groupName})",
                                    true
                                    );
                            }
                        }
                    }
                    else
                    {
                        SetStatusLabelText(
                            StatusLabel2,
                            $"{DateTime.Now.ToYYYYMMDDHHMM()} Ошибка получения данных для отчета: {nameReport}",
                            true
                            );
                    }
                }
            }
            titleOfbodyMail = null;

            dtTempIntermediate?.Dispose();
            dtPersonTemp?.Clear();
        }

        //send e-mail
        private static string SendEmailAsync(MailServer server, MailUser from, MailUser to, string _subject, BodyBuilder builder)
        {
            logger.Trace($"-= {nameof(SendEmailAsync)} =-");

            MailSender mailSender = new MailSender(server);
            mailSender.SetFrom(from);
            mailSender.SetTo(to);

            mailSender.SendEmailAsync(_subject, builder).GetAwaiter().GetResult();
            return mailSender.Status;
        }

        //Compose Standart Report and send e-mail to recepient
        private static void SendReport(string to, string period, string department, string pathToFile, string messageAfterPicture)
        {
            logger.Trace($"-= {nameof(SendReport)} =-");

            MailUser _to = new MailUser(to.Split('@')[0], to);
            string subject = "Отчет по посещаемости за период: " + period;

            BodyBuilder builder = BuildMessage(period, department);
            builder.Attachments.Add(pathToFile);

            string statusOfSentEmail = SendEmailAsync(_mailServer, _mailUser, _to, subject, builder);    //send e-mail

            if (resultOfSendingReports == null)
            {
                resultOfSendingReports = new List<Mailing>();
            }

            resultOfSendingReports.Add(new Mailing
            {
                _recipient = to,
                _nameReport = pathToFile,
                _descriptionReport = department,
                _period = period,
                _status = statusOfSentEmail
            });

            stimerPrev = "отчет " + to + " отправлен " + statusOfSentEmail;
            logger.Info("SendStandartReport: " + statusOfSentEmail);
        }

        private static void MakeByteLogo(Bitmap logo)
        {
            //create e-mail logo
            Bitmap b = new Bitmap(logo, new Size(50, 50));
            ImageConverter ic = new ImageConverter();
            byteLogo = (Byte[])ic.ConvertTo(b, typeof(Byte[]));
            //convert embedded resources into memory stream to attach at an email
            // System.IO.MemoryStream logo = new System.IO.MemoryStream(byteLogo);
            // mailLogo = new System.Net.Mail.LinkedResource(logo, "image/jpeg");
            // mailLogo.ContentId = Guid.NewGuid().ToString(); //myAppLogo for email's reports
        }

        private static BodyBuilder BuildMessage(string period, string department)
        {
            var builder = new BodyBuilder();
            //plain-text version of the message text
            /*    builder.TextBody = @"

                     Hey User,
                     What are you up to this weekend? I am throwing one of my parties on
                     Saturday and I was hoping you could make it.

                     Will you be my +1?

                     -- Yuri
                    ";
            */

            //add the app's Logo
            // to builder.LinkedResources and then use its Content-Id value in the img src.
            MimeEntity image = builder.LinkedResources.Add("asta.jpg", byteLogo, new ContentType("image", "jpg"));
            image.ContentId = MimeKit.Utils.MimeUtils.GenerateMessageId();

            //  var generic = new MimePart("image", "jpg") { ContentObject = new ContentObject(logo), IsAttachment = true };
            //original - var generic = new MimePart("application", "octet-stream") { ContentObject = new ContentObject(new MemoryStream()), IsAttachment = true };

            // Set the html version of the message text
            builder.HtmlBody = $@"<p><font size='3' color='black' face='Arial'>Здравствуйте,</p>
            Во вложении «Отчет по учету рабочего времени сотрудников».<p>
            <b>Период: </b>{period}
            <br/><b>Подразделение: </b>'{department}'<br/><p>Уважаемые руководители,</p>
            <p>согласно Приказу ГК АИС «О функционировании процессов кадрового делопроизводства»,<br/><br/>
            <b>Внесите,</b> пожалуйста, <b>до конца текущего месяца</b> по сотрудникам подразделения
            информацию о командировках, больничных, отпусках, прогулах и т.п. <b>на сайт</b> www.ais .<br/><br/>
            Руководители <b>подразделений</b> ЦОК, <b>не отображающихся на сайте,<br/>вышлите, </b>пожалуйста,
            <b>Табель</b> учета рабочего времени<br/>
            <b>в отдел компенсаций и льгот до последнего рабочего дня месяца.</b><br/></p>
            <font size='3' color='black' face='Arial'>С, Уважением,<br/>{NAME_OF_SENDER_REPORTS}
            </font><br/><br/><font size='2' color='black' face='Arial'><i>
            Данное сообщение и отчет созданы автоматически<br/>программой по учету рабочего времени сотрудников.
            </i></font><br/><font size='1' color='red' face='Arial'><br/>{DateTime.Now.ToYYYYMMDDHHMM()}
            </font></p><hr><img alt='ASTA' src='cid:{image.ContentId}'/><br/><a href='mailto:ryik.yuri@gmail.com'>_</a>";

            // We may also want to attach a calendar event for Yuri's party...
            //  builder.Attachments.Add(@"C:\Users\Yuri\Documents\party.ics");
            return builder;
        }

        //Compose Admin Report and send e-mail to Administrator
        private static void SendReport()
        {
            logger.Trace($"-= {nameof(SendReport)} =-");

            string period = DateTime.Now.ToYYYYMMDD();
            string subject = $"Результат отправки отчетов за {period}";
            MailUser _to = new MailUser(mailJobReportsOfNameOfReceiver.Split('@')[0], mailJobReportsOfNameOfReceiver);

            BodyBuilder builder = BuildMessage(period, resultOfSendingReports);
            string statusOfSentEmail = SendEmailAsync(_mailServer, _mailUser, _to, subject, builder);

            logger.Trace($"Try to send, From: {mailSenderAddress}| To:{mailJobReportsOfNameOfReceiver}| {period}");
            logger.Info($"SendAdminReport: {statusOfSentEmail}");
            stimerPrev = $"Административный отчет отправлен {statusOfSentEmail}";
        }

        private static BodyBuilder BuildMessage(string period, List<Mailing> reportOfResultSending)
        {
            var builder = new BodyBuilder();

            // to builder.LinkedResources and then use its Content-Id value in the img src.
            MimeEntity image = builder.LinkedResources.Add("asta.jpg", byteLogo, new ContentType("image", "jpg"));
            image.ContentId = MimeKit.Utils.MimeUtils.GenerateMessageId();

            var sb = new StringBuilder();
            int order = 0;
            sb.Append(@"<style type = 'text/css'> A {text - decoration: none;}</ style >");
            sb.Append(@"<p><font size='4' color='black' face='Arial'>Здравствуйте,</p><p>Результаты отправки отчетов</p>");

            sb.Append(@"<b>Дата: </b>");
            sb.Append(period);
            sb.Append(@"<p><b>Результат отправки " + reportOfResultSending.Count + @" отчета/отчетов:</b></p><p>");

            foreach (var mailing in reportOfResultSending)
            {
                if (mailing != null && mailing._status != null)
                {
                    order += 1;
                    if (
                    mailing._status.ToLower().Contains("error") || mailing._status.ToLower().Contains("not been sent") || mailing._status.ToLower().Contains("unknown"))
                    {
                        sb.Append(@"<font size='2' color='red' face='Arial'>");
                        sb.Append(order + @". Получатель: " + mailing._recipient +
                                  @", отчет по " + mailing._descriptionReport +
                                  @" не получил: " + mailing._status + @"</font><br/>");
                    }
                    else
                    {
                        sb.Append(@"<font size='2' color='black' face='Arial'>");
                        sb.Append(order + @". Получатель: " + mailing._recipient +
                                  @", отчет по группе " + mailing._descriptionReport +
                                  @", доставлен " + mailing._status + @" </font><br/>");
                    }
                }
            }

            sb.Append(@"</p><font size='2' color='black' face='Arial'>С, Уважением,<br/>");
            sb.Append(NAME_OF_SENDER_REPORTS);
            sb.Append(@"</font><br/><br/><font size='1' color='black' face='Arial'><i>");
            sb.Append(@"Данное сообщение и отчет созданы автоматически<br/>программой по учету рабочего времени сотрудников.");
            sb.Append(@"</i></font><br/><font size='1' color='red' face='Arial'><br/>");
            sb.Append(DateTime.Now.ToYYYYMMDDHHMM());
            sb.Append(@"</font></p>");
            sb.Append(@"<hr><img alt='ASTA' src='cid:");
            sb.Append(image.ContentId);
            sb.Append(@"'/><br/><a href='mailto:ryik.yuri@gmail.com'>_</a>");

            // Set the html version of the message text
            builder.HtmlBody = sb.ToString();

            return builder;
        }

        //---  End. Schedule Functions ---//

        /// <summary>
        /// Colorizing Controls
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Set_ComboBox_DrawItem(object sender, DrawItemEventArgs e) //Colorize the Combobox
        {
            Font font = (sender as ComboBox).Font;
            Brush backgroundColor;
            Brush textColor = Brushes.Black;

            if (e.Index % 2 != 0)
            {
                if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                {
                    backgroundColor = Brushes.SandyBrown;
                    //  textColor = Brushes.Black;
                }
                else
                {
                    backgroundColor = SystemBrushes.Window;
                    //   textColor = Brushes.Black;
                }
            }
            else
            {
                if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                {
                    backgroundColor = Brushes.SandyBrown;// SystemBrushes.Highlight;
                                                         // textColor = Brushes.Black;// SystemBrushes.HighlightText;
                }
                else
                {
                    backgroundColor = Brushes.PaleTurquoise; // SystemBrushes.Window;
                                                             //  textColor = Brushes.Black;// SystemBrushes.WindowText;
                }
            }
            e.Graphics.FillRectangle(backgroundColor, e.Bounds);
            e.Graphics.DrawString((sender as ComboBox).Items[e.Index].ToString(), font, textColor, e.Bounds);
        }

        private void Set_ListBox_DrawItem(object sender, DrawItemEventArgs e) //Colorize the Combobox
        {
            Font font = (sender as ListBox).Font;
            Brush backgroundColor;
            Brush textColor = Brushes.Black;

            if (e.Index % 2 != 0)
            {
                if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                {
                    backgroundColor = Brushes.SandyBrown;
                    //  textColor = Brushes.Black;
                }
                else
                {
                    backgroundColor = SystemBrushes.Window;
                    //   textColor = Brushes.Black;
                }
            }
            else
            {
                if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                {
                    backgroundColor = Brushes.SandyBrown;// SystemBrushes.Highlight;
                                                         // textColor = Brushes.Black;// SystemBrushes.HighlightText;
                }
                else
                {
                    backgroundColor = Brushes.PaleTurquoise; // SystemBrushes.Window;
                                                             //  textColor = Brushes.Black;// SystemBrushes.WindowText;
                }
            }
            e.Graphics.FillRectangle(backgroundColor, e.Bounds);
            e.Graphics.DrawString((sender as ListBox).Items[e.Index].ToString(), font, textColor, e.Bounds);
        }

        //clear all registered Click events on the selected button
        private void ClearButtonClickEvent(Button b)
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
        private string ReturnTextOfControl(Control control) //add string into  from other threads
        {
            string text = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { text = control?.Text?.Trim(); }));
            else
                text = control?.Text?.Trim();
            return text;
        }

        private void AddComboBoxItem(ComboBox comboBx, string s) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { comboBx.Items.Add(s); }));
            else
                comboBx.Items.Add(s);
        }

        private void ClearComboBox(ComboBox comboBx) //add string into  from other threads
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

        private void SetComboBoxIndex(ComboBox comboBx, int i) //from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { try { comboBx.SelectedIndex = i; } catch { } }));
            else
                try { comboBx.SelectedIndex = i; } catch { }
        }

        private int? ReturnComboBoxCountItems(ComboBox comboBx) //from other threads
        {
            int? count = 0;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { count = comboBx?.Items?.Count; }));
            else
            { count = comboBx?.Items?.Count; }
            return count;
        }

        private string ReturnComboBoxSelectedItem(ComboBox comboBox) //from other threads
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

        private string ReturnListBoxSelectedItem(ListBox listBox) //from other threads
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

        private void SetNumUpDown(NumericUpDown numericUpDown, decimal i) //add string into comboBoxTargedPC from other threads
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

        private decimal ReturnNumUpDown(NumericUpDown numericUpDown)
        {
            decimal iCombo = 0;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { iCombo = numericUpDown.Value; }));
            else
                iCombo = numericUpDown.Value;
            return iCombo;
        }

        private DateTime ReturnDateTimePicker(DateTimePicker dateTimePicker) //add string into  from other threads
        {
            DateTime result = DateTime.Now;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    result = dateTimePicker.Value;
                }
                ));
            else
            {
                result = dateTimePicker.Value;
            }
            return result;
        }

        private int[] ReturnDateTimePickerArray(DateTimePicker dateTimePicker) //add string into  from other threads
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

        private void SetStatusLabelText(ToolStripStatusLabel statusLabel, string text, bool error = false) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    statusLabel.Text = text;
                    if (error) statusLabel.BackColor = Color.DarkOrange;
                }));
            else
            {
                statusLabel.Text = text;
                if (error) statusLabel.BackColor = Color.DarkOrange;
            }

            AddLoggerTraceText($"{nameof(SetStatusLabelText)} set text as {text}");
        }

        private void SetStatusLabelForeColor(ToolStripStatusLabel statusLabel, Color color)
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { statusLabel.ForeColor = color; }));
            else
                statusLabel.ForeColor = color;
        }

        private void SetStatusLabelBackColor(ToolStripStatusLabel statusLabel, Color color) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { statusLabel.BackColor = color; }));
            else
                statusLabel.BackColor = color;
        }

        private string ReturnStatusLabelText(ToolStripStatusLabel statusLabel)
        {
            string s = null;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    s = statusLabel.Text;
                }));
            else
            {
                s = statusLabel.Text;
            }
            return s;
        }

        private Color ReturnStatusLabelBackColor(ToolStripStatusLabel statusLabel) //add string into  from other threads
        {
            Color s = SystemColors.ControlText;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { s = statusLabel.BackColor; }));
            else
                s = statusLabel.BackColor;

            return s;
        }

        private void SetMenuItemBackColor(ToolStripMenuItem menuItem, Color color) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { menuItem.BackColor = color; }));
            else
                menuItem.BackColor = color; ;
        }

        private void SetMenuItemText(ToolStripMenuItem menuItem, string text) //access from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    menuItem.Text = text;
                }));
            else
            {
                menuItem.Text = text;
            }

            AddLoggerTraceText($"{nameof(SetMenuItemText)}: {nameof(menuItem.Name)} set text as '{text}'");
        }

        private void SetMenuItemTooltip(ToolStripMenuItem menuItem, string text) //access from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    menuItem.ToolTipText = text;
                }));
            else
            {
                menuItem.ToolTipText = text;
            }

            AddLoggerTraceText($"{nameof(SetMenuItemTooltip)}: {nameof(menuItem.Name)} set ToolTip as '{text}'");
        }

        private void EnableMenuItem(ToolStripMenuItem tMenuItem, bool status) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { tMenuItem.Enabled = status; }));
            else
                tMenuItem.Enabled = status;
        }

        private void VisibleMenuItem(ToolStripMenuItem tMenuItem, bool status) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { tMenuItem.Visible = status; }));
            else
                tMenuItem.Visible = status;
        }

        private void VisibleMenuSeparator(ToolStripSeparator separator, bool status) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { separator.Visible = status; }));
            else
                separator.Visible = status;
        }

        private string ReturnMenuItemText(ToolStripMenuItem menuItem) //access from other threads
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

        private void VisibleControl(Control control, bool state) //access from other threads
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

        private void EnableControl(Control control, bool state) //access from other threads
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

        private void DisposeControl(Control control) //access from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    if (control != null)
                    {
                        groupBoxProperties?.Controls?.Remove(control);
                        control?.Dispose();
                    }
                }));
            else
            {
                if (control != null)
                {
                    groupBoxProperties?.Controls?.Remove(control);
                    control?.Dispose();
                }
            }
        }

        private void ChangeControlBackColor(Control control, Color color) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { control.BackColor = color; }));
            else
                control.BackColor = color; ;
        }

        private void SetControlToolTip(Control control, string text)
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    toolTip1.SetToolTip(control, text);
                }));
            else
            {
                toolTip1.SetToolTip(control, text);
            }

            AddLoggerTraceText($"{nameof(SetControlToolTip)}: {nameof(control.Name)} set text as '{text}'");
        }

        private void SetControlText(Control control, string text)
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    control.Text = text;
                }));
            else
            {
                control.Text = text;
            }

            AddLoggerTraceText($"{nameof(SetControlText)}: {nameof(control.Name)} set ToolTip as '{text}'");
        }

        private void SetCheckBoxCheckedStatus(CheckBox checkBox, bool checkboxChecked) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { checkBox.Checked = checkboxChecked; }));
            else
                checkBox.Checked = checkboxChecked;
        }

        private bool ReturnCheckboxCheckedStatus(CheckBox checkBox) //add string into  from other threads
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

        private void ResumePpanel(Panel panel) //access from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    panel?.ResumeLayout();
                }));
            else
            {
                panel?.ResumeLayout();
            }
        }

        private void SetPanelAutoSizeMode(Panel panel, AutoSizeMode state) //access from other threads
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

        private void SetPanelAutoScroll(Panel panel, bool state) //access from other threads
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

        private void SetPanelAnchor(Panel panel, AnchorStyles anchorStyles) //access from other threads
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

        private void SetPanelHeight(Panel panel, int height) //access from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    panel.Height = height;
                }));
            else
            { panel.Height = height; }
        }

        private int ReturnPanelParentHeight(Panel panel) //access from other threads
        {
            if (panel == null) return 0;

            int height = 0;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    if (panel?.Parent?.Height > 0)
                    { height = panel.Parent.Height; }
                }));
            else
            {
                if (panel?.Parent?.Height > 0)
                { height = panel.Parent.Height; }
            }
            return height;
        }

        private int ReturnPanelHeight(Panel panel) //access from other threads
        {
            if (panel == null) return 0;

            int height = 0;
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate
                {
                    if (panel?.Height > 0)
                    { height = panel.Height; }
                }));
            }
            else
            {
                if (panel?.Height > 0)
                { height = panel.Height; }
            }
            return height;
        }

        private int ReturnPanelWidth(Panel panel) //access from other threads
        {
            if (panel == null) return 0;

            int width = 0;
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate
                {
                    if (panel?.Width > 0)
                    { width = panel.Width; }
                }));
            }
            else
            {
                if (panel?.Width > 0)
                { width = panel.Width; }
            }
            return width;
        }

        private int ReturnPanelControlsCount(Panel panel) //access from other threads
        {
            if (panel == null) return 0;

            int count = 0;
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate
                  {
                      if (panel?.Controls?.Count > 0)
                          count = panel.Controls.Count;
                  }));
            }
            else
            {
                if (panel?.Controls?.Count > 0)
                    count = panel.Controls.Count;
            }
            return count;
        }

        private void RefreshPictureBox(PictureBox picBox, Bitmap picImage) // не работает
        {
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate
                 {
                     if (picBox != null)
                     {
                         picBox.Image = RefreshBitmap(picImage, ReturnPanelWidth(panelView) - 2, ReturnPanelHeight(panelView) - 2); //сжатая картина
                         picBox.Refresh();
                     }
                 }));
            }
            else
            {
                if (picBox != null)
                {
                    picBox.Image = RefreshBitmap(picImage, ReturnPanelWidth(panelView) - 2, ReturnPanelHeight(panelView) - 2); //сжатая картина
                    picBox.Refresh();
                }
            }
        }

        /// <summary>
        /// Change a Color of the Font on Status by the Timer
        /// stimerCurr = "Ждите!"
        /// stimerPrev - Информация СтатусТулСтрип
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Timer1_Tick(object sender, EventArgs e)
        {
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate
                 {
                     if (StatusLabel2.ForeColor == Color.DarkBlue)
                     {
                         StatusLabel2.ForeColor = Color.DarkRed;
                         if (!string.IsNullOrEmpty(stimerCurr)) StatusLabel2.Text = stimerCurr;
                     }
                     else
                     {
                         StatusLabel2.ForeColor = Color.DarkBlue;
                         StatusLabel2.Text = stimerPrev;
                     }
                 }));
            }
            else
            {
                if (StatusLabel2.ForeColor == Color.DarkBlue)
                {
                    StatusLabel2.ForeColor = Color.DarkRed;
                    if (!string.IsNullOrEmpty(stimerCurr)) StatusLabel2.Text = stimerCurr;
                }
                else
                {
                    StatusLabel2.ForeColor = Color.DarkBlue;
                    StatusLabel2.Text = stimerPrev;
                }
            }
        }

        /// <summary>
        /// Rise Value of progressBar at 1
        /// </summary>
        private void ProgressWork1Step()
        {
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate
                  {
                      if (ProgressBar1.Value > 99)
                      { ProgressBar1.Value = 0; }
                      ProgressBar1.Maximum = 100;
                      ProgressBar1.Value += 1;
                  }));
            }
            else
            {
                if (ProgressBar1.Value > 99)
                { ProgressBar1.Value = 0; }
                ProgressBar1.Maximum = 100;
                ProgressBar1.Value += 1;
            }
        }

        /// <summary>
        /// Run timer, set value of progressbar at 0, set StatusLabel2.BackColor = SystemColors.Control
        /// </summary>
        private void ProgressBar1Start() //Set progressBar Value into 0 from other threads
        {
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate
               {
                   timer1.Enabled = true;
                   ProgressBar1.Value = 0;
                   StatusLabel2.BackColor = SystemColors.Control;
               }));
            }
            else
            {
                timer1.Enabled = true;
                ProgressBar1.Value = 0;
                StatusLabel2.BackColor = SystemColors.Control;
            }
        }

        /// <summary>
        /// Stop timer, set progressbar at 100, set StatusLabel2.ForeColor =  Color.Black
        /// </summary>
        private void ProgressBar1Stop() //Set progressBar Value into 100 from other threads
        {
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate
                {
                    timer1.Stop();
                    StatusLabel2.Text = stimerPrev;
                    ProgressBar1.Value = 100;
                    StatusLabel2.ForeColor = Color.Black;
                }));
            }
            else
            {
                timer1.Stop();
                StatusLabel2.Text = stimerPrev;
                ProgressBar1.Value = 100;
                StatusLabel2.ForeColor = Color.Black;
            }
        }

        //---- End. Access to Controls from other threads ----//

        //---- Start. Convertors of data types ----//
        /// <summary>
        /// result is total_seconds
        /// </summary>
        /// <param name="hours"></param>
        /// <param name="minutes"></param>
        /// <returns></returns>
        private int ConvertStringsTimeToSeconds(string hours, string minutes)
        {
            int[] result = ConvertStringsTimeToInt(hours, minutes);
            return result[0] * 60 * 60 + result[1] * 60;
        }

        /// <summary>
        ///  result is 'hh:MM'
        /// </summary>
        /// <param name="hours"></param>
        /// <param name="minutes"></param>
        /// <returns></returns>
        private static string ConvertStringsTimeToStringHHMM(string hours, string minutes)
        {
            int[] result = ConvertStringsTimeToInt(hours, minutes);
            return $"{result[0]:d2}:{result[1]:d2}";
        }

        /// <summary>
        /// result int[]{ hours, minutes }
        /// </summary>
        /// <param name="hours"></param>
        /// <param name="minutes"></param>
        /// <returns></returns>
        private static int[] ConvertStringsTimeToInt(string hours, string minutes)
        {
            int.TryParse(hours ?? "0", out int h);
            int.TryParse(minutes ?? "0", out int m);
            return new int[] { h, m };
        }

        /// <summary>
        ///  result is 'hh:MM'
        /// </summary>
        /// <param name="hours"></param>
        /// <param name="minutes"></param>
        /// <returns></returns>
        private static string ConvertDecimalTimeToStringHHMM(decimal hours, decimal minutes)
        {
            return $"{(int)hours:d2}:{(int)minutes:d2}";
        }

        /// <summary>
        /// result decimal[] { hour, minute, total_hours, total_minutes, total_seconds }
        /// </summary>
        /// <param name="timeInHHMM"></param>
        /// <returns></returns>
        private decimal[] ConvertStringTimeHHMMToDecimalArray(string timeInHHMM) //time HH:MM converted to decimal value
        {
            int[] time = timeInHHMM?.ConvertTimeIntoStandartTimeIntArray();
            decimal[] result = new decimal[5];

            result[0] = (decimal)time[0];                                   // hour in decimal          22
            result[1] = (decimal)time[1];                                   // Minute in decimal        15
            result[2] = result[0] +                                          // hours in decimal         22.25
                (decimal)TimeSpan.FromMinutes((double)time[1]).TotalHours +
                (decimal)TimeSpan.FromSeconds((double)time[2]).TotalHours;
            result[3] = 60 * result[0] + result[1];                         // minutes in decimal       1335
            result[4] = 60 * 60 * result[0] + 60 * result[1] + (decimal)time[2];    // result in seconds       1335

            return result;
        }

        //---- End. Convertors of data types ----//

        private void SetStatusLabelText(object sender, TextEventArgs e)
        { SetStatusLabelText(StatusLabel2, e.Message); }

        private void SetStatusLabel2BackColor(object sender, ColorEventArgs e)
        { SetStatusLabelBackColor(StatusLabel2, e.Color); }

        private void SetUploadingStatus(object sender, BoolEventArgs e)
        { isUploaded = e.Status; }

        private void AddLoggerTraceText(object sender, TextEventArgs e)
        { logger.Trace(e.Message); }

        private void AddLoggerTraceText(string e)
        { logger.Trace(e); }

        private void CreateDBItem_Click(object sender, EventArgs e)
        { TryMakeLocalDB(); }

        private void GetCurrentSchemeItem_Click(object sender, EventArgs e)
        { GetSQLiteDbScheme(); }

        private void GetSQLiteDbScheme()
        {
            StringBuilder sb = new StringBuilder();
            OpenFileDialogExtentions filePath = new OpenFileDialogExtentions("Выберите файл c базой SQLite", "SQLite DB (main.db)|main.db|Все files (*.*)|*.*");
            string fpath = filePath.FilePath;

            if (fpath == null)
                fpath = dbApplication.FullName.ToString();

            if (fpath == null)
                return;

            Cursor = Cursors.WaitCursor;
            System.Threading.Thread worker = new System.Threading.Thread(new System.Threading.ThreadStart(delegate
            {
                try
                {
                    SQLiteConnectionStringBuilder builder = new SQLiteConnectionStringBuilder
                    {
                        DataSource = fpath,
                        PageSize = 4096,
                        UseUTF16Encoding = true
                    };

                    using (SQLiteConnection conn = new SQLiteConnection(builder.ConnectionString))
                    {
                        conn.Open();

                        //    SQLiteCommand count = new SQLiteCommand(@"SELECT COUNT(*) FROM SQLITE_MASTER", conn);
                        //   long num = (long)count.ExecuteScalar();

                        SQLiteCommand query = new SQLiteCommand(@"SELECT * FROM SQLITE_MASTER", conn);
                        using (SQLiteDataReader reader = query.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string type = (string)reader["type"];
                                string name = (string)reader["name"];
                                string tblName = (string)reader["tbl_name"];

                                // Ignore SQLite internal indexes and tables
                                if (name.StartsWith("sqlite_"))
                                    continue;
                                if (reader["sql"] == DBNull.Value)
                                    continue;

                                string sql = (string)reader["sql"];

                                MethodInvoker mi = delegate
                                { sb.Append(sql + ";\r\n\r\n"); };
                                this.Invoke(mi);
                            } // while
                        } // using
                    } // using

                    MethodInvoker mi3 = delegate
                    {
                        MessageBox.Show(this,
                            sb.ToString(),
                            $"Схема SQLite БД: {fpath}",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        logger.Info($"Схема SQLite БД: {fpath}");
                        logger.Info(sb.ToString());
                    };
                    this.Invoke(mi3);
                }
                catch (Exception ex)
                {
                    MethodInvoker mi1 = delegate
                    {

                        logger.Warn($"Ошибка получения схемы БД: {fpath}");
                        logger.Warn(ex.ToString());

                        MessageBox.Show(this,
                            ex.Message,
                            $"Ошибка получения схемы БД {fpath}",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                    };
                    this.Invoke(mi1);
                }
                finally
                {
                    MethodInvoker mi2 = delegate
                    {
                        Cursor = Cursors.Default;
                    };
                    this.Invoke(mi2);
                }
            }));
            worker.Start();
        }

        private void RunUpdate(UpdatingParameters parameters, IADUser user)
        {
            logger.Trace($"Update URL: {parameters.appUpdateURL}");

            //https://github.com/ravibpatel/AutoUpdater.NET
            //http://www.cyberforum.ru/csharp-beginners/thread2169711.html

            //Basic Authetication for XML, Update file and Change Log
            // BasicAuthentication basicAuthentication = new BasicAuthentication("myUserName", "myPassword");
            // AutoUpdater.BasicAuthXML = AutoUpdater.BasicAuthDownload = AutoUpdater.BasicAuthChangeLog = basicAuthentication;

            // AutoUpdater.CheckForUpdateEvent += AutoUpdaterOnCheckForUpdateEvent; //check manualy only
            // AutoUpdater.ReportErrors = true; // will show error message, if there is no update available or if it can't get to the XML file from web server.
            // AutoUpdater.CheckForUpdateEvent -= AutoUpdaterOnAutoCheckForUpdateEvent;
            // AutoUpdater.RemindLaterTimeSpan = RemindLaterFormat.Minutes;
            // AutoUpdater.RemindLaterTimeSpan = RemindLaterFormat.Minutes;
            // AutoUpdater.RemindLaterAt = 1;
            // AutoUpdater.ApplicationExitEvent += ApplicationExit;
            //  AutoUpdater.ApplicationExitEvent += AutoUpdater_ApplicationExitEvent;    //https://archive.codeplex.com/?p=autoupdaterdotnet
            // AutoUpdater.BasicAuthChangeLog

            AutoUpdater.AppTitle = "ASTA's update page";
            AutoUpdater.RunUpdateAsAdmin = false;
            AutoUpdater.DownloadPath = appFolderUpdatePath;
            AutoUpdater.CheckForUpdateEvent += CheckUpdate_Event; //write errors if had no access to the folder
            AutoUpdater.ApplicationExitEvent += ApplicationExit;    //https://archive.codeplex.com/?p=autoupdaterdotnet

            AutoUpdater.Start(parameters.appUpdateURL, new System.Net.NetworkCredential(user.Login, user.Password, user.Domain));

            //AutoUpdater.Start("ftp://kv-sb-server.corp.ais/Common/ASTA/ASTA.xml", new NetworkCredential("FtpUserName", "FtpPassword")); //download from FTP
        }

        /// <summary>
        /// Run Update immidiately
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RunUpdateItem_Click(object sender, EventArgs e)
        {
            IADUser user;
            if (!uploadingStatus)
            {
                UpdatingParameters parameters = MakeParametersToUpdate();
                user = ReturnADUserFromLocalDb();

                //Changing settings
                AutoUpdater.Mandatory = true;
                AutoUpdater.ReportErrors = true;
                AutoUpdater.UpdateMode = Mode.Normal;

                RunUpdate(parameters, user);

                AutoUpdater.CheckForUpdateEvent -= CheckUpdate_Event;
                AutoUpdater.ApplicationExitEvent -= ApplicationExit;
            }
            else
            {
                SetStatusLabelText(StatusLabel2, "Ждите! На сервер загружается новая версия ПО");
            }
        }

        private async Task CheckUpdates()
        {
            IADUser user;
            //Check updates frequently
            System.Timers.Timer timer = new System.Timers.Timer
            {
                Interval = 2 * 60 * 60 * 1000,       // the interval of checking is set in 2 hours='2 * 60 * 60 * 1000'
                SynchronizingObject = this
            };
            timer.Elapsed += delegate
            {
                if (!uploadingStatus)
                {
                    UpdatingParameters parameters = MakeParametersToUpdate();
                    user = ReturnADUserFromLocalDb();

                    AutoUpdater.Mandatory = true;
                    AutoUpdater.UpdateMode = Mode.ForcedDownload;
                    AutoUpdater.LetUserSelectRemindLater = false;
                    AutoUpdater.RemindLaterTimeSpan = RemindLaterFormat.Days;
                    AutoUpdater.RemindLaterAt = 2;

                    RunUpdate(parameters, user);
                }
                else
                { logger.Trace("Обновление приостановлено. На сервер сейчас загружается новая версия ПО"); }
            };
            await Task.Run(() =>
            {
                timer.Start();
            });
        }

        private void CheckUpdate_Event(UpdateInfoEventArgs args)
        {
            if (args != null)
            {
                if (args.IsUpdateAvailable)
                {
                    urlUpdateReachError = false;
                    try
                    {
                        if (AutoUpdater.DownloadUpdate())
                        {
                            SetStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                            UpdatingParameters parameters = MakeParametersToUpdate();

                            System.Xml.XmlDocument xmldoc = new System.Xml.XmlDocument();
                            System.Xml.XmlNodeList xmlnode;
                            xmldoc.Load(parameters.appUpdateURL);
                            xmlnode = xmldoc.GetElementsByTagName("version");
                            string foundNewVersionApp = xmlnode[0].InnerText;

                            logger.Info("=-------------------------------------------------------=");
                            logger.Info("");
                            logger.Trace("-= Update =-");
                            logger.Trace("...");
                            SetStatusLabelText(
                                StatusLabel2,
                                $"Oбнаружена новая версия {appName} ver. {foundNewVersionApp}");
                            logger.Trace("...");
                            logger.Trace("-= Update =-");
                            logger.Info("");
                            logger.Info("=-------------------------------------------------------=");

                            if (replaceBrokenRemoteFolderUpdateURL && urlUpdateReachError)
                            {
                                string message = SaveParameterInConfigASTA(new ParameterConfig()
                                {
                                    name = "RemoteFolderUpdateURL",
                                    value = remoteFolderUpdateURL,
                                    description = $"Параметр обновлен {DateTime.Now.ToYYYYMMDDHHMMSS()}",
                                    isSecret = false,
                                    isExample = "no"
                                });

                                System.Threading.Thread.Sleep(200);
                            }   //update broken RemoteFolderUpdateURL by correct URL

                            ApplicationExit();
                        }
                    }
                    catch (Exception exception)
                    { logger.Warn($"Update's check was failed: {exception.Message} | {exception.GetType().ToString()}"); }
                    // Uncomment the following line if you want to show standard update dialog instead.
                    // AutoUpdater.ShowUpdateForm();
                }
                else
                {
                    SetStatusLabelBackColor(StatusLabel2, SystemColors.Control);
                    SetStatusLabelText(StatusLabel2, $"Новых версий ПО '{appName}' не обнаружено");
                }
            }
            else
            {
                SetStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                SetStatusLabelText(StatusLabel2, $"Обновление не возможно. Проверьте URL: {remoteFolderUpdateURL}");
                logger.Warn($"Не верный URL сервера обновлений: {remoteFolderUpdateURL}");

                urlUpdateReachError = true;
            }
        }

        private bool urlUpdateReachError = false;

        private UpdatingParameters PrepareUpdating()
        {
            logger.Info(remoteFolderUpdateURL);

            //Make an archive with the currrent app's version
            MakeZip(appAllFiles, fileNameToUpdateZip);

            //Make MD5 of ZIP archive
            appFileMD5 = CalculateHash(fileNameToUpdateZip);

            UpdatingParameters parameters = new UpdatingParameters
            {
                localFolderUpdatingURL = localAppFolderPath,
                remoteFolderUpdatingURL = remoteFolderUpdateURL,
                appVersion = appVersionAssembly,
                appFileXml = fileNameToUpdateXML,
                appUpdateMD5 = appFileMD5,
                appFileZip = fileNameToUpdateZip
            };

            MakerOfLinks makerLinks = new MakerOfLinks();
            makerLinks.status += AddLoggerTraceText;

            MakerOfUpdateXmlFile makerXML = new MakerOfUpdateXmlFile();
            makerXML.status += AddLoggerTraceText;

            UpdatePreparing preparing = new UpdatePreparing(makerLinks, makerXML, parameters);
            preparing.status += SetStatusLabelText;

            preparing.Do();

            makerXML.status -= AddLoggerTraceText;
            makerLinks.status -= AddLoggerTraceText;
            preparing.status -= SetStatusLabelText;
            makerXML = null;
            makerLinks = null;
            parameters = null;

            return preparing.GetParameters();
        }

        //Upload App's files to Server
        private void UploadApplicationItem_Click(object sender, EventArgs e) //Uploading()
        {
            firstAttemptsUpdate = true;
            UploadUpdatinfAplication();
        }

        private async void UploadUpdatinfAplication()
        {
            string temp = ReturnStatusLabelText(StatusLabel1);
            SetStatusLabelText(StatusLabel1, "Обновление выгружается!");
            timer1.Start();
            await Task.Run(() => Uploading());

            if (replaceBrokenRemoteFolderUpdateURL && isUploaded)
            {
                string message = SaveParameterInConfigASTA(new ParameterConfig()
                {
                    name = "RemoteFolderUpdateURL",
                    value = remoteFolderUpdateURL,
                    description = "Параметр обновлен " + DateTime.Now.ToYYYYMMDDHHMMSS(),
                    isSecret = false,
                    isExample = "no"
                });
                SetStatusLabelText(StatusLabel2, message);
            }
            timer1.Stop();

            stimerCurr = temp;
            SetStatusLabelText(StatusLabel1, temp);
            StatusLabel2.ForeColor = Color.Black;
        }

        public System.IO.FileInfo ReturnNewFileInfo(string filePath)
        {
            //System.IO.Abstractions.FileInfoBase;
            return new System.IO.FileInfo(filePath);
        }

        private void Uploading() //UploadApplicationToShare()
        {
            SetStatusLabelBackColor(StatusLabel2, SystemColors.Control);

            uploadingStatus = true;
            isUploaded = false;

            UpdatingParameters parameters = PrepareUpdating();

            List<string> source = new List<string> {
                parameters.localFolderUpdatingURL + @"\" + parameters.appFileXml,
                parameters.localFolderUpdatingURL + @"\" + parameters.appFileZip
            };

            List<System.IO.Abstractions.IFileInfo> listSource = new List<System.IO.Abstractions.IFileInfo>();
            source.ForEach(p => listSource.Add((System.IO.Abstractions.FileInfoBase)ReturnNewFileInfo(p)));

            List<string> target = new List<string> {
               parameters.appUpdateFolderURI + parameters.appFileXml,
                parameters.appUpdateFolderURI + parameters.appFileZip
            };
            List<System.IO.Abstractions.IFileInfo> listTarget = new List<System.IO.Abstractions.IFileInfo>();
            target.ForEach(p => listTarget.Add((System.IO.Abstractions.FileInfoBase)ReturnNewFileInfo(p)));

            using (Uploader uploader = new Uploader(parameters, listSource, listTarget))
            {
                uploader.StatusText += AddLoggerTraceText;
                uploader.StatusColor += SetStatusLabel2BackColor;
                uploader.StatusFinishedUploading += SetUploadingStatus;
                uploader.Upload();

                uploader.StatusText -= AddLoggerTraceText;
                uploader.StatusColor -= SetStatusLabel2BackColor;
                uploader.StatusFinishedUploading -= SetUploadingStatus;
            }

            if (!isUploaded && uploadingStatus && firstAttemptsUpdate)
            {
                ReplaceBrokenRemoteFolderUpdateURL();

                System.Threading.Thread.Sleep(200);

                firstAttemptsUpdate = false;
                Uploading();
            }

            uploadingStatus = false;

            foreach (var file in source)
            {
                if (System.IO.File.Exists(file))
                {
                    try
                    {
                        System.IO.File.Delete(file);
                        AddLoggerTraceText($"Файл удален: {file}");
                    }
                    catch (Exception expt)
                    {
                        AddLoggerTraceText($"Файл {file} не удален из-за {expt.ToString()}");
                    }
                }
            }

            if (!isUploaded)
            {
                SetStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                SetStatusLabelText(StatusLabel2, $"Загрузить обновление ver.{parameters.appVersion} на {parameters.remoteFolderUpdatingURL} не удалось");
            }
            else
            {
                SetStatusLabelBackColor(StatusLabel2, Color.PaleGreen);
                SetStatusLabelText(StatusLabel2, $"Обновление ver.{parameters.appVersion} загружено в {parameters.remoteFolderUpdatingURL}");
            }
            parameters = null;
        }

        private static bool firstAttemptsUpdate = true;
        private static bool isUploaded = false;

        private static bool replaceBrokenRemoteFolderUpdateURL = false;

        private string ChangeFormUrl(string serverUrl)
        {
            AddLoggerTraceText($"-= {nameof(ChangeFormUrl)} =-");
            string url;
            if (serverUrl.ToLower().Contains(domain.ToLower()))//change url: server.domain -> server
            { url = $"{serverUrl.Remove(serverUrl.IndexOf('.'))}"; }
            else
            {
                if (serverUrl.Contains('/'))
                { url = $"{serverUrl.Remove(serverUrl.IndexOf('/'))}.{domain}"; }
                else if (serverUrl.Contains('\\'))
                { url = $"{serverUrl.Remove(serverUrl.IndexOf('\\'))}.{domain}"; }
                else
                { url = $"{serverUrl}.{domain}"; }
            }
            AddLoggerTraceText($"Old URL: {serverUrl} - > New URL: {url}");

            return url;
        }

        private void ReplaceBrokenRemoteFolderUpdateURL()
        {
            AddLoggerTraceText("-= ReplaceBrokenRemoteFolderUpdateURL =-");

            string oldUrl = remoteFolderUpdateURL;
            logger.Info($"old {nameof(remoteFolderUpdateURL)}: {oldUrl}");
            remoteFolderUpdateURL = ChangeFormUrl(oldUrl);
            logger.Info($"{nameof(remoteFolderUpdateURL)}: {oldUrl} -> {remoteFolderUpdateURL}");
            replaceBrokenRemoteFolderUpdateURL = true;
        }

        private UpdatingParameters MakeParametersToUpdate()
        {
            if (urlUpdateReachError)
            { ReplaceBrokenRemoteFolderUpdateURL(); }

            UpdatingParameters parameters = new UpdatingParameters
            {
                localFolderUpdatingURL = localAppFolderPath,
                remoteFolderUpdatingURL = remoteFolderUpdateURL,
                appVersion = appVersionAssembly,
                appFileXml = fileNameToUpdateXML,
                appUpdateMD5 = appFileMD5,
                appFileZip = fileNameToUpdateZip
            };

            MakerOfLinks makerLinks = new MakerOfLinks();
            makerLinks.SetParameters(parameters);
            makerLinks.Make();

            urlUpdateReachError = false;

            return makerLinks.GetParameters();
        }

        //Calculate MD5 checksum of local file
        private string CalculateHash(string filePath)
        {
            CalculatorHash calculatedHash = new CalculatorHash(filePath);
            return calculatedHash.Calculate();
        }

        private void CalculateHashItem_Click(object sender, EventArgs e) //Selectfiles()
        { SelectfilesForCalculatingHash(); }

        private void SelectfilesForCalculatingHash() //SelectFileOpenFileDialog() CalculateFileHash()
        {
            DialogResult selectTwoFiles = MessageBox.Show("Выбрать 2 файла для сравнения?", "Сравнение файлов",
                MessageBoxButtons.YesNo,
                      MessageBoxIcon.Question,
                      MessageBoxDefaultButton.Button1);

            OpenFileDialogExtentions ofde = new OpenFileDialogExtentions("Выберите файл для вычисления его хэша");
            string filePath = ofde.FilePath;

            CalculatorHash calculatedHash = new CalculatorHash(filePath);
            string result = calculatedHash.Calculate();

            if (selectTwoFiles == DialogResult.Yes)
            {
                ofde = new OpenFileDialogExtentions("Выберите следующий файл для вычисления его хэша");
                filePath = ofde.FilePath;

                calculatedHash = new CalculatorHash(filePath);
                string secondFile = calculatedHash.Calculate();
                bool equalString = string.Equals(result, secondFile);
                result += "\n" + secondFile;
                string caption = equalString ? "MD5 хэш файлов совпадает" : "MD5 хэш файлов не совпадает";
                MessageBox.Show(result, caption, MessageBoxButtons.OK, ReturnMessageBoxIcon(equalString));
            }
            else
            { MessageBox.Show(result, "Результат вычисления хэша", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        }

        private MessageBoxIcon ReturnMessageBoxIcon(bool ok)
        {
            if (ok) return MessageBoxIcon.Information;
            else return MessageBoxIcon.Exclamation;
        }

        private void OpenMenuItemsAsLocalAdmin_Click(object sender, EventArgs e)
        {
            VisibleOfAdminMenuItems(true);
            SetStatusLabelText(StatusLabel2, "Включено отображение меню для Администратора системы");
        }

        private void VisibleOfAdminMenuItems(bool show)
        {
            VisibleMenuItem(RefreshConfigInMainDBItem, show);
            VisibleMenuItem(CalculateHashItem, show);
            VisibleMenuItem(AddParameterInConfigItem, show);
            VisibleMenuItem(SelectedToLoadCityItem, show);
            VisibleMenuItem(MailingAddItem, show);
            VisibleMenuItem(TestToSendAllMailingsItem, show);
            VisibleMenuItem(UploadApplicationItem, show);
            VisibleMenuItem(GetADUsersItem, show);
            VisibleMenuItem(GetCurrentSchemeItem, show);
            VisibleMenuItem(CreateDBItem, show);
            VisibleMenuItem(ClearRegistryItem, show);
            VisibleMenuItem(EditAnualDaysItem, show);
            VisibleMenuItem(ModeItem, show);
            VisibleMenuItem(MailingsExceptItem, show);
            VisibleMenuItem(SettingsProgrammItem, show);
            VisibleMenuItem(DeletePersonFromGroupItem, show);
            VisibleMenuItem(ImportPeopleInLocalDBItem, show);
            VisibleMenuItem(OpenMenuAsLocalAdminItem, !show);

            VisibleMenuSeparator(toolStripSeparator2, !show);
        }

        private void EnableMainMenuItems(bool show)
        {
            EnableMenuItem(FunctionMenuItem, show);
            EnableMenuItem(GroupsMenuItem, show);
            EnableMenuItem(SettingsMenuItem, show);
            EnableMenuItem(HelpMenuItem, show);

            CheckBoxesFiltersAll_Enable(show);
        }
    }
}