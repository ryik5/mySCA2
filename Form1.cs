using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.Data.Common;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Security.Cryptography;  // for Crypography

using MimeKit;

using ASTA.Classes;
using ASTA.Classes.People;
using ASTA.Classes.Common;
using ASTA.Classes.Updating;
using ASTA.Security;
using AutoUpdaterDotNET;

//using NLog;
//Project\Control NuGet\console 
//install-package nlog
//install-package nlog.config


namespace ASTA
{
    public partial class WinFormASTA : Form
    {
        //todo!!!!!!!!!
        //Check of All variables, const and controls
        //they will be needed to Remove if they are not needed
        private DataGridViewOperations dgvo;

        //datatime
        DateTime today = DateTime.Now;

        //logging
        static NLog.Logger logger;
        static string method = System.Reflection.MethodBase.GetCurrentMethod().Name;
        // logger.Trace("-= " + method + " =-");

        //System settings
        readonly static string appVersionAssembly = System.Reflection.Assembly.GetEntryAssembly().GetName().Version.ToString();
        readonly static System.Diagnostics.FileVersionInfo appFileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath);
        readonly static string appCopyright = appFileVersionInfo.LegalCopyright;
        readonly static string appName = appFileVersionInfo.ProductName;
        readonly static string appNameXML = appName + @".xml";
        readonly static string appFileZip = appName + @".zip";
        static string appFileMD5;

        static string remoteFolderUpdateURL;// = @"kv-sb-server.corp.ais\Common\ASTA";
       // static string appUpdateFolderURL;// = @"file://" + serverUpdateURL.Replace(@"\", @"/") + @"/"; //  @"file://kv-sb-server.corp.ais/Common/ASTA/";
        static string appUpdateURL;// = appUpdateFolderURL + appNameXML;
      //  static string appUpdateFolderURI;// = @"\\" + serverUpdateURL + @"\"; //@"\\kv-sb-server.corp.ais\Common\ASTA\";

        static bool uploadingUpdate = false;


        readonly static string appFilePath = Application.ExecutablePath;
        readonly static string localAppFolderPath = System.IO.Path.GetDirectoryName(appFilePath); //Environment.CurrentDirectory
        readonly static string appFolderTempPath = System.IO.Path.Combine(localAppFolderPath, "Temp");
        readonly static string appFolderUpdatePath = System.IO.Path.Combine(localAppFolderPath, "Update");
        readonly static string appFolderBackUpPath = System.IO.Path.Combine(localAppFolderPath, "Backup");
        readonly static string[] appAllFiles = new string[] {
                @"ASTA.exe" , @"NLog.config", @"AutoUpdater.NET.dll",
                @"ASTA.sql", @"Google.Protobuf.dll", @"NLog.dll",
                @"MySql.Data.dll", @"MySql.Data.dll", @"Renci.SshNet.dll",
                @"MailKit.dll", @"MimeKit.dll",
                @"BouncyCastle.Crypto.dll", @"System.Data.SQLite.dll",
                @"x64\SQLite.Interop.dll", @"x86\SQLite.Interop.dll",
                
                //Analysing packages
                @"ASTA.exe.config",
                        };
        static string appQueryCreatingDB = System.IO.Path.Combine(localAppFolderPath, System.IO.Path.GetFileNameWithoutExtension(appFilePath) + @".sql");

        string guid = System.Runtime.InteropServices.Marshal.GetTypeLibGuidForAssembly(System.Reflection.Assembly.GetExecutingAssembly()).ToString(); // получаем GIUD приложения// получаем GIUD приложения

        readonly static string appRegistryKey = @"SOFTWARE\RYIK\ASTA";
        readonly static byte[] keyEncryption = Convert.FromBase64String(@"OCvesunvXXsxtt381jr7vp3+UCwDbE4ebdiL1uinGi0="); //Key Encrypt
        readonly static byte[] keyDencryption = Convert.FromBase64String(@"NO6GC6Zjl934Eh8MAJWuKQ=="); //Key Decrypt

        readonly static System.IO.FileInfo dbApplication = new System.IO.FileInfo(@".\main.db");
        readonly static string appDbPath = dbApplication.FullName;
        readonly static string appDbName = System.IO.Path.GetFileName(appDbPath);

        readonly static string sqLiteLocalConnectionString = string.Format("Data Source = {0}; Version=3;", dbApplication); ////$"Data Source={dbApplication.FullName};Version=3;"

        static string sqlServerConnectionString;// = "Data Source=" + serverName + "\\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID=" + userName + ";Password=" + userPasswords + "; Connect Timeout=5";
        static string mysqlServerConnectionStringDB1;//@"server=" + mysqlServer + @";User=" + mysqlServerUserName + @";Password=" + mysqlServerUserPassword + @";database=wwwais;convert zero datetime=True;Connect Timeout=60";
        static string mysqlServerConnectionStringDB2;//@"server=" + mysqlServer + @";User=" + mysqlServerUserName + @";Password=" + mysqlServerUserPassword + @";database=wwwais;convert zero datetime=True;Connect Timeout=60";


        //todo
        //Все константы в локальную БД
        //todo
        //преобразовать в дикшенери сущностей с дефолтовыми значениями и описаниями параметров
        //  Dictionary<string, ParameterOfConfiguration> config = new Dictionary<string, ParameterOfConfiguration>();

        static string sLastSelectedElement = "MainForm";
        static string nameOfLastTable = "PersonRegistrationsList";
        static string currentAction = "";
        static bool currentModeAppManual = true;

        static string statusBar = appName + " ver." + appVersionAssembly + " by " + appCopyright;

        // taskbar and logo
        static Bitmap bmpLogo = Properties.Resources.LogoRYIK;
        NotifyIcon notifyIcon = new NotifyIcon();
        ContextMenu contextMenu;
        bool buttonAboutForm;
        static Byte[] byteLogo;


        int iCounterLine = 0;

        //collecting of data
        static List<Employee> listFIO = new List<Employee>(); // List of FIO and identity of data

        //Controls "NumUpDown"
        decimal numUpHourStart = 9;
        decimal numUpMinuteStart = 0;
        decimal numUpHourEnd = 18;
        decimal numUpMinuteEnd = 0;


        //View of registrations
        PictureBox pictureBox1 = new PictureBox();
        Bitmap bmp = new Bitmap(1, 1);


        //reports
        static int offsetTimeIn = 60;    // смещение времени прихода после учетного, в секундах, в течении которого не учитывается опоздание
        static int offsetTimeOut = 60;   // смещение времени ухода до учетного, в секундах в течении которого не учитывается ранний уход
        string[] myBoldedDates;
        string[] workSelectedDays;
        static string reportStartDay = "";
        static string reportLastDay = "";
        bool reportExcelReady = false;
        string filePathExcelReport;

        //mailing
        const string NAME_OF_SENDER_REPORTS = @"Отдел компенсаций и льгот";
        const string START_OF_MONTH = @"START_OF_MONTH";
        const string MIDDLE_OF_MONTH = @"MIDDLE_OF_MONTH";
        const string END_OF_MONTH = @"END_OF_MONTH";
        static System.Threading.Timer timer;
        static object synclock = new object();
        static bool sent = false;
        static string DEFAULT_DAY_OF_SENDING_REPORT = @"END_OF_MONTH";
        static int ShiftDaysBackOfSendingFromLastWorkDay = 3; //shift back of sending email before a last working day within the month

        //Page of Mailing
        Label labelMailServerName;
        TextBox textBoxMailServerName;
        Label labelMailServerUserName;
        TextBox textBoxMailServerUserName;
        Label labelMailServerUserPassword;
        TextBox textBoxMailServerUserPassword;
        static string mailServer = "";
        static string mailServerDB = "";
        static int mailServerSMTPPort = 25;
        static string mailServerSMTPPortDB = "";
        static string mailSenderAddress = "";
        static string mailsOfSenderOfNameDB = "";
        static string mailsOfSenderOfPassword = "";
        static string mailsOfSenderOfPasswordDB = "";

        static MailServer _mailServer;
        static MailUser _mailUser;
        static string mailJobReportsOfNameOfReceiver = ""; //Receiver of job reports
        static List<Mailing> resultOfSendingReports = new List<Mailing>();
        //  static System.Net.Mail.LinkedResource mailLogo;


        //Page of "Settings' Programm"
        bool bServer1Exist = false;
        Label labelServer1;
        TextBox textBoxServer1;
        Label labelServer1UserName;
        TextBox textBoxServer1UserName;
        Label labelServer1UserPassword;
        TextBox textBoxServer1UserPassword;
        string sServer1;
        string sServer1Registry;
        string sServer1DB;
        string sServer1UserName;
        string sServer1UserNameRegistry;
        string sServer1UserNameDB;
        string sServer1UserPassword;
        string sServer1UserPasswordRegistry;
        string sServer1UserPasswordDB;

        Label labelmysqlServer;
        TextBox textBoxmysqlServer;
        Label labelmysqlServerUserName;
        TextBox textBoxmysqlServerUserName;
        Label labelmysqlServerUserPassword;
        TextBox textBoxmysqlServerUserPassword;
        static string mysqlServer;
        static string mysqlServerRegistry;
        static string mysqlServerDB;
        static string mysqlServerUserName;
        static string mysqlServerUserNameRegistry;
        static string mysqlServerUserNameDB;
        static string mysqlServerUserPassword;
        static string mysqlServerUserPasswordRegistry;
        static string mysqlServerUserPasswordDB;

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


        List<string> listGroups = new List<string>();
        int countGroups = 0;
        int countUsers = 0;
        int countMailers = 0;

        int numberPeopleInLoading = 1;

        //todo
        static string stimerPrev = "";
        static string stimerCurr = "Ждите!";

        static List<UserAD> usersAD = new List<UserAD>(); //Users of AD. Got data from Domain

        static List<OutReasons> outResons = new List<OutReasons>();
        static List<OutPerson> outPerson = new List<OutPerson>();
        static List<PeopleShift> peopleShifts = new List<PeopleShift>();

        //DataTables with people data
        DataTable dtPeople = new DataTable("People");
        DataColumn[] dcPeople =
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
        static DataTable dtPersonTemp = new DataTable("PersonTemp");
        static DataTable dtPersonTempAllColumns = new DataTable("PersonTempAllColumns");
        static DataTable dtPersonRegistrationsFullList = new DataTable("PersonRegistrationsFullList");
        static DataTable dtPeopleGroup = new DataTable("PeopleGroup");
        static DataTable dtPeopleListLoaded = new DataTable("PeopleLoaded");

        //Color of User's Control elements which depend on the selected MenuItem  
        Color labelGroupCurrentBackColor;
        Color textBoxGroupCurrentBackColor;
        Color labelGroupDescriptionCurrentBackColor;
        Color textBoxGroupDescriptionCurrentBackColor;
        Color comboBoxFioCurrentBackColor;
        Color textBoxFIOCurrentBackColor;
        Color textBoxNavCurrentBackColor;

        static CollectionOfPassagePoints collectionOfPassagePoints;

        public WinFormASTA()
        { InitializeComponent(); }

        private void Form1_Load(object sender, EventArgs e)
        { Form1Load(); }

        private async void Form1Load()
        {
            //Начало лога
            logger = NLog.LogManager.GetCurrentClassLogger();
            logger.Info("");
            logger.Info("");
            logger.Info("-= " + statusBar + " =-");
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


            //Clear temporary folder 
            ClearItemsInApplicationFolders(appFolderTempPath + @"\*.*");
            ClearItemsInApplicationFolders(appFolderUpdatePath + @"\*.*");
            System.IO.Directory.CreateDirectory(appFolderBackUpPath + @"\");


            //Make archives:
            //1. from app's *.exe and main lib files of the app
            ClearItemsInApplicationFolders(localAppFolderPath + @"\"+ appFileZip);
            MakeZip(appAllFiles, appFileZip);
            if (System.IO.File.Exists(appFileZip))
            {
                System.IO.File.Move(appFileZip, System.IO.Path.Combine(appFolderBackUpPath, appName + "." + GetSafeFilename(DateTime.Now.ToYYYYMMDDHHMMSS(), "") + @".zip"));
            }
            //refresh temp folder
            ClearItemsInApplicationFolders(appFolderTempPath + @"\*.*");

            //2. from the main DB of application
            string dbZipPath = appDbName + "." + GetSafeFilename(DateTime.Now.ToYYYYMMDDHHMMSS(), "") + @".zip";
            ClearItemsInApplicationFolders(localAppFolderPath + @"\" + dbZipPath);
            MakeZip(appDbName, dbZipPath);
            System.IO.File.Move(dbZipPath, System.IO.Path.Combine(appFolderBackUpPath, System.IO.Path.GetFileName(dbZipPath)));

            //refresh temp folder
            ClearItemsInApplicationFolders(appFolderTempPath + @"\*.*");


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

            string txtSQLs = String.Join(" ", ReadTXTFile(appQueryCreatingDB)); //List<string> -> a single string
            DbSchema schemaTXT = DbSchema.ParseSql(txtSQLs);
            logger.Trace("txtSQL is wanted to make: " + schemaTXT?.Tables?.Count + " tables from: " + appQueryCreatingDB);

            foreach (var table in schemaTXT.Tables)
            { logger.Trace("the table is wanted: " + table.Key); }

            bool equalDB = schemaTXT.Equals(schemaDB);
            if (equalDB) { logger.Trace("Схема конфигурации текущей БД соответствует схеме загруженной с файла: " + appQueryCreatingDB); }
            else { logger.Info("Схема конфигурации текущей БД Отличается от схеме загруженной с файла: " + appQueryCreatingDB); }

            logger.Trace("tables in loaded DB: " + schemaDB?.Tables?.Count + ", " + " must be tables: " + schemaTXT?.Tables?.Count);

            if (currentDbEmpty || !equalDB || !(schemaDB.Tables.Equals(schemaTXT.Tables)))
            {
                logger.Info("Заполняю схему локальной DB");
                TryMakeLocalDB(appQueryCreatingDB);
            }


            //Refresh Configuration of the application
            AddExceptedParametersIntoConfigurationDb();

            logger.Info("Загружаю/проверяю настройки программы...");

            LoadPreviouslySavedParameters();
            logger.Info("Вычисляю ближайшие праздничные и выходные дни...");
            DataTable dtEmpty = new DataTable();
            EmployeeFull personEmpty = new EmployeeFull();
            var startDay = DateTime.Now.AddDays(-60).ToYYYYMMDD();
            var endDay = DateTime.Now.AddDays(30).ToYYYYMMDD();

            SeekAnualDays(ref dtEmpty, ref personEmpty, false,
                ConvertStringDateToIntArray(startDay), ConvertStringDateToIntArray(endDay),
                ref myBoldedDates, ref workSelectedDays
                );
            dtEmpty = null;
            personEmpty = null;

            dgvo = new DataGridViewOperations();
            monthCalendar.SelectionStart = DateTime.Now;
            monthCalendar.SelectionEnd = DateTime.Now;
            monthCalendar.Update();
            monthCalendar.Refresh();

            //todo
            //rewrite to access from other threads
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

            logger.Info("Настраиваю переменные....");

            //Prepare DataTables
            dtPeople.Columns.AddRange(dcPeople);
            dtPeople.DefaultView.Sort = Names.GROUP + ", " + Names.FIO + ", " + Names.DATE_REGISTRATION + ", " + Names.TIME_REGISTRATION + ", " + Names.REAL_TIME_IN + ", " + Names.REAL_TIME_OUT + " ASC";
            // dtPeople.DefaultView.Sort = "[Группа], [Фамилия Имя Отчество], [Дата регистрации], [Время регистрации], [Фактич. время прихода ЧЧ:ММ:СС], [Фактич. время ухода ЧЧ:ММ:СС] ASC";

            //Clone default column name and structure from 'dtPeople' to other DataTables
            dtPersonTemp = dtPeople.Clone();  //Copy only structure(Name of columns)
            dtPersonRegistrationsFullList = dtPeople.Clone();  //Copy only structure(Name of columns)
            dtPeopleGroup = dtPeople.Clone();  //Copy only structure(Name of columns)

            if (currentDbEmpty)
            {
                logger.Warn("Form loading is finishing, but the local db is still empty!");
            }
            else
            {
                if (currentModeAppManual)
                {
                    nameOfLastTable = "ListFIO";
                    SeekAndShowMembersOfGroup("");
                    logger.Info("Программа запущена в интерактивном режиме...");
                }
                else
                {
                    // nameOfLastTable = "Mailing";
                    _ControlEnable(comboBoxFio, false);
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
                    Task.Run(() => AutoUpdate());
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
                ParameterOfConfigurationInSQLiteDB parametersInDb = new ParameterOfConfigurationInSQLiteDB(dbApplication);
                listParameters = parametersInDb.GetParameters("%%"); //load only real data
                List<ParameterConfig> parameters = ReturnListParametersWithEmptyValue(listParameters);
                if (parameters?.Count > 0)
                {
                    string resultParameters = null;
                    foreach (var p in parameters)
                    {
                        resultParameters += (p.parameterName + " is empty\n\r");
                    }

                    logger.Warn("Empty parameters in local config db: " + resultParameters);
                }
            }


            logger.Info("Настраиваю интерфейс....");
            bmpLogo = Properties.Resources.LogoRYIK;
            MakeByteLogo(bmpLogo); //logo for mailing

            this.Icon = Icon.FromHandle(bmpLogo.GetHicon());
            notifyIcon.Icon = this.Icon;
            notifyIcon.Visible = true;
            notifyIcon.BalloonTipText = "Developed by " + appCopyright;

            notifyIcon.ShowBalloonTip(500);

            this.Text = appFileVersionInfo.Comments;
            notifyIcon.Text = appName + "\nv." + appVersionAssembly + " (" + appFileVersionInfo.FileVersion + ")" + "\n" + appFileVersionInfo.CompanyName;

            StatusLabel1.Text = statusBar;
            StatusLabel1.Alignment = ToolStripItemAlignment.Right;

            contextMenu = new ContextMenu();  //Context Menu on notify Icon
            notifyIcon.ContextMenu = contextMenu;
            contextMenu.MenuItems.Add("About", AboutSoft);
            contextMenu.MenuItems.Add("-", AboutSoft);
            contextMenu.MenuItems.Add("Exit", ApplicationExit);


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
            dateTimePickerStart.Value = DateTime.Now.FirstDayOfMonth();
            dateTimePickerEnd.Value = DateTime.Now.LastDayOfMonth();

            if (_comboBoxCountItems(comboBoxFio) > 0)
            {
                _MenuItemVisible(listFioItem, true);
                _comboBoxSelectIndex(comboBoxFio, 0);
            }

            numUpDownHourStart.Value = 9;
            numUpDownMinuteStart.Value = 0;
            numUpDownHourEnd.Value = 18;
            numUpDownMinuteEnd.Value = 0;

            _MenuItemSetText(ModeItem, "Включить режим автоматических e-mail рассылок");
            _MenuItemSetTooltip(ModeItem, "Включен интерактивный режим. Все рассылки остановлены.");

            _MenuItemSetText(EditAnualDaysItem, Names.DAY_OFF_OR_WORK);
            _MenuItemSetTooltip(EditAnualDaysItem, Names.DAY_OFF_OR_WORK_EDIT);

            _MenuItemSetText(PersonOrGroupItem, Names.WORK_WITH_A_PERSON);

            SetMenuItemsTextAfterClosingDateTimePicker();

            _ControlSetToolTip(textBoxGroup, "Создать или добавить в группу");
            _ControlSetToolTip(textBoxGroupDescription, "Изменить описание группы");

            logger.Trace("SetTechInfoIntoDB");
            SetTechInfoIntoDB();

            logger.Info("");
            logger.Info("Загрузка и настройка интерфейса ПО завершена....");
            logger.Info("");
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
        {
            logger.Info("");
            logger.Info("");
            logger.Info("-=-=  Завершение работы ПО  =-=-");
            logger.Info("-----------------------------------------");
            logger.Info("");

            //taskkill /F /IM ASTA.exe
            Text = @"Closing application...";
            System.Threading.Thread.Sleep(1000);

            Application.Exit();
        }



        private List<string> ReadTXTFile(string fpath = null, int listMaxLength = 10000)
        {
            List<string> txt = new List<string>(listMaxLength);
            string query = string.Empty, s = string.Empty, result = string.Empty, log = string.Empty;

            Cursor = Cursors.WaitCursor;

            MethodInvoker mi = delegate
            {
                string prevText = _toolStripStatusLabelReturnText(StatusLabel2);
                Color prevColor = _toolStripStatusLabelReturnBackColor(StatusLabel2);
                bool readOk = true;

                if (!(fpath?.Length > 0))
                {
                    fpath = SelectFileOpenFileDialog("Выберите файл", "Текстовые файлы (*.txt)|*.txt|SQL файлы (*.sql)|*.sql|All files (*.*)|*.*");
                    if (fpath == null) return;
                }

                logger.Trace("ReadTXTFile");
                _toolStripStatusLabelSetText(StatusLabel2, "Читаю файл: " + fpath);
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
                    logger.Warn("Не могу прочитать файл: " + fpath + " \nException: " + err.ToString());
                }

                if (!readOk)
                {
                    _toolStripStatusLabelSetText(StatusLabel2, "Не могу прочитать файл: " + fpath, true);
                }
                else
                {
                    _toolStripStatusLabelBackColor(StatusLabel2, prevColor);
                    if (prevText?.Length > 0)
                    { _toolStripStatusLabelSetText(StatusLabel2, prevText); }
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

        private void TryMakeLocalDB(string fpath = null, DbSchema schema = null)
        {
            List<string> txt = ReadTXTFile(fpath);
            string query = string.Empty, result = string.Empty, log = string.Empty;

            _toolStripStatusLabelSetText(StatusLabel2, "Создаю таблицы в БД на основе запроса из текстового файла: " + fpath);
            using (SqLiteDbWriter dbWriter = new SqLiteDbWriter(sqLiteLocalConnectionString, dbApplication))
            {
                dbWriter.ExecuteQueryBegin();
                foreach (var s in txt)
                {
                    if (s.StartsWith("CREATE TABLE"))
                    { query = s.Trim(); }
                    else { query += s.Trim(); }

                    if (s.EndsWith(";"))
                    {
                        dbWriter.ExecuteQueryForBulkStepByStep(query);
                        log += "query: " + query + "\nresult: " + dbWriter.Status + "\n";
                    }
                }//foreach

                dbWriter.ExecuteQueryEnd();
                _toolStripStatusLabelSetText(StatusLabel2, "Таблицы в БД созданы.");
            }
        }


        private void SetTechInfoIntoDB() //Write Technical Info in DB 
        {
            string result = string.Empty;
            string query = null;
            if (dbApplication.Exists)
            {
                query = "INSERT OR REPLACE INTO 'TechnicalInfo' (PCName, POName, POVersion, LastDateStarted, CurrentUser, FreeRam, GuidAppication) " +
                        " VALUES (@PCName, @POName, @POVersion, @LastDateStarted, @CurrentUser, @FreeRam, @GuidAppication)";

                using (SqLiteDbWriter dbWriter = new SqLiteDbWriter(sqLiteLocalConnectionString, dbApplication))
                {
                    using (SQLiteCommand SqlQuery = new SQLiteCommand(query, dbWriter._sqlConnection))
                    {
                        SqlQuery.Parameters.Add("@PCName", DbType.String).Value = Environment.MachineName + "|" + Environment.OSVersion;
                        SqlQuery.Parameters.Add("@POName", DbType.String).Value = appFileVersionInfo.FileName + "(" + appName + ")";
                        SqlQuery.Parameters.Add("@POVersion", DbType.String).Value = appFileVersionInfo.FileVersion;
                        SqlQuery.Parameters.Add("@LastDateStarted", DbType.String).Value = DateTime.Now.ToYYYYMMDDHHMM();
                        SqlQuery.Parameters.Add("@CurrentUser", DbType.String).Value = Environment.UserName;
                        SqlQuery.Parameters.Add("@FreeRam", DbType.String).Value = "RAM: " + Environment.WorkingSet.ToString();
                        SqlQuery.Parameters.Add("@GuidAppication", DbType.String).Value = guid;

                        dbWriter.ExecuteQuery(SqlQuery);
                        result = dbWriter.Status;
                    }
                }
            }
            logger.Trace("SetTechInfoIntoDB: query: " + query + "\n" + result);//method = System.Reflection.MethodBase.GetCurrentMethod().Name;
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
            if (dbApplication.Exists)
            {
                string query;
                int count = 0;
                using (SqLiteDbReader dbReader = new SqLiteDbReader(sqLiteLocalConnectionString, dbApplication))
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
                                    _comboBoxAdd(comboBoxFio, record["ComboList"].ToString().Trim());
                                    count++;
                                }
                            }
                            catch (Exception err) { logger.Info(err.ToString()); }
                        }
                    }
                    logger.Trace("LoadPreviouslySavedParameters: query: " + query + "\n" + count + " rows loaded from 'LastTakenPeopleComboList'");

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
                    logger.Trace("LoadPreviouslySavedParameters: query:" + query + "\n" + count + " rows loaded from 'PeopleGroup'");
                }

                //loading parameters
                ParameterOfConfigurationInSQLiteDB parameters = new ParameterOfConfigurationInSQLiteDB(dbApplication);
                listParameters = parameters.GetParameters("%%").FindAll(x => x?.isExample == "no"); //load only real data

                DEFAULT_DAY_OF_SENDING_REPORT = GetValueOfConfigParameter(listParameters, @"DEFAULT_DAY_OF_SENDING_REPORT", END_OF_MONTH);
                int days = 0;
                int.TryParse(GetValueOfConfigParameter(listParameters, @"ShiftDaysBackOfSendingFromLastWorkDay", ""), out days);

                if (days < 0 || days > 27)
                { ShiftDaysBackOfSendingFromLastWorkDay = 3; }
                else
                { ShiftDaysBackOfSendingFromLastWorkDay = days; }

                clrRealRegistrationRegistry = Color.FromName(GetValueOfConfigParameter(listParameters, @"clrRealRegistration", "PaleGreen"));

                sServer1DB = GetValueOfConfigParameter(listParameters, @"SKDServer", null);
                sServer1UserNameDB = GetValueOfConfigParameter(listParameters, @"SKDUser", null);
                sServer1UserPasswordDB = GetValueOfConfigParameter(listParameters, @"SKDUserPassword", null, true);

                mysqlServerDB = GetValueOfConfigParameter(listParameters, @"MySQLServer", null);
                mysqlServerUserNameDB = GetValueOfConfigParameter(listParameters, @"MySQLUser", null);
                mysqlServerUserPasswordDB = GetValueOfConfigParameter(listParameters, @"MySQLUserPassword", null, true);

                mailServerDB = GetValueOfConfigParameter(listParameters, @"MailServer", null);
                mailServerSMTPPortDB = GetValueOfConfigParameter(listParameters, @"MailServerSMTPport", null);
                mailsOfSenderOfNameDB = GetValueOfConfigParameter(listParameters, @"MailUser", null);
                mailsOfSenderOfPasswordDB = GetValueOfConfigParameter(listParameters, @"MailUserPassword", null, true);

                mailJobReportsOfNameOfReceiver = GetValueOfConfigParameter(listParameters, @"JobReportsReceiver", null, true);
                string defaultURL = remoteFolderUpdateURL;
                remoteFolderUpdateURL = GetValueOfConfigParameter(listParameters, @"RemoteFolderUpdateURL", defaultURL);
            }

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


                sqlServerConnectionString = "Data Source=" + sServer1 + "\\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID=" + sServer1UserName + ";Password=" + sServer1UserPassword + "; Connect Timeout=30";
                mysqlServerConnectionStringDB1 = @"server=" + mysqlServer + @";User=" + mysqlServerUserName + @";Password=" + mysqlServerUserPassword + @";database=wwwais;convert zero datetime=True;Connect Timeout=60";
                mysqlServerConnectionStringDB2 = @"server=" + mysqlServer + @";User=" + mysqlServerUserName + @";Password=" + mysqlServerUserPassword + @";database=wwwais;convert zero datetime=True;Connect Timeout=60";


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

            logger.Trace("LoadPreviouslySavedParameters: " + nameof(modeApp) + " - " + modeApp + ", " + nameof(currentModeAppManual) + " - " + currentModeAppManual);
        }

        private string GetValueOfConfigParameter(List<ParameterConfig> listOfParameters, string nameParameter, string defaultValue, bool pass = false)
        {
            return listOfParameters.FindLast(x => x?.parameterName?.Trim() == nameParameter)?.Value?.Trim() != null ?
                   listParameters.FindLast(x => x?.parameterName?.Trim() == nameParameter)?.Value?.Trim() :
                   defaultValue;
        }

        private List<ParameterConfig> ReturnListParametersWithEmptyValue(List<ParameterConfig> listOfParameters)
        {
            List<ParameterConfig> parameterConfigs = new List<ParameterConfig>();
            foreach (var parameter in listOfParameters)
            {
                if (string.IsNullOrWhiteSpace(parameter.Value))
                {
                    parameterConfigs.Add(parameter);
                }
            }
            return parameterConfigs;
        }

        private void RefreshConfigInMainDBItem_Click(object sender, EventArgs e)
        {
            AddExceptedParametersIntoConfigurationDb();
        }

        private void AddExceptedParametersIntoConfigurationDb()    //add not existed example of parameters into ConfigTable in the Main Local DB
        {
            _toolStripStatusLabelSetText(StatusLabel2, "Проверяю список параметров конфигурации локальной БД...");

            ParameterOfConfigurationInSQLiteDB configInDB = new ParameterOfConfigurationInSQLiteDB(dbApplication);
            ParameterOfConfiguration parameterOfConfiguration = null;
            listParameters = configInDB.GetParameters("%%");//.FindAll(x => x.isExample == "no");//update work parameters

            foreach (string sParameter in Names.allParametersOfConfig)
            {
                logger.Trace("looking for: " + sParameter + " in local DB");
                if (!listParameters.Any(x => x?.parameterName == sParameter))
                {
                    parameterOfConfiguration = new ParameterOfConfigurationBuilder().
                        SetParameterName(sParameter).
                        SetParameterValue("").
                        SetParameterDescription("").
                        IsPassword(false).
                        SetIsExample("yes");

                    string resultSaving = configInDB.SaveParameter(parameterOfConfiguration);
                    logger.Info("Попытка добавить новый параметр в конфигурацию: " + resultSaving);
                }
            }
            _toolStripStatusLabelSetText(StatusLabel2, "Обновление параметров конфигурации локальной БД завершено");
        }

        private void AddParameterInConfigItem_Click(object sender, EventArgs e)
        {
            AddParameterInConfig();
        }

        private void AddParameterInConfig()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            _ControlVisible(panelView, false);
            btnPropertiesSave.Text = "Сохранить параметр";
            RemoveClickEvent(btnPropertiesSave);
            btnPropertiesSave.Click += new EventHandler(ButtonPropertiesSave_inConfig);

            listParameters = new List<ParameterConfig>();
            ParameterOfConfigurationInSQLiteDB parameter = new ParameterOfConfigurationInSQLiteDB(dbApplication);

            listParameters = parameter.GetParameters("%%");

            foreach (string sParameter in Names.allParametersOfConfig)
            {
                if (!(listParameters.FindLast(x => x?.parameterName?.Trim() == sParameter)?.Value?.Length > 0))
                {
                    listParameters.Add(new ParameterConfig()
                    {
                        parameterName = sParameter,
                        parameterDescription = "Example",
                        Value = "",
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
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            panelViewResize(numberPeopleInLoading);

            periodCombo = new ListBox
            {
                Location = new Point(20, 120),
                Size = new Size(590, 100),
                Parent = groupBoxProperties,
                DrawMode = DrawMode.OwnerDrawFixed,
                Sorted = true
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
            periodCombo.SelectedIndexChanged += new EventHandler(ListBox_SelectedIndexChanged);
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
            textBoxSettings16.Text = listParameters.FindLast(x => x.parameterName == result)?.Value;
            tooltip = listParameters.FindLast(x => x.parameterName == result)?.parameterDescription;
            toolTip1.SetToolTip(textBoxSettings16, tooltip);
        }


        private void textboxDate_KeyPress(object sender, KeyPressEventArgs e)
        {
            string inputed = null;
            if (e.KeyChar == (char)13)//если нажата Enter
            {
                inputed = ReturnStrongNameDayOfSendingReports((sender as TextBox).Text);
                (sender as TextBox).Text = inputed;
            }
        }


        private void ButtonPropertiesSave_inConfig(object sender, EventArgs e) //SaveProperties()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");
            string textLabel = _ReturnTextOfControl(labelSettings9);
            string description = textLabel?.ToLower() == "example" ? "" : textLabel;

            ParameterOfConfigurationInSQLiteDB parameter = new ParameterOfConfigurationInSQLiteDB(dbApplication);

            ParameterOfConfiguration parameterOfConfiguration = new ParameterOfConfigurationBuilder().
                SetParameterName(labelServer1.Text).
                SetParameterValue(textBoxSettings16.Text).
                SetParameterDescription(description).
                IsPassword(checkBox1.Checked).
                SetIsExample("no");

            string resultSaving = parameter.SaveParameter(parameterOfConfiguration);
            MessageBox.Show(parameterOfConfiguration.ParameterName + " (" + parameterOfConfiguration.ParameterValue + ")\n" + resultSaving);

            DisposeTemporaryControls();
            _ControlVisible(panelView, true);

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
        }



        //void ShowDataTableDbQuery(
        private void ShowDataTableDbQuery(System.IO.FileInfo dbApplication, string myTable, string mySqlQuery, string mySqlWhere) //Query data from the Table of the DB
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            DataTable dt;
            using (dt = new DataTable())
            {
                string query = mySqlQuery + " FROM '" + myTable + "' " + mySqlWhere + "; ";

                if (dbApplication.Exists)
                {
                    using (SqLiteDbReader dbReader = new SqLiteDbReader(sqLiteLocalConnectionString, dbApplication))
                    {
                        dt = dbReader.GetDataTable(query);
                    }
                }

                dgvo.ShowData(dataGridView1, dt);
                iCounterLine = dgvo.RowsCount(dataGridView1);
            }
            logger.Trace("ShowDataTableDbQuery: " + iCounterLine);

            nameOfLastTable = myTable;
            sLastSelectedElement = "dataGridView";
        }

        private void ShowDatatableOnDatagridview(DataTable dt, string nameLastTable, string[] nameHidenColumnsArray1 = null) //Query data from the Table of the DB
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            using (DataTable dataTable = dt.Copy())
            {
                for (int i = 0; i < nameHidenColumnsArray1?.Length; i++)
                {
                    if (nameHidenColumnsArray1[i]?.Length > 0)
                        try { dataTable.Columns[nameHidenColumnsArray1[i]].ColumnMapping = MappingType.Hidden; } catch { }
                }

                dgvo.ShowData(dataGridView1, dataTable);
                _toolStripStatusLabelSetText(StatusLabel2, "Всего записей: " + dgvo.RowsCount(dataGridView1));
            }
            nameOfLastTable = nameLastTable;
            sLastSelectedElement = "dataGridView";
        }


        private async Task ExecuteQueryOnLocalDB(System.IO.FileInfo dbApplication, string query) //Delete All data from the selected Table of the DB (both parameters are string)
        {
            string result = string.Empty;
            if (dbApplication.Exists)
            {
                using (SqLiteDbWriter dbWriter = new SqLiteDbWriter(sqLiteLocalConnectionString, dbApplication))
                {
                    using (SQLiteCommand SqlQuery = new SQLiteCommand(query, dbWriter._sqlConnection))
                    {
                        dbWriter.ExecuteQuery(SqlQuery);
                        result += dbWriter.Status;
                    }
                }
            }

            logger.Trace("ExecuteQueryOnLocalDB: query: " + query + "| result: " + result);
        }

        private async Task DeleteDataTableQueryParameters(System.IO.FileInfo dbApplication, string myTable, string sqlParameter1, string sqlData1,
            string sqlParameter2 = "", string sqlData2 = "", string sqlParameter3 = "", string sqlData3 = "",
            string sqlParameter4 = "", string sqlData4 = "", string sqlParameter5 = "", string sqlData5 = "", string sqlParameter6 = "", string sqlData6 = "") //Delete data from the Table of the DB by NAV (both parameters are string)
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;

            string result = string.Empty;
            string query = "DELETE FROM '" + myTable + "' Where " + sqlParameter1 + "= @" + sqlParameter1 +
                           " AND " + sqlParameter2 + "= @" + sqlParameter2 + " AND " + sqlParameter3 + "= @" + sqlParameter3 +
                           " AND " + sqlParameter4 + "= @" + sqlParameter4 + " AND " + sqlParameter5 + "= @" + sqlParameter5 +
                           " AND " + sqlParameter6 + "= @" + sqlParameter6 + ";";

            if (dbApplication.Exists)
            {
                using (SqLiteDbWriter dbWriter = new SqLiteDbWriter(sqLiteLocalConnectionString, dbApplication))
                {
                    SQLiteCommand sqlCommand = null;
                    if (sqlParameter1.Length > 0 && sqlParameter2.Length > 0 && sqlParameter3.Length > 0 && sqlParameter4.Length > 0
                    && sqlParameter5.Length > 0 && sqlParameter6.Length > 0)
                    {
                        sqlCommand = new SQLiteCommand(query, dbWriter._sqlConnection);

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
                        query = "DELETE FROM '" + myTable + "' Where " + sqlParameter1 + "= @" + sqlParameter1 +
                            " AND " + sqlParameter2 + "= @" + sqlParameter2 + " AND " + sqlParameter3 + "= @" + sqlParameter3 +
                            " AND " + sqlParameter4 + "= @" + sqlParameter4 + " AND " + sqlParameter5 + "= @" + sqlParameter5 + ";";

                        sqlCommand = new SQLiteCommand(query, dbWriter._sqlConnection);
                        sqlCommand.Parameters.Add("@" + sqlParameter1, DbType.String).Value = sqlData1;
                        sqlCommand.Parameters.Add("@" + sqlParameter2, DbType.String).Value = sqlData2;
                        sqlCommand.Parameters.Add("@" + sqlParameter3, DbType.String).Value = sqlData3;
                        sqlCommand.Parameters.Add("@" + sqlParameter4, DbType.String).Value = sqlData4;
                        sqlCommand.Parameters.Add("@" + sqlParameter5, DbType.String).Value = sqlData5;
                    }
                    else if (sqlParameter1.Length > 0 && sqlParameter2.Length > 0 && sqlParameter3.Length > 0 && sqlParameter4.Length > 0)
                    {
                        sqlCommand = new SQLiteCommand("DELETE FROM '" + myTable + "' Where " + sqlParameter1 + "= @" + sqlParameter1 +
                            " AND " + sqlParameter2 + "= @" + sqlParameter2 + " AND " + sqlParameter3 + "= @" + sqlParameter3 +
                            " AND " + sqlParameter4 + "= @" + sqlParameter4 + ";");
                        sqlCommand.Parameters.Add("@" + sqlParameter1, DbType.String).Value = sqlData1;
                        sqlCommand.Parameters.Add("@" + sqlParameter2, DbType.String).Value = sqlData2;
                        sqlCommand.Parameters.Add("@" + sqlParameter3, DbType.String).Value = sqlData3;
                        sqlCommand.Parameters.Add("@" + sqlParameter4, DbType.String).Value = sqlData4;
                    }
                    else if (sqlParameter1.Length > 0 && sqlParameter2.Length > 0 && sqlParameter3.Length > 0)
                    {
                        query = "DELETE FROM '" + myTable + "' Where " + sqlParameter1 + "= @" +
                            sqlParameter1 + " AND " + sqlParameter2 + "= @" + sqlParameter2 + " AND " +
                            sqlParameter3 + "= @" + sqlParameter3 + ";";

                        sqlCommand = new SQLiteCommand(query, dbWriter._sqlConnection);
                        sqlCommand.Parameters.Add("@" + sqlParameter1, DbType.String).Value = sqlData1;
                        sqlCommand.Parameters.Add("@" + sqlParameter2, DbType.String).Value = sqlData2;
                        sqlCommand.Parameters.Add("@" + sqlParameter3, DbType.String).Value = sqlData3;

                    }
                    else if (sqlParameter1.Length > 0 && sqlParameter2.Length > 0)
                    {
                        query = "DELETE FROM '" + myTable + "' Where " + sqlParameter1 + "= @" + sqlParameter1 + " AND " +
                            sqlParameter2 + "= @" + sqlParameter2 + ";";

                        sqlCommand = new SQLiteCommand(query, dbWriter._sqlConnection);
                        sqlCommand.Parameters.Add("@" + sqlParameter1, DbType.String).Value = sqlData1;
                        sqlCommand.Parameters.Add("@" + sqlParameter2, DbType.String).Value = sqlData2;
                    }
                    else if (sqlParameter1.Length > 0)
                    {
                        query = "DELETE FROM '" + myTable + "' Where " + sqlParameter1 + "= @" + sqlParameter1 + ";";

                        sqlCommand = new SQLiteCommand(query, dbWriter._sqlConnection);
                        sqlCommand.Parameters.Add("@" + sqlParameter1, DbType.String).Value = sqlData1;
                    }
                    dbWriter.ExecuteQuery(sqlCommand);
                    result += dbWriter.Status;
                }
            }

            logger.Trace(method + ": query: " + query + "| result: " + result);
        }


        private async Task CheckAliveIntellectServer(string serverName, string userName, string userPasswords) //Check alive the SKD Intellect-server and its DB's 'intellect'
        {
            //stop checking last registrations
            checkInputsOutputs = false;
            logger.Trace("-= CheckAliveIntellectServer =-");

            logger.Trace("Проверка доступности " + serverName + ". Ждите окончания процесса...");
            bServer1Exist = false;

            string query = "SELECT database_id FROM sys.databases WHERE Name ='intellect' ";

            try
            {
                using (SqlDbReader sqlDbTableReader = new SqlDbReader(sqlServerConnectionString))
                {
                    sqlDbTableReader.GetData(query);
                    bServer1Exist = true;
                }
            }
            catch (Exception err)
            {
                _toolStripStatusLabelSetText(StatusLabel2, "Ошибка доступа к " + serverName + " SQL БД СКД-сервера!", true, "server: " + serverName + "|user: " + userName + "|password: " + userPasswords + "\n" + err.ToString());
            }

            logger.Trace("CheckAliveIntellectServer: query: " + query + "| result: " + bServer1Exist);
        }


        private async void GetADUsersItem_Click(object sender, EventArgs e)
        {
            await Task.Run(() => GetUsersFromAD());
        }

        private void GetUsersFromAD()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            logger.Trace("GetUsersFromAD: ");
            string user = null;
            string password = null;
            string domain = null;
            string domainController = null;

            ADData ad;
            usersAD = new List<UserAD>();

            listParameters = new List<ParameterConfig>();
            ParameterOfConfigurationInSQLiteDB parameters = new ParameterOfConfigurationInSQLiteDB(dbApplication);

            listParameters = parameters.GetParameters("%%").FindAll(x => x.isExample == "no"); //load only real data

            user = listParameters.FindLast(x => x?.parameterName == @"UserName")?.Value;
            password = listParameters.FindLast(x => x?.parameterName == @"UserPassword")?.Value;
            domainController = listParameters.FindLast(x => x?.parameterName == @"DomainController")?.Value;
            domain = listParameters.FindLast(x => x?.parameterName == @"DomainOfUser")?.Value;

            logger.Trace("user, domain, password, server: " + user + " |" + domain + " |" + password + " |" + domainController);

            if (user?.Length > 0 && password?.Length > 0 && domain?.Length > 0 && domainController?.Length > 0)
            {
                _toolStripStatusLabelSetText(StatusLabel2, "Получаю данные из домена: " + domain);

                ad = new ADData(user, domain, password, domainController);
                ad.ADUsersCollection.CollectionChanged += Users_CollectionChanged;
                usersAD = ad.GetADUsers().ToList();
                usersAD.Sort();

                logger.Trace("GetUsersFromAD: count users in list: " + usersAD.Count);

                //передать дальше в обработку:
                foreach (var person in usersAD)
                {
                    logger.Trace(person?.fio + " |" + person?.login + " |" + person?.code + " |" + person?.mail + " |" + person?.lastLogon);
                }
                countUsers = usersAD.Count;
                _toolStripStatusLabelSetText(StatusLabel2, "Из домена " + domain + " получено " + countUsers + " ФИО сотрудников");
            }
            else
            {
                _toolStripStatusLabelSetText(
                    StatusLabel2,
                    "Ошибка доступа к домену " + domain,
                    true,
                    "It hasn't access to AD: user: " + user + "| domain: " + domain + "| password: " + password + "| server: " + domainController);
            }
        }

        //уведомление о количестве и последнем полученном из AD пользователей
        private void Users_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            switch (e.Action)
            {
                case NotifyCollectionChangedAction.Add: // если добавление
                    UserAD newUser = e.NewItems[0] as UserAD;
                    stimerPrev = "Получено из AD: " + newUser.id + " пользователей, последний: " + ShortFIO(newUser.fio);
                    break;
                /*   case NotifyCollectionChangedAction.Remove: // если удаление
                       UserAD oldUser = e.OldItems[0] as UserAD;
                       break;
                   case NotifyCollectionChangedAction.Replace: // если замена
                       UserAD replacedUser = e.OldItems[0] as UserAD;
                       UserAD replacingUser = e.NewItems[0] as UserAD;
                       break;*/
                default:
                    break;
            }
        }

        private async void GetFio_Click(object sender, EventArgs e)  //DoListsFioGroupsMailings()
        {
            _ProgressBar1Start();
            currentAction = "GetFIO";
            CheckBoxesFiltersAll_SetState(false);
            CheckBoxesFiltersAll_Enable(false);
            _MenuItemEnabled(LoadDataItem, false);
            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(SettingsMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            _MenuItemEnabled(GetFioItem, false);
            _ControlEnable(dataGridView1, false);

            await Task.Run(() => CheckAliveIntellectServer(sServer1, sServer1UserName, sServer1UserPassword).GetAwaiter().GetResult());

            if (bServer1Exist)
            {
                await Task.Run(() => DoListsFioGroupsMailings().GetAwaiter().GetResult());

                _MenuItemVisible(listFioItem, true);
                _MenuItemEnabled(SettingsMenuItem, true);
                _MenuItemEnabled(GetFioItem, true);
                _MenuItemEnabled(FunctionMenuItem, true);
                _MenuItemEnabled(LoadDataItem, true);
                _MenuItemEnabled(GroupsMenuItem, true);

                _ControlEnable(dataGridView1, true);
                _ControlVisible(dataGridView1, true);
                _ControlVisible(pictureBox1, false);
                _ControlEnable(comboBoxFio, true);

                _toolStripStatusLabelForeColor(StatusLabel1, Color.Black);
                _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);
            }
            else
            {
                _MenuItemEnabled(SettingsMenuItem, true);
            }
            _ProgressBar1Stop();
        }

        private async Task DoListsFioGroupsMailings()  //  GetDataFromRemoteServers()  ImportTablePeopleToTableGroupsInLocalDB()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            countGroups = 0;
            countUsers = 0;
            countMailers = 0;

            if (currentAction != @"sendEmail")
            { _toolStripStatusLabelSetText(StatusLabel2, "Получаю данные с серверов..."); }
            GetUsersFromAD();

            DataTable dtTempIntermediate = dtPeople.Clone();
            GetDataFromRemoteServers(dtTempIntermediate, peopleShifts);

            if (currentAction != @"sendEmail")
            { _toolStripStatusLabelSetText(StatusLabel2, "Формирую и записываю группы в локальную базу..."); }
            WriteGroupsMailingsInLocalDb(dtTempIntermediate, peopleShifts);

            if (currentAction != @"sendEmail")
            { _toolStripStatusLabelSetText(StatusLabel2, "Записываю ФИО в локальную базу..."); }

            WritePeopleInLocalDB(dbApplication.ToString(), dtTempIntermediate);

            if (currentAction != @"sendEmail")
            {
                dtPersonTemp = LeaveAndOrderColumnsOfDataTable(dtTempIntermediate, Names.orderColumnsFIO);

                ShowDatatableOnDatagridview(dtPersonTemp, "ListFIO");

                _toolStripStatusLabelSetText(StatusLabel2, "Записано в локальную базу: " + countUsers + " ФИО, " + countGroups + " групп и " + countMailers + " рассылок");
            }
            dtTempIntermediate?.Dispose();
        }

        //Get the list of registered users
        private void GetDataFromRemoteServers(DataTable dataTablePeople, List<PeopleShift> peopleShifts)
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            EmployeeFull personFromServer = new EmployeeFull();
            DataRow row;

            string query;
            string fio = "";
            string nav = "";
            string groupName = "";
            string idGroup = "";
            string depName = "";
            string depBoss = "";

            string timeStart = ConvertDecimalTimeToStringHHMM(_numUpDownReturn(numUpDownHourStart), _numUpDownReturn(numUpDownMinuteStart));
            string timeEnd = ConvertDecimalTimeToStringHHMM(_numUpDownReturn(numUpDownHourEnd), _numUpDownReturn(numUpDownMinuteEnd));
            string dayStartShift = "";
            string dayStartShift_ = "";

            listFIO = new List<Employee>();
            Dictionary<string, Department> departments = new Dictionary<string, Department>();
            Department departmentFromDictionary;

            _comboBoxClr(comboBoxFio);
            _toolStripStatusLabelSetText(StatusLabel2, "Запрашиваю данные с " + sServer1 + ". Ждите окончания процесса...");

            string confitionToLoad = "";
            using (var sqlConnection = new SQLiteConnection(sqLiteLocalConnectionString))
            {
                sqlConnection.Open();

                _toolStripStatusLabelSetText(StatusLabel2, "Получаю из локальной базы список городов для загрузки списка ФИО из MySQL базы...");
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
                query = "SELECT id,level_id,name,owner_id,parent_id,region_id,schedule_id FROM OBJ_DEPARTMENT";
                using (SqlDbReader sqlDbTableReader = new SqlDbReader(sqlServerConnectionString))
                {
                    System.Data.SqlClient.SqlDataReader sqlData = sqlDbTableReader.GetData(query);

                    //import group from SCA server
                    foreach (DbDataRecord record in sqlData)
                    {
                        idGroup = record["id"]?.ToString()?.Trim();
                        groupName = record?["name"]?.ToString()?.Trim();
                        if (groupName?.Length > 0 && idGroup?.Length > 0 && !departments.ContainsKey(idGroup))
                        {
                            departments.Add(idGroup, new Department()
                            {
                                _departmentId = idGroup,
                                _departmentDescription = groupName,
                                _departmentBossCode = sServer1
                            });
                        }
                        _ProgressWork1Step();
                    }
                }

                //import users from с SCA server
                query = "SELECT id, name, surname, patronymic, post, tabnum, parent_id, facility_code, card FROM OBJ_PERSON WHERE is_locked = '0' AND facility_code NOT LIKE '' AND card NOT LIKE '' ";
                logger.Trace(query);
                using (SqlDbReader sqlDbTableReader = new SqlDbReader(sqlServerConnectionString))
                {
                    System.Data.SqlClient.SqlDataReader sqlData = sqlDbTableReader.GetData(query);
                    foreach (DbDataRecord record in sqlData)
                    {
                        if (record["name"]?.ToString()?.Trim()?.Length > 0)
                        {
                            row = dataTablePeople.NewRow();
                            fio = (record["name"]?.ToString()?.Trim() + " " + record["surname"]?.ToString()?.Trim() + " " + record["patronymic"]?.ToString()?.Trim()).Replace(@"  ", @" ");
                            groupName = record["parent_id"]?.ToString()?.Trim();
                            nav = record["tabnum"]?.ToString()?.Trim()?.ToUpper();

                            departmentFromDictionary = new Department();

                            if (departments.TryGetValue(groupName, out departmentFromDictionary))
                            {
                                depName = departmentFromDictionary._departmentDescription;
                            }
                            else { depName = ""; }

                            row[Names.N_ID] = Convert.ToInt32(record["id"]?.ToString()?.Trim());
                            row[Names.FIO] = fio;
                            row[Names.CODE] = nav;

                            row[Names.GROUP] = groupName;
                            row[Names.DEPARTMENT] = depName;
                            row[Names.DEPARTMENT_ID] = sServer1.IndexOf('.') > -1 ? sServer1.Remove(sServer1.IndexOf('.')) : sServer1;

                            row[Names.EMPLOYEE_POSITION] = record["post"]?.ToString()?.Trim();

                            row[Names.DESIRED_TIME_IN] = timeStart;
                            row[Names.DESIRED_TIME_OUT] = timeEnd;

                            dataTablePeople.Rows.Add(row);

                            listFIO.Add(new Employee { fio = fio, code = nav });

                            _ProgressWork1Step();
                        }
                    }
                }

                // import users, shifts and group from web DB
                int tmpSeconds = 0;
                groupName = mysqlServer;
                _toolStripStatusLabelSetText(StatusLabel2, "Запрашиваю данные с " + mysqlServer + ". Ждите окончания процесса...");


                // import departments from web DB
                logger.Trace(mysqlServerConnectionStringDB1);
                query = "SELECT id, parent_id, name, boss_code FROM dep_struct ORDER by id";
                logger.Trace(query);
                using (MySqlDbReader mysqlDbTableReader = new MySqlDbReader(mysqlServerConnectionStringDB1))
                {
                    MySql.Data.MySqlClient.MySqlDataReader mysqlData = mysqlDbTableReader.GetData(query);
                    while (mysqlData.Read())
                    {
                        idGroup = mysqlData?.GetString(@"id");

                        if (mysqlData?.GetString(@"name")?.Length > 0 && idGroup?.Length > 0 && !departments.ContainsKey(idGroup))
                        {
                            departments.Add(idGroup, new Department()
                            {
                                _departmentId = idGroup,
                                _departmentDescription = mysqlData.GetString(@"name"),
                                _departmentBossCode = mysqlData?.GetString(@"boss_code")
                            });
                        }
                        _ProgressWork1Step();
                    }
                }

                // import individual shifts of people from web DB
                query = "Select code,start_date,mo_start,mo_end,tu_start,tu_end,we_start,we_end,th_start,th_end,fr_start,fr_end, " +
                                "sa_start,sa_end,su_start,su_end,comment FROM work_time ORDER by start_date";
                logger.Trace(query);
                using (MySqlDbReader mysqlDbTableReader = new MySqlDbReader(mysqlServerConnectionStringDB1))
                {
                    MySql.Data.MySqlClient.MySqlDataReader mysqlData = mysqlDbTableReader.GetData(query);
                    while (mysqlData.Read())
                    {
                        if (mysqlData.GetString(@"code")?.Length > 0)
                        {
                            try { dayStartShift = DateTime.Parse(mysqlData.GetMySqlDateTime(@"start_date").ToString()).ToYYYYMMDD(); }
                            catch
                            { dayStartShift = DateTime.Parse("1980-01-01").ToYYYYMMDD(); }

                            peopleShifts.Add(new PeopleShift()
                            {
                                _nav = mysqlData.GetString(@"code").Replace('C', 'S'),
                                _dayStartShift = dayStartShift,
                                _MoStart = Convert.ToInt32(mysqlData.GetString(@"mo_start")),
                                _MoEnd = Convert.ToInt32(mysqlData.GetString(@"mo_end")),
                                _TuStart = Convert.ToInt32(mysqlData.GetString(@"tu_start")),
                                _TuEnd = Convert.ToInt32(mysqlData.GetString(@"tu_end")),
                                _WeStart = Convert.ToInt32(mysqlData.GetString(@"we_start")),
                                _WeEnd = Convert.ToInt32(mysqlData.GetString(@"we_end")),
                                _ThStart = Convert.ToInt32(mysqlData.GetString(@"th_start")),
                                _ThEnd = Convert.ToInt32(mysqlData.GetString(@"th_end")),
                                _FrStart = Convert.ToInt32(mysqlData.GetString(@"fr_start")),
                                _FrEnd = Convert.ToInt32(mysqlData.GetString(@"fr_end")),
                                _SaStart = Convert.ToInt32(mysqlData.GetString(@"sa_start")),
                                _SaEnd = Convert.ToInt32(mysqlData.GetString(@"sa_end")),
                                _SuStart = Convert.ToInt32(mysqlData.GetString(@"su_start")),
                                _SuEnd = Convert.ToInt32(mysqlData.GetString(@"su_end")),
                                _Status = "",
                                _Comment = mysqlData.GetString(@"comment")
                            });
                            _ProgressWork1Step();
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
                        }
                        catch {/*nothing todo on any errors */ }
                    }
                }
                // import people from web DB
                query = "Select code, family_name, first_name, last_name, vacancy, department, boss_id, city FROM personal " + confitionToLoad;//where hidden=0
                logger.Trace(query);
                using (MySqlDbReader mysqlDbTableReader = new MySqlDbReader(mysqlServerConnectionStringDB1))
                {
                    MySql.Data.MySqlClient.MySqlDataReader mysqlData = mysqlDbTableReader.GetData(query);
                    while (mysqlData.Read())
                    {
                        if (mysqlData.GetString(@"family_name")?.Trim()?.Length > 0)
                        {
                            row = dataTablePeople.NewRow();
                            personFromServer = new EmployeeFull();

                            fio = (mysqlData.GetString(@"family_name")?.Trim() + " " + mysqlData.GetString(@"first_name")?.Trim() + " " + mysqlData.GetString(@"last_name")?.Trim())?.Replace(@"  ", @" ");

                            personFromServer.fio = fio.Replace("&acute;", "'");
                            personFromServer.code = mysqlData.GetString(@"code")?.Trim()?.ToUpper()?.Replace('C', 'S');
                            personFromServer.DepartmentId = mysqlData.GetString(@"department")?.Trim();

                            departmentFromDictionary = new Department();

                            if (departments.TryGetValue(personFromServer?.DepartmentId, out departmentFromDictionary))
                            {
                                depName = departmentFromDictionary?._departmentDescription;
                                depBoss = departmentFromDictionary?._departmentBossCode;
                            }

                            personFromServer.Department = depName ?? personFromServer?.DepartmentId;

                            personFromServer.DepartmentBossCode = depBoss?.Length > 0 ? depBoss : mysqlData.GetString(@"boss_id")?.Trim();

                            personFromServer.City = mysqlData.GetString(@"city")?.Trim();

                            personFromServer.PositionInDepartment = mysqlData.GetString(@"vacancy")?.Trim();
                            personFromServer.GroupPerson = groupName;

                            personFromServer.Shift = dayStartShift_;

                            personFromServer.ControlInHHMM = timeStart;
                            personFromServer.ControlOutHHMM = timeEnd;

                            dayStartShift = peopleShifts.FindLast((x) => x._nav == personFromServer.code)._dayStartShift;
                            if (dayStartShift?.Length > 0)
                            {
                                personFromServer.Shift = "Индивидуальный график с " + dayStartShift;

                                tmpSeconds = peopleShifts.FindLast((x) => x._nav == personFromServer.code)._MoStart;
                                personFromServer.ControlInHHMM = ConvertSecondsTimeToStringHHMMArray(tmpSeconds)[2];

                                tmpSeconds = peopleShifts.FindLast((x) => x._nav == personFromServer.code)._MoEnd;
                                personFromServer.ControlOutHHMM = ConvertSecondsTimeToStringHHMMArray(tmpSeconds)[2];

                                personFromServer.Comment = peopleShifts.FindLast((x) => x._nav == personFromServer.code)._Comment;

                                logger.Trace("Индивидуальный график с " + dayStartShift + " " + personFromServer.code + " " + personFromServer.Comment);
                            }
                            row[Names.FIO] = personFromServer.fio;
                            row[Names.CODE] = personFromServer.code;

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

                            listFIO.Add(new Employee { fio = personFromServer.fio, code = personFromServer.code });

                            _ProgressWork1Step();
                        }
                    }
                }

                dataTablePeople.AcceptChanges();
                logger.Trace("departments.count: " + departments.Count);
                _toolStripStatusLabelSetText(StatusLabel2, "ФИО и наименования департаментов получены.");
            }
            catch (Exception err)
            {
                _toolStripStatusLabelSetText(
                    StatusLabel2,
                    "Возникла ошибка во время получения данных с серверов.",
                    true, err.ToString()
                    );
            }

            query = fio = nav = groupName = depName = depBoss = timeStart = timeEnd = dayStartShift = dayStartShift_ = confitionToLoad = null;
            row = null; departments = null; departmentFromDictionary = null; personFromServer = null;
        }

        private void WriteGroupsMailingsInLocalDb(DataTable dataTablePeople, List<PeopleShift> peopleShifts)
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            _toolStripStatusLabelSetText(StatusLabel2, "Формирую обновленные списки ФИО, департаментов и рассылок...");

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
            _ProgressWork1Step();

            string skdName = sServer1.Split('.')[0];
            int iSKD = 0;
            int iMysql = 0;
            foreach (var dr in dataTablePeople.AsEnumerable())
            {
                depId = dr[Names.DEPARTMENT_ID]?.ToString();
                depBossEmail = usersAD.Find((x) => x.code == dr[Names.CHIEF_ID]?.ToString())?.mail;
                if (depId?.Length > 0 && !groups.ContainsKey("@" + depId))
                {
                    if (depId == skdName && iSKD < 1)
                    {
                        groups.Add("@" + depId, new DepartmentFull()
                        {
                            _departmentId = "@" + depId,
                            _departmentDescription = "skd",
                            _departmentBossCode = dr[Names.CHIEF_ID]?.ToString(),
                            _departmentBossEmail = depBossEmail
                        });
                        iSKD++;
                    }
                    else if (depId != skdName)
                    {
                        groups.Add("@" + depId, new DepartmentFull()
                        {
                            _departmentId = "@" + depId,
                            _departmentDescription = dr[Names.DEPARTMENT]?.ToString(),
                            _departmentBossCode = dr[Names.CHIEF_ID]?.ToString(),
                            _departmentBossEmail = depBossEmail
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
                            _departmentId = depId,
                            _departmentDescription = "web",
                            _departmentBossCode = "GetCodeFromDB",
                            _departmentBossEmail = "GetEmailFromDB"
                        });
                        iMysql++;
                    }
                    else if (depId != mysqlServer)
                    {
                        groups.Add(depId, new DepartmentFull()
                        {
                            _departmentId = depId,
                            _departmentDescription = dr[Names.DEPARTMENT]?.ToString(),
                            _departmentBossCode = "GetCodeFromDB",
                            _departmentBossEmail = "GetEmailFromDB"
                        });
                    }
                }
                _ProgressWork1Step();
            }
            foreach (var dep in groups)
            {
                logger.Trace(dep.Value._departmentId + " " + dep.Value._departmentDescription + " " + dep.Value._departmentBossCode + " " + dep.Value._departmentBossEmail);
            }
            logger.Trace("groups.count: " + groups.Distinct().Count());

            foreach (var strDepartment in groups)
            {
                if (strDepartment.Value?._departmentId?.Length > 0)
                {
                    departmentsUniq.Add(new Department
                    {
                        _departmentId = strDepartment.Value._departmentId,
                        _departmentDescription = strDepartment.Value._departmentDescription,
                        _departmentBossCode = strDepartment.Value._departmentBossCode
                    });

                    if (strDepartment.Value?._departmentBossEmail?.Length > 0)
                    {
                        departmentsEmailUniq.Add(new DepartmentFull
                        {
                            _departmentId = strDepartment.Value._departmentId,
                            _departmentDescription = strDepartment.Value._departmentDescription,
                            _departmentBossEmail = strDepartment.Value._departmentBossEmail
                        });
                    }
                }
            }
            _ProgressWork1Step();

            if (dbApplication.Exists)
            {
                logger.Trace("Чищу базу от старых списков с ФИО...");

                ExecuteQueryOnLocalDB(dbApplication, "DELETE FROM 'LastTakenPeopleComboList';").GetAwaiter().GetResult();

                foreach (var department in departmentsUniq?.ToList()?.Distinct())
                {
                    DeleteDataTableQueryParameters(dbApplication, "PeopleGroup", "GroupPerson", department?._departmentId).GetAwaiter().GetResult();
                    _ProgressWork1Step();
                }
                _ProgressWork1Step();

                using (var sqlConnection = new SQLiteConnection(sqLiteLocalConnectionString))
                {
                    sqlConnection.Open();

                    var sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();

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
                                    departmentsEmailUniq.RemoveWhere(x => x._departmentBossEmail == dbRecordTemp);
                                }
                            }
                        }
                    }
                    sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();

                    logger.Info("Записываю новые группы ...");
                    sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                    foreach (var deprtment in departmentsUniq.ToList().Distinct())
                    {
                        depName = deprtment._departmentId;
                        depDescr = deprtment._departmentDescription;

                        depBoss = deprtment._departmentBossCode?.Length > 0 ? deprtment._departmentBossCode : "Default_Recepient_code_From_Db";
                        using (var command = new SQLiteCommand("INSERT OR REPLACE INTO 'PeopleGroupDescription' (GroupPerson, GroupPersonDescription, Recipient) " +
                                               "VALUES (@GroupPerson, @GroupPersonDescription, @Recipient)", sqlConnection))
                        {
                            command.Parameters.Add("@GroupPerson", DbType.String).Value = depName;
                            command.Parameters.Add("@GroupPersonDescription", DbType.String).Value = depDescr;
                            command.Parameters.Add("@Recipient", DbType.String).Value = depBoss;
                            try { command.ExecuteNonQuery(); } catch { }
                        }

                        logger.Trace("CreatedGroup: " + depName + "(" + depDescr + ")");
                        _ProgressWork1Step();
                    }
                    sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();

                    logger.Info("Записываю новые рассылки по группам с учетом исключений...");
                    string recipientEmail = "";
                    sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                    foreach (var deprtment in departmentsEmailUniq?.ToList()?.Distinct())
                    {
                        depName = deprtment?._departmentId;
                        depDescr = deprtment?._departmentDescription;
                        depBoss = deprtment?._departmentBossCode;
                        recipientEmail = deprtment?._departmentBossEmail;

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

                            logger.Trace("SaveMailing: " + recipientEmail + " " + depName + " " + depDescr);
                        }
                        _ProgressWork1Step();
                    }
                    sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();

                    logger.Info("Записываю новые индивидуальные графики...");
                    sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();
                    foreach (var shift in peopleShifts?.ToArray())
                    {
                        if (shift._nav?.Length > 0)
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
                                try { sqlCommand.ExecuteNonQuery(); } catch (Exception err) { MessageBox.Show(err.ToString()); }
                                _ProgressWork1Step();
                            }
                        }
                    }
                    sqlCommand1 = new SQLiteCommand("end", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();

                    int.TryParse(departmentsUniq?.ToArray()?.Distinct()?.Count().ToString(), out countGroups);
                    int.TryParse(departmentsEmailUniq?.ToArray()?.Distinct()?.Count().ToString(), out countMailers);

                    logger.Info("Записано групп: " + countGroups);
                    logger.Info("Записано рассылок: " + countMailers);
                }
            }

            departmentsUniq = null;
            departmentsEmailUniq = null;

            _ProgressWork1Step();
            _toolStripStatusLabelSetText(StatusLabel2, "Списки ФИО и департаментов получены.");
        }

        private void listFioItem_Click(object sender, EventArgs e) //ListFioReturn()
        {
            nameOfLastTable = "ListFIO";

            _ControlEnable(comboBoxFio, true);
            SeekAndShowMembersOfGroup("");
        }

        private async void Export_Click(object sender, EventArgs e)
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            _ProgressBar1Start();
            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(SettingsMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            _ControlEnable(dataGridView1, false);

            filePathExcelReport = System.IO.Path.Combine(localAppFolderPath, "InputOutputs " + DateTime.Now.ToString("yyyyMMdd_HHmmss"));

            await Task.Run(() => ExportDatatableSelectedColumnsToExcel(dtPersonTemp, "InputOutputsOfStaff", filePathExcelReport).GetAwaiter().GetResult());
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("explorer.exe", " /select, " + filePathExcelReport + @".xlsx")); // //System.Reflection.Assembly.GetExecutingAssembly().Location)

            _MenuItemEnabled(FunctionMenuItem, true);
            _MenuItemEnabled(SettingsMenuItem, true);
            _MenuItemEnabled(GroupsMenuItem, true);
            _ControlEnable(dataGridView1, true);
            _ProgressBar1Stop();
        }

        private string TryToMakeNameFile(string fileName)
        {
            string pathToFile = fileName;
            if (System.IO.File.Exists(fileName + @".xlsx"))
            {
                pathToFile = fileName + "_1";
                TryToMakeNameFile(pathToFile);
            }

            return pathToFile;
        }

        private async Task ExportDatatableSelectedColumnsToExcel(DataTable dataTable, string nameReport, string filePath)  //Export DataTable to Excel 
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            string pathToFile = TryToMakeNameFile(filePath);

            reportExcelReady = false;


            dataTable.SetColumnsOrder(Names.orderColumnsFinacialReport);
            DataTable dtExport;
            string sort = Names.DEPARTMENT + ", " + Names.FIO + ", " + Names.DATE_REGISTRATION + " ASC";
            // Order of collumns
            DataView dv = dataTable.DefaultView;
            try
            {
                dv.Sort = sort;
                dtExport = dv.ToTable();
            }
            catch
            {
                sort = Names.GROUP + ", " + Names.FIO + ", " + Names.DATE_REGISTRATION + " ASC";
                dv.Sort = sort;
                dtExport = dv.ToTable();
            }

            logger.Trace("Сортировка: " + sort);
            logger.Trace("В таблице " + dataTable.TableName + " столбцов всего - " + dtExport.Columns.Count + ", строк - " + dtExport.Rows.Count);
            _toolStripStatusLabelSetText(StatusLabel2, "Генерирую Excel-файл по отчету: '" + nameReport + "'");
            _ProgressWork1Step();

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
                _ProgressWork1Step();

                int rows = 1;
                int rowsInTable = dtExport.Rows.Count;
                int columnsInTable = indexColumns.Length - 1; // p.Length;
                sheet.Name = nameReport;
                //sheet.Names.Add("next", "=" + Path.GetFileNameWithoutExtension(filePathExcelReport) + "!$A$1", true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _ProgressWork1Step();

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
                catch (Exception err) { logger.Warn("Нарушения: " + err.ToString()); }
                _ProgressWork1Step();

                try
                {
                    Microsoft.Office.Interop.Excel.Range rangeColumnC = sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(@"Отпуск")) + 1)];
                    rangeColumnC.Cells.EntireColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                }
                catch (Exception err) { logger.Warn("Отпуск: " + err.ToString()); }

                try
                {
                    Microsoft.Office.Interop.Excel.Range rangeColumnD = sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(Names.EMPLOYEE_HOOKY)) + 1)];
                    rangeColumnD.Cells.EntireColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                }
                catch (Exception err) { logger.Warn("Отгул: " + err.ToString()); }

                try
                {
                    Microsoft.Office.Interop.Excel.Range rangeColumnE = sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(Names.EMPLOYEE_SICK_LEAVE)) + 1)];
                    rangeColumnE.Cells.EntireColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                }
                catch (Exception err) { logger.Warn("Больничный: " + err.ToString()); }

                try
                {
                    Microsoft.Office.Interop.Excel.Range rangeColumnF = sheet.Columns[GetExcelColumnName(Array.IndexOf(indexColumns, dtExport.Columns.IndexOf(Names.EMPLOYEE_ABSENCE)) + 1)];
                    rangeColumnF.Cells.EntireColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                }
                catch (Exception err) { logger.Warn("Отсутствовал: " + err.ToString()); }
                _ProgressWork1Step();

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
                _ProgressWork1Step();

                foreach (DataRow row in dtExport.Rows)
                {
                    rows++;
                    for (int column = 0; column < columnsInTable; column++)
                    {
                        sheet.Cells[rows, column + 1].Value = row[indexColumns[column]];
                    }
                    _ProgressWork1Step();
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
                _ProgressWork1Step();

                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                _ProgressWork1Step();

                //Autofilter
                range.Select();
                range.AutoFilter(1, Type.Missing, Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);

                //save document
                workbook.SaveAs(pathToFile + @".xlsx",
                    Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                    false, false,
                    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, true,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value);
                _ProgressWork1Step();

                _toolStripStatusLabelSetText(StatusLabel2, "Отчет сохранен в файл: " + pathToFile + @".xlsx");

                filePath = pathToFile;
                _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);
                reportExcelReady = true;
                releaseObject(range);
                releaseObject(rangeColumnName);
            }
            catch (Exception err)
            {
                _toolStripStatusLabelSetText(
                    StatusLabel2,
                    "Ошибка генерации файла. Проверьте наличие установленного Excel",
                    true, "| ExportDatatableSelectedColumnsToExcel: " + err.ToString());
            }
            finally
            {
                //close document
                workbook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                workbooks.Close();

                //clear temporary objects
                releaseObject(sheet);
                releaseObject(workbook);
                releaseObject(workbooks);
                excel.Quit();
                releaseObject(excel);

                dv?.Dispose();
                dtExport?.Dispose();
            }
            _ProgressWork1Step();

            sLastSelectedElement = "ExportExcel";
        }

        private void releaseObject(object obj) //for function - ExportToExcel()
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                logger.Trace("Exception releasing object of Excel: " + ex.Message);
            }
            finally
            { GC.Collect(); }
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
        { SelectFioAndNavFromCombobox(); }

        private void SelectFioAndNavFromCombobox()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            textBoxFIO.Text = "";
            textBoxNav.Text = "";
            CheckBoxesFiltersAll_Enable(false);
            string[] sComboboxFIO = comboBoxFio?.SelectedItem?.ToString()?.Trim()?.Split('|');
            //    try
            //    {
            if (sComboboxFIO?.Length > 0)
            {
                textBoxFIO.Text = sComboboxFIO[0]?.Trim();
                StatusLabel2.Text = @"Выбран: " + ShortFIO(textBoxFIO?.Text);

                if (sComboboxFIO?.Length > 1)
                {
                    textBoxNav.Text = sComboboxFIO[1]?.Trim();
                }
            }

            //    }
            //    catch { }


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
                _selectedEmployeeVisitor = new EmployeeVisitor { fio = textBoxFIO.Text, code = textBoxNav.Text };
                logger.Trace(_selectedEmployeeVisitor.fio + " " + _selectedEmployeeVisitor.code);
                ShowVisitorsOnDataGridView(visitors, _selectedEmployeeVisitor);
            }
        }


        private void CreateGroupItem_Click(object sender, EventArgs e)
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            if (textBoxGroup.Text.Trim().Length > 0)
            {
                CreateGroupInDB(dbApplication, textBoxGroup.Text.Trim(), textBoxGroupDescription.Text.Trim());
            }

            PersonOrGroupItem.Text = Names.WORK_WITH_A_PERSON;
            nameOfLastTable = "PeopleGroup";
            ListGroups();
        }

        private void CreateGroupInDB(System.IO.FileInfo fileInfo, string nameGroup, string descriptionGroup)
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            if (nameGroup.Length > 0)
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
                _toolStripStatusLabelSetText(StatusLabel2, "Группа - \"" + nameGroup + "\" создана");
            }
        }

        private void ListGroupsItem_Click(object sender, EventArgs e)
        {
            ListGroups();
        }

        private void ListGroups()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            _ControlVisible(groupBoxProperties, false);
            _ControlVisible(dataGridView1, false);

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

            _ControlEnable(comboBoxFio, false);

            _ControlVisible(dataGridView1, true);
            dataGridView1.Select();
        }

        private void UpdateAmountAndRecepientOfPeopleGroupDescription()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            logger.Trace("UpdateAmountAndRecepientOfPeopleGroupDescription");
            List<string> groupsUncount = new List<string>();
            List<ParametersOfGroup> amounts = new List<ParametersOfGroup>();
            HashSet<string> groupsUniq = new HashSet<string>();
            List<DepartmentFull> emails = new List<DepartmentFull>();
            string tmpRec = "";
            string query = "";

            if (dbApplication.Exists)
            {
                using (var sqlConnection = new SQLiteConnection(sqLiteLocalConnectionString))
                {
                    sqlConnection.Open();

                    SQLiteCommand sqlCommand1 = new SQLiteCommand("begin", sqlConnection);
                    sqlCommand1.ExecuteNonQuery();

                    //set to empty for amounts and recepients in the PeopleGroupDescription
                    query = "UPDATE 'PeopleGroupDescription' SET AmountStaffInDepartment='0';";
                    using (var command = new SQLiteCommand(query, sqlConnection))
                    { command.ExecuteNonQuery(); }
                    query = "UPDATE 'PeopleGroupDescription' SET Recipient='';";
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

                    query = "SELECT GroupPerson FROM PeopleGroupDescription;";
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

                    logger.Trace("groupsUncount: " + (new HashSet<string>(groupsUncount)).Count());
                    foreach (string group in new HashSet<string>(groupsUncount))
                    {
                        amounts.Add(new ParametersOfGroup()
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
                            query = "UPDATE 'PeopleGroupDescription' SET AmountStaffInDepartment='" + group._amountMembers + "' WHERE GroupPerson like '" + group._groupName + "';";
                            logger.Trace(query);
                            using (var command = new SQLiteCommand(query, sqlConnection))
                            { command.ExecuteNonQuery(); }

                            query = "UPDATE 'PeopleGroupDescription' SET Recipient='" + group._emails + "' WHERE GroupPerson like '" + group._groupName + "';";
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
            if (nameOfLastTable == "PeopleGroup" || nameOfLastTable == "PeopleGroupDescription")
            {
                dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                    Names.GROUP
                });
                SeekAndShowMembersOfGroup(dgvo.cellValue[0]);
            }
            else if (nameOfLastTable == "Mailing")
            {
                dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                    @"Отчет по группам"
                });
                SeekAndShowMembersOfGroup(dgvo.cellValue[0]);
            }
        }

        private void SeekAndShowMembersOfGroup(string nameGroup)
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

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

            using (SqLiteDbReader dbReader = new SqLiteDbReader(sqLiteLocalConnectionString, dbApplication))
            {
                System.Data.SQLite.SQLiteDataReader data = null;

                try
                {
                    data = dbReader?.GetData(query);
                }
                catch { logger.Info("SeekAndShowMembersOfGroup: no any fio in 'selected'"); }

                if (data != null)
                {
                    foreach (DbDataRecord record in data)
                    {
                        if (record["FIO"]?.ToString()?.Length > 0 && record["NAV"]?.ToString()?.Length > 0)
                        {
                            try { dprtmnt = record[@"Department"].ToString(); } catch { dprtmnt = record[@"GroupPerson"]?.ToString(); }

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
                            numberPeopleInLoading++;
                        }
                    }
                }
                data = null;
            }

            if (numberPeopleInLoading > 0)
            {
                dtPersonTemp = LeaveAndOrderColumnsOfDataTable(dtTemp, Names.orderColumnsFIO);
                ShowDatatableOnDatagridview(dtPersonTemp, "PeopleGroup");
                _MenuItemVisible(DeletePersonFromGroupItem, true);
            }

            query = null;
            dataRow = null;
            dtTemp.Dispose();
        }

        private void AddPersonToGroupItem_Click(object sender, EventArgs e) //AddPersonToGroup() //Add the selected person into the named group
        { AddPersonToGroup(); }

        private void AddPersonToGroup() //Add the selected person into the named group
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            string group = _ReturnTextOfControl(textBoxGroup);
            string groupDescription = _ReturnTextOfControl(textBoxGroupDescription);
            logger.Trace("AddPersonToGroup: group " + group);
            if (dgvo.RowsCount(dataGridView1) > -1)
            {
                DataGridViewOperations dgvo = new DataGridViewOperations();
                dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
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

                    if (group?.Length > 0 && dgvo.cellValue[1]?.Length > 0)
                    {
                        using (var sqlCommand = new SQLiteCommand("INSERT OR REPLACE INTO 'PeopleGroup' (FIO, NAV, GroupPerson, ControllingHHMM, ControllingOUTHHMM, Department, PositionInDepartment, Comment, Shift, DepartmentId, City, Boss) " +
                                                    "VALUES (@FIO, @NAV, @GroupPerson, @ControllingHHMM, @ControllingOUTHHMM, @Department, @PositionInDepartment, @Comment, @Shift, @DepartmentId, @City, @Boss)", connection))
                        {
                            sqlCommand.Parameters.Add("@FIO", DbType.String).Value = dgvo.cellValue[0];
                            sqlCommand.Parameters.Add("@NAV", DbType.String).Value = dgvo.cellValue[1];
                            sqlCommand.Parameters.Add("@GroupPerson", DbType.String).Value = group;
                            sqlCommand.Parameters.Add("@ControllingHHMM", DbType.String).Value = dgvo.cellValue[4];
                            sqlCommand.Parameters.Add("@ControllingOUTHHMM", DbType.String).Value = dgvo.cellValue[5];
                            sqlCommand.Parameters.Add("@Department", DbType.String).Value = dgvo.cellValue[2];
                            sqlCommand.Parameters.Add("@PositionInDepartment", DbType.String).Value = dgvo.cellValue[3];
                            sqlCommand.Parameters.Add("@Comment", DbType.String).Value = "";
                            sqlCommand.Parameters.Add("@DepartmentId", DbType.String).Value = dgvo.cellValue[8];
                            sqlCommand.Parameters.Add("@City", DbType.String).Value = dgvo.cellValue[9];
                            sqlCommand.Parameters.Add("@Boss", DbType.String).Value = dgvo.cellValue[6];
                            sqlCommand.Parameters.Add("@Shift", DbType.String).Value = dgvo.cellValue[7];
                            try { sqlCommand.ExecuteNonQuery(); } catch (Exception ept) { logger.Warn("PeopleGroup: " + ept.ToString()); }
                        }
                        _toolStripStatusLabelSetText(StatusLabel2, "'" + ShortFIO(dgvo.cellValue[0]) + "'" + " добавлен в группу '" + group + "'");
                        _toolStripStatusLabelBackColor(StatusLabel2, SystemColors.Control);
                    }
                    else if (group?.Length > 0 && dgvo?.cellValue[1]?.Length == 0)
                    {
                        _toolStripStatusLabelSetText(StatusLabel2, "Отсутствует NAV-код у: " + ShortFIO(textBoxFIO.Text));
                        _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                    }
                    else if (group?.Length == 0 && dgvo?.cellValue[1]?.Length > 0)
                    {
                        _toolStripStatusLabelSetText(StatusLabel2, "Не указана группа, в которую нужно добавить!");
                        _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                    }
                }
            }

            SeekAndShowMembersOfGroup(group);

            labelGroup.BackColor = SystemColors.Control;
            _MenuItemSetText(PersonOrGroupItem, Names.WORK_WITH_A_PERSON);
            nameOfLastTable = "PeopleGroup";
        }


        private void importPeopleInLocalDBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

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
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

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
                                _departmentId = cell[2],
                                _departmentDescription = cell[3],
                            });

                            checkHourS = cell[5];
                            if (TryParseStringToDecimal(checkHourS) == 0)
                            { checkHourS = numUpHourStart.ToString(); }
                            row[Names.DESIRED_TIME_IN] = ConvertStringsTimeToStringHHMM(checkHourS, cell[6]);

                            checkHourE = cell[7];
                            if (TryParseStringToDecimal(checkHourE) == 0)
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
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");
            logger.Trace("WritePeopleInLocalDB: table - " + dtSource + ", row - " + dtSource.Rows.Count);

            string result = string.Empty;
            string query = null;
            if (dbApplication.Exists)
            {
                query = "INSERT OR REPLACE INTO 'PeopleGroup' (FIO, NAV, GroupPerson, ControllingHHMM, ControllingOUTHHMM, Shift, Comment, Department, PositionInDepartment, DepartmentId, City, Boss) " +
                        "VALUES (@FIO, @NAV, @GroupPerson, @ControllingHHMM, @ControllingOUTHHMM, @Shift, @Comment, @Department, @PositionInDepartment, @DepartmentId, @City, @Boss)";

                using (SqLiteDbWriter dbWriter = new SqLiteDbWriter(sqLiteLocalConnectionString, dbApplication))
                {
                    result = string.Empty;
                    dbWriter.ExecuteQueryBegin();
                    foreach (var dr in dtSource.AsEnumerable())
                    {
                        if (dr[Names.FIO]?.ToString()?.Length > 0 && dr[Names.CODE]?.ToString()?.Length > 0)
                        {
                            using (SQLiteCommand SqlQuery = new SQLiteCommand(query, dbWriter._sqlConnection))
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

                                dbWriter.ExecuteQueryForBulkStepByStep(SqlQuery);
                                result += dbWriter.Status;
                                _ProgressWork1Step();
                            }
                        }
                    }
                    logger.Trace(method + ": query: " + query + "\nresult: " + result);//method = System.Reflection.MethodBase.GetCurrentMethod().Name;
                    dbWriter.ExecuteQueryEnd();

                    result = string.Empty;
                    dbWriter.ExecuteQueryBegin();
                    query = "INSERT OR REPLACE INTO 'LastTakenPeopleComboList' (ComboList) VALUES (@ComboList)";
                    foreach (var str in listFIO)
                    {
                        if (str.fio?.Length > 0 && str.code?.Length > 0)
                        {
                            using (SQLiteCommand SqlQuery = new SQLiteCommand(query, dbWriter._sqlConnection))
                            {
                                SqlQuery.Parameters.Add("@ComboList", DbType.String).Value = str.fio + "|" + str.code;

                                dbWriter.ExecuteQueryForBulkStepByStep(SqlQuery);
                                result += dbWriter.Status;
                                _ProgressWork1Step();
                            }
                        }
                    }
                    logger.Trace(method + ": query: " + query + "\nresult:" + result);//method = System.Reflection.MethodBase.GetCurrentMethod().Name;
                    dbWriter.ExecuteQueryEnd();
                }
            }

            foreach (var str in listFIO)
            { _comboBoxAdd(comboBoxFio, str.fio + "|" + str.code); }
            _ProgressWork1Step();
            if (_comboBoxCountItems(comboBoxFio) > 0)
            { _comboBoxSelectIndex(comboBoxFio, 0); }
            _ProgressWork1Step();

            int.TryParse(listFIO.Count.ToString(), out countUsers);
            logger.Info("Записано ФИО: " + countUsers);
        }

        private void ImportListGroupsDescriptionInLocalDB(string pathToPersonDB, HashSet<Department> departmentsUniq) //use listGroups
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");


            string result = string.Empty;
            string query = null;
            if (dbApplication.Exists)
            {
                query = "INSERT OR REPLACE INTO 'PeopleGroupDescription' (GroupPerson, GroupPersonDescription, Recipient) " +
                                        "VALUES (@GroupPerson, @GroupPersonDescription, @Recipient)";

                using (SqLiteDbWriter dbWriter = new SqLiteDbWriter(sqLiteLocalConnectionString, dbApplication))
                {
                    dbWriter.ExecuteQueryBegin();
                    foreach (var group in departmentsUniq)
                    {
                        if (group?._departmentId?.Length > 0)
                        {
                            using (SQLiteCommand SqlQuery = new SQLiteCommand(query, dbWriter._sqlConnection))
                            {
                                SqlQuery.Parameters.Add("@GroupPerson", DbType.String).Value = group._departmentId;
                                SqlQuery.Parameters.Add("@GroupPersonDescription", DbType.String).Value = group._departmentDescription;
                                SqlQuery.Parameters.Add("@Recipient", DbType.String).Value = group._departmentBossCode;

                                dbWriter.ExecuteQueryForBulkStepByStep(SqlQuery);
                                result += dbWriter.Status;
                            }
                        }
                    }
                    dbWriter.ExecuteQueryEnd();
                    logger.Info(method + ": query: " + query + "\n" + result);//method = System.Reflection.MethodBase.GetCurrentMethod().Name;
                }
            }
        }


        private void GetNamesOfPassagePoints() //Get names of the pass by points
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            logger.Trace("GetNamePoints");
            collectionOfPassagePoints = new CollectionOfPassagePoints();

            if (dbApplication.Exists)
            {
                string query = "Select id, name FROM OBJ_ABC_ARC_READER;";
                logger.Trace("query: " + query);
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
            }
            foreach (var tmp in collectionOfPassagePoints.GetCollection())
            {
                logger.Trace(tmp.Key + " " + tmp.Value._idPoint + " " + tmp.Value._namePoint + " " + tmp.Value._direction + " " + tmp.Value._connectedToServer);
            }
        }


        private void ResetFilterLoadLastIputsOutput_Click(object sender, EventArgs e)
        {
            _selectedEmployeeVisitor = null;
            ShowVisitorsOnDataGridView(visitors);
        }


        private async void LoadLastIputsOutputs_Click(object sender, EventArgs e) //LoadInputsOutputsOfVisitors()
        {
            //how many times continiously to check registrations at the server
            int timesCheckingRegistration = 10;

            //clear painting;
            _paintedEmployeeVisitor = null;
            _selectedEmployeeVisitor = null;

            //status of repeatedly loading of registrations cards from server
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

            string dayStart = _dateTimePickerReturn(dateTimePickerStart).ToYYYYMMDD();
            string dayEnd = _dateTimePickerReturn(dateTimePickerEnd).ToYYYYMMDD();

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

            string dayStart = _dateTimePickerReturn(dateTimePickerStart).ToYYYYMMDD();
            string dayEnd = _dateTimePickerReturn(dateTimePickerEnd).ToYYYYMMDD();

            if (DateTime.Now.ToYYYYMMDD() != dayStart)
            { timesCheckingRegistration = 1; }

            await Task.Run(() => LoadInputsOutputsOfVisitors(dayStart, dayEnd, timesCheckingRegistration));
        }

        private static Visitors visitors = new Visitors();

        //lock to loading of registrations cards from server
        private static object lockerToLoadInsOuts = new object();

        //lock to show data on datagridview
        private static object lockerToShowData = new object();

        //status of repeatedly loading of registrations cards from server
        private static bool checkInputsOutputs = true;

        //Timer for waiting next loading
        IStartStopTimer startStopTimer;

        private void LoadInputsOutputsOfVisitors(string startDay, string endDay, int timesChecking)
        {
            _ControlEnable(comboBoxFio, true);

            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");
            CheckAliveIntellectServer(sServer1, sServer1UserName, sServer1UserPassword).GetAwaiter().GetResult();
            nameOfLastTable = "LastIputsOutputs";
            List<Visitor> visitorsTillNow;

            //Get names of the points
            GetNamesOfPassagePoints();

            visitors = new Visitors();

            //subscribe on action of changed data in collection of visitors
            visitors.collection.CollectionChanged += VisitorsCollectionChanged;

            checkInputsOutputs = true;
            bool firstStage = true;

            startStopTimer = new StartStopTimer(15);

            string startTime = "00:00:00";
            string endTime = "23:59:59";
            do
            {
                lock (lockerToLoadInsOuts)
                {
                    visitorsTillNow = new List<Visitor>();

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

                timesChecking--;
                if (timesChecking <= 0)
                { checkInputsOutputs = false; }

                _toolStripStatusLabelSetText(StatusLabel2, "Загружены данные о регистрации пропусков до: " + startDay + " " + startTime);

                startStopTimer.WaitTime();

            } while (checkInputsOutputs);

            _toolStripStatusLabelSetText(StatusLabel2, "Сбор данных регистрации пропусков завершен");
        }

        private List<Visitor> GetInputsOutputs(ref string startDay, ref string startTime, ref string endDay, ref string endTime)
        {
            _ProgressBar1Start();

            List<Visitor> visitors = new List<Visitor>();
            bool startTimeNotSet = true;

            SideOfPassagePoint sideOfPassagePoint;
            string time, date, fullPointName, fio, action, action_descr, fac, card;
            int idCard = 0; string idCardDescr;

            string query = "SELECT p.param0 as param0, p.param1 as param1, p.objid as objid, p.objtype, p.action as action, " +
            " pe.tabnum as nav, pe.facility_code as fac, pe.card as card, " +
            " CONVERT(varchar, p.date, 120) AS date, CONVERT(varchar, p.time, 114) AS time" +
            " FROM protocol p " +
            " LEFT JOIN OBJ_PERSON pe ON  p.param1=pe.id " +
            " where p.objtype like 'ABC_ARC_READER' " +
            " AND p.param0 like '%%' " +
            // " AND p.param1 like '" + person.idCard + "' " +
            " AND date > '" + startDay + " " + startTime + "' AND date <= '" + endDay + " " + endTime + "' " +
            " ORDER BY date DESC, time DESC;"; //sorting 

            logger.Trace("query: " + query);

            using (SqlDbReader sqlDbTableReader = new SqlDbReader(sqlServerConnectionString))
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
                        time = ConvertStringsTimeToStringHHMMSS(record["time"]?.ToString()?.Trim());

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
            }
            stimerPrev = "";
            _ProgressBar1Stop();
            return visitors;
        }

        private void VisitorsCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (nameOfLastTable != "LastIputsOutputs")//Visitor selected=null
            {
                checkInputsOutputs = false;
            }

            lock (lockerToShowData)
            {
                ShowVisitorsOnDataGridView(visitors, _selectedEmployeeVisitor);
            }
        }

        private void ShowVisitorsOnDataGridView(Visitors visitors_, Visitor selected = null)
        {
            using (DataTable dtTemp = dtPeople.Clone())
            {
                SendListLastRegistrationsToDataTable(visitors_.collection, dtTemp, selected);
                using (DataTable dt = LeaveAndOrderColumnsOfDataTable(dtTemp, Names.orderColumnsLastRegistrations))
                {
                    dgvo.ShowData(dataGridView1, dt);
                    nameOfLastTable = "LastIputsOutputs";

                    //  ShowDatatableOnDatagridview(dt, "LastIputsOutputs");
                    if (_paintedEmployeeVisitor != null)
                    {
                        PaintDataGridViewBy(dataGridView1, Names.FIO, _paintedEmployeeVisitor.fio);
                    }
                }
            }
        }

        private void SendListLastRegistrationsToDataTable(ObservableRangeCollection<Visitor> _visitors, DataTable dt, Visitor selected = null)
        {
            DataRow row = dt.NewRow();
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
            PaintDataGridViewBy(dataGridView1, Names.CHECKPOINT_ACTION, _paintedEmployeeVisitor.action);
        }

        private void PaintRowsFioItem_Click(object sender, EventArgs e)
        {
            _paintedEmployeeVisitor = LookForSelectedVisitor(dataGridView1);
            PaintDataGridViewBy(dataGridView1, Names.FIO, _paintedEmployeeVisitor.fio);
        }

        private EmployeeVisitor LookForSelectedVisitor(DataGridView dgv)
        {
            EmployeeVisitor visitor = new EmployeeVisitor();
            dgvo = new DataGridViewOperations();
            dgvo.FindValuesInCurrentRow(dgv, new string[] {
                Names.FIO,
                Names.N_ID_STRING,
                Names.CHECKPOINT_ACTION
                    });
            visitor.fio = dgvo.cellValue[0];
            visitor.idCard = dgvo.cellValue[1];
            visitor.action = dgvo.cellValue[2];

            return visitor;
        }

        private void PaintDataGridViewBy(DataGridView dgv, string columnName, string desiredData)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row?.Cells[columnName]?.Value?.ToString() == desiredData)
                    row.DefaultCellStyle.BackColor = Color.Red;
                else
                    row.DefaultCellStyle.BackColor = Color.White;
            }
        }



        private void GetDataOfGroup_Click(object sender, EventArgs e) //LoadIdCardRegistrations()
        {
            string group = _ReturnTextOfControl(textBoxGroup);

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
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            _ProgressBar1Start();

            await Task.Run(() => CheckAliveIntellectServer(sServer1, sServer1UserName, sServer1UserPassword).GetAwaiter().GetResult());

            _ControlChangeBackColor(groupBoxPeriod, SystemColors.Control);
            _ControlChangeBackColor(groupBoxTimeStart, SystemColors.Control);
            _ControlChangeBackColor(groupBoxTimeEnd, SystemColors.Control);
            _MenuItemBackColorChange(LoadDataItem, SystemColors.Control);

            _MenuItemEnabled(LoadDataItem, false);
            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(SettingsMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            CheckBoxesFiltersAll_Enable(false);

            _ControlVisible(dataGridView1, false);

            if (bServer1Exist)
            {
                reportStartDay = _dateTimePickerStartReturn().Split(' ')[0];
                reportLastDay = _dateTimePickerEndReturn().Split(' ')[0];

                await Task.Run(() => GetData(_group, reportStartDay, reportLastDay));

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

                _ControlVisible(dataGridView1, true);

                _ControlEnable(checkBoxReEnter, true);
                _ControlEnable(checkBoxTimeViolations, false);
                _ControlEnable(checkBoxWeekend, false);
                _ControlEnable(checkBoxCelebrate, false);
                CheckBoxesFiltersAll_SetState(false);

                panelViewResize(numberPeopleInLoading);
                _ControlChangeBackColor(groupBoxFilterReport, Color.PaleGreen);
            }
            else
            {
                GetInfoSetup();
                _MenuItemEnabled(SettingsMenuItem, true);
            }
            stimerPrev = "";
            _ProgressBar1Stop();

            if (dtPersonTemp?.Rows.Count > 0)
            {
                _MenuItemVisible(TableExportToExcelItem, true);
                _toolStripStatusLabelSetText(StatusLabel2, "Данные регистрации пропусков загружены");
            }
        }

        private void GetData(string _group, string _reportStartDay, string _reportLastDay)
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

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
            LoadRecords(_group, _dateTimePickerStartReturn(), _dateTimePickerEndReturn(), "");

            dtPersonTemp = LeaveAndOrderColumnsOfDataTable(dtPersonRegistrationsFullList.Copy(), Names.orderColumnsRegistrations);

            //show selected data  within the selected collumns   
            ShowDatatableOnDatagridview(dtPersonTemp, "PeopleGroup");

        }

        private void LoadRecords(string nameGroup, string startDate, string endDate, string doPostAction)
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            EmployeeFull person = new EmployeeFull();
            outPerson = new List<OutPerson>();
            outResons = new List<OutReasons>
            {
                new OutReasons()
                { _id = "0", _hourly = 1, _name = @"Unknow", _visibleName = @"Неизвестная" }
            };

            string query = "Select id, name,hourly, visibled_name FROM out_reasons";
            logger.Trace(query);
            using (MySqlDbReader mysqlDbTableReader = new MySqlDbReader(mysqlServerConnectionStringDB1))
            {
                MySql.Data.MySqlClient.MySqlDataReader mysqlData = mysqlDbTableReader.GetData(query);

                while (mysqlData.Read())
                {
                    if (mysqlData?.GetString(@"id")?.Length > 0 && mysqlData?.GetString(@"name")?.Length > 0)
                    {
                        outResons.Add(new OutReasons()
                        {
                            _id = mysqlData.GetString(@"id"),
                            _hourly = Convert.ToInt32(mysqlData.GetString(@"hourly")),
                            _name = mysqlData.GetString(@"name"),
                            _visibleName = mysqlData.GetString(@"visibled_name")
                        });
                    }
                }
            }
            _ProgressWork1Step();


            string date = "";
            string resonId = "";
            query = "Select * FROM out_users where reason_date >= '" + startDate.Split(' ')[0] + "' AND reason_date <= '" + endDate.Split(' ')[0] + "' ";
            logger.Trace(query);
            using (MySqlDbReader mysqlDbTableReader = new MySqlDbReader(mysqlServerConnectionStringDB1))
            {
                MySql.Data.MySqlClient.MySqlDataReader mysqlData = mysqlDbTableReader.GetData(query);

                while (mysqlData.Read())
                {
                    if (mysqlData?.GetString(@"reason_id")?.Length > 0 && mysqlData?.GetString(@"user_code")?.Length > 0)
                    {
                        resonId = outResons.Find((x) => x._id == mysqlData.GetString(@"reason_id"))._id;
                        try { date = DateTime.Parse(mysqlData.GetString(@"reason_date")).ToYYYYMMDD(); } catch { date = ""; }

                        outPerson.Add(new OutPerson()
                        {
                            _reason_id = resonId,
                            _nav = mysqlData.GetString(@"user_code").ToUpper()?.Replace('C', 'S'),
                            _date = date,
                            _from = ConvertStringsTimeToSeconds(mysqlData.GetString(@"from_hour"), mysqlData.GetString(@"from_min")),
                            _to = ConvertStringsTimeToSeconds(mysqlData.GetString(@"to_hour"), mysqlData.GetString(@"to_min")),
                            _hourly = 0
                        });
                    }

                }
                logger.Trace("Всего с " + startDate.Split(' ')[0] + " по " + endDate.Split(' ')[0] + " на сайте есть - " + outPerson.Count + " записей с отсутствиями");
            }
            _ProgressWork1Step();

            if ((nameOfLastTable == "PeopleGroupDescription" || nameOfLastTable == "PeopleGroup" || nameOfLastTable == "Mailing" ||
                nameOfLastTable == "ListFIO" || doPostAction == "sendEmail") && nameGroup?.Length > 0)
            {
                _toolStripStatusLabelSetText(StatusLabel2, "Получаю данные по группе " + nameGroup);
                dtPeopleGroup = LoadGroupMembersFromDbToDataTable(nameGroup);

                logger.Trace("LoadRecords, DT - " + dtPeopleGroup.TableName + " , всего записей - " + dtPeopleGroup.Rows.Count);
                foreach (DataRow row in dtPeopleGroup.Rows)
                {
                    if (row[Names.FIO]?.ToString()?.Length > 0 && (row[Names.GROUP]?.ToString() == nameGroup || ("@" + row[Names.DEPARTMENT_ID]?.ToString()) == nameGroup))
                    {
                        person = new EmployeeFull()
                        {
                            fio = row[Names.FIO].ToString(),
                            code = row[Names.CODE]?.ToString(),
                            GroupPerson = nameGroup,
                            Department = row[Names.DEPARTMENT]?.ToString(),
                            City = row[Names.PLACE_EMPLOYEE]?.ToString(),
                            DepartmentBossCode = row[Names.CHIEF_ID]?.ToString(),
                            PositionInDepartment = row[Names.EMPLOYEE_POSITION]?.ToString(),
                            DepartmentId = row[Names.DEPARTMENT_ID]?.ToString(),
                            ControlInHHMM = row[Names.DESIRED_TIME_IN]?.ToString(),
                            ControlOutHHMM = row[Names.DESIRED_TIME_OUT]?.ToString(),
                            Comment = row[Names.EMPLOYEE_SHIFT_COMMENT]?.ToString(),
                            Shift = row[Names.EMPLOYEE_SHIFT]?.ToString()
                        };

                        person.ControlInSeconds = ConvertStringTimeHHMMSSToSeconds(person.ControlInHHMM);
                        person.ControlOutSeconds = ConvertStringTimeHHMMSSToSeconds(person.ControlOutHHMM);

                        GetPersonRegistrationFromServer(ref dtPersonRegistrationsFullList, person, startDate, endDate);     //Search Registration at checkpoints of the selected person
                    }
                }
                nameOfLastTable = "PeopleGroup";
                _toolStripStatusLabelSetText(StatusLabel2, "Данные по группе \"" + nameGroup + "\" получены");
            }
            else
            {
                person = new EmployeeFull();
                person.code = _ReturnTextOfControl(textBoxNav);
                person.fio = _ReturnTextOfControl(textBoxFIO);

                _toolStripStatusLabelSetText(StatusLabel2, "Получаю данные по \"" + ShortFIO(person.fio) + "\" ");

                person.GroupPerson = "One User";
                person.Department = "";
                person.DepartmentId = "";
                person.City = "";
                person.DepartmentBossCode = "";
                person.PositionInDepartment = "Сотрудник";

                person.Shift = "";
                person.Comment = "";

                person.ControlInHHMM = ConvertStringsTimeToStringHHMM(_numUpDownReturn(numUpDownHourStart).ToString(), _numUpDownReturn(numUpDownMinuteStart).ToString());
                person.ControlOutHHMM = ConvertStringsTimeToStringHHMM(_numUpDownReturn(numUpDownHourEnd).ToString(), _numUpDownReturn(numUpDownMinuteEnd).ToString());

                logger.Trace("LoadRecords,One User: " + person.fio);

                GetPersonRegistrationFromServer(ref dtPersonRegistrationsFullList, person, startDate, endDate);

                _toolStripStatusLabelSetText(StatusLabel2, "Данные с СКД по \"" + ShortFIO(_ReturnTextOfControl(textBoxFIO)) + "\" получены!");
            }
        }

        private void GetPersonRegistrationFromServer(ref DataTable dtTarget, EmployeeFull person, string startDay, string endDay)
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            SideOfPassagePoint sideOfPassagePoint;

            logger.Trace("GetPersonRegistrationFromServer, person - " + person.code);

            SeekAnualDays(ref dtTarget, ref person, false,
                ConvertStringDateToIntArray(startDay), ConvertStringDateToIntArray(endDay),
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
                query = "Select id, tabnum FROM OBJ_PERSON where tabnum like '" + person.code + "';";
                logger.Trace(query);

                //is looking for the idCard of the person's NAV
                using (SqlDbReader sqlDbTableReader = new SqlDbReader(sqlServerConnectionString))
                {
                    System.Data.SqlClient.SqlDataReader sqlData = sqlDbTableReader.GetData(query);
                    foreach (DbDataRecord record in sqlData)
                    {
                        if (record?["tabnum"]?.ToString()?.Trim() == person.code)
                        {
                            person.idCard = Convert.ToInt32(record["id"].ToString().Trim());
                            break;
                        }
                        _ProgressWork1Step();
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
                                seconds = ConvertStringTimeHHMMSSToSeconds(record["time"]?.ToString()?.Trim());

                                fullPointName = record["objid"]?.ToString()?.Trim();
                                sideOfPassagePoint = collectionOfPassagePoints.GetPoint(fullPointName);
                                namePoint = sideOfPassagePoint._namePoint;
                                direction = sideOfPassagePoint._direction;

                                personWorkedDays.Add(stringDataNew);

                                rowPerson = dtTarget.NewRow();
                                rowPerson[Names.FIO] = record["param0"].ToString().Trim();
                                rowPerson[Names.CODE] = person.code;
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
                                rowPerson[Names.REAL_TIME_IN] = ConvertSecondsToStringHHMMSS(seconds);

                                dtTarget.Rows.Add(rowPerson);

                                logger.Trace(rowPerson[Names.FIO] + " " + stringDataNew + " " + seconds + " " + namePoint + " " + direction);
                                _ProgressWork1Step();
                            }
                        }
                        catch (DbException dbexpt)
                        { logger.Warn(@"Ошибка доступа к данным БД: " + dbexpt.ToString()); }
                        catch (Exception err) { logger.Warn("GetPersonRegistrationFromServer: " + err.ToString()); }
                    }
                }

                stringDataNew = null; query = null;
                _ProgressWork1Step();
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
                rowPerson[Names.CODE] = person.code;
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
                _ProgressWork1Step();
            }
            dtTarget.AcceptChanges();

            foreach (DataRow row in dtTarget.Rows)
            {
                if (row[Names.CODE].ToString() == person.code)
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
                    if (row[Names.DATE_REGISTRATION]?.ToString() == day && row[Names.CODE]?.ToString() == person.code)
                    {
                        try
                        {
                            row[Names.EMPLOYEE_SHIFT_COMMENT] = outPerson.Find((x) => x._date == day && x._nav == person.code)._reason_id;
                            logger.Trace("GetPersonRegistrationFromServer, outPerson " + person.code + ", outReason - " + row[Names.EMPLOYEE_SHIFT_COMMENT].ToString());
                        }
                        catch { }
                        break;
                    }
                }
            }
            dtTarget.AcceptChanges();

            _ProgressWork1Step();

            rowPerson = null;
            namePoint = null; direction = null;
            cellData = new string[1];
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
        private DataTable LoadGroupMembersFromDbToDataTable(string namePointedGroup) // dtPeopleGroup //"Select * FROM PeopleGroup where GroupPerson like '" + _textBoxReturnText(textBoxGroup) + "';"
        {
            DataTable peopleGroup = dtPeople.Clone();
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            DataRow dataRow;

            string query = "Select FIO, NAV, GroupPerson, ControllingHHMM, ControllingOUTHHMM, Shift, Comment, Department, PositionInDepartment, DepartmentId, City, Boss  FROM PeopleGroup ";
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
            else if (EditAnualDaysItem.Text.Contains(@"Завершить редактирование"))
            {
                ExitEditAnual();
                EditAnualDaysItem.Font = new Font(this.Font, FontStyle.Regular);
                AddAnualDateItem.Font = new Font(this.Font, FontStyle.Regular);
            }
        }

        private void EnterEditAnual()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            ShowDataTableDbQuery(dbApplication, "BoldedDates",
                "SELECT DayBolded AS '" + Names.DAYOFF_DATE + "', DayType AS '" + Names.DAYOFF_TYPE + "', " +
                "NAV AS '" + Names.DAYOFF_USED_BY + "', DayDescription AS 'Описание', DateCreated AS '" + Names.DAYOFF_ADDED + "'",
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
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            comboBoxFio.Items?.RemoveAt(comboBoxFio.FindStringExact(""));
            if (comboBoxFio.Items.Count > 0)
            { comboBoxFio.SelectedIndex = 0; }

            _MenuItemEnabled(FunctionMenuItem, true);
            _MenuItemEnabled(GroupsMenuItem, true);
            _MenuItemEnabled(AddAnualDateItem, false);

            CheckBoxesFiltersAll_Visible(true);

            EditAnualDaysItem.Text = Names.DAY_OFF_OR_WORK;
            EditAnualDaysItem.ToolTipText = Names.DAY_OFF_OR_WORK_EDIT;

            toolTip1.SetToolTip(textBoxGroup, "Создать или добавить в группу");
            toolTip1.SetToolTip(textBoxGroupDescription, "Изменить описание группы");
            labelGroup.Text = "Группа";
            textBoxGroup.Text = "";

            _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);
            _toolStripStatusLabelSetText(StatusLabel2, @"Завершен 'Режим редактирования в локальной БД дат праздников и выходных'");

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
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

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

            if (dbApplication.Exists)
            {
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
        }

        private void DeleteAnualDateItem_Click(object sender, EventArgs e) //DeleteAnualDay()
        { DeleteAnualDay(); }

        private void DeleteAnualDay()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            DataGridViewOperations dgvo = new DataGridViewOperations();
            dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                 Names.DAYOFF_DATE, Names.DAYOFF_TYPE, Names.DAYOFF_USED_BY, Names.DAYOFF_ADDED });

            DeleteDataTableQueryParameters(dbApplication, "BoldedDates",
                "DayBolded", dgvo.cellValue[0], "DayType", dgvo.cellValue[1],
                "NAV", dgvo.cellValue[2], "DateCreated", dgvo.cellValue[3]).GetAwaiter().GetResult();

            ShowDataTableDbQuery(dbApplication, "BoldedDates", "SELECT DayBolded AS '" + Names.DAYOFF_DATE + "', DayType AS '" + Names.DAYOFF_TYPE + "', " +
            "NAV AS '" + Names.DAYOFF_USED_BY + "', DayDescription AS 'Описание', DateCreated AS '" + Names.DAYOFF_ADDED + "'",
            " ORDER BY DayBolded desc, NAV asc; ");
        }

        private void CheckBoxesFiltersAll_SetState(bool state)
        {
            _CheckboxCheckedSet(checkBoxTimeViolations, state);
            _CheckboxCheckedSet(checkBoxReEnter, state);
            _CheckboxCheckedSet(checkBoxCelebrate, state);
            _CheckboxCheckedSet(checkBoxWeekend, state);
        }

        private void CheckBoxesFiltersAll_Enable(bool state)
        {
            _ControlEnable(checkBoxTimeViolations, state);
            _ControlEnable(checkBoxReEnter, state);
            _ControlEnable(checkBoxCelebrate, state);
            _ControlEnable(checkBoxWeekend, state);
        }

        private void CheckBoxesFiltersAll_Visible(bool state)
        {
            _ControlVisible(checkBoxTimeViolations, state);
            _ControlVisible(checkBoxReEnter, state);
            _ControlVisible(checkBoxCelebrate, state);
            _ControlVisible(checkBoxWeekend, state);
        }

        private async void checkBox_CheckStateChanged(object sender, EventArgs e)
        { await Task.Run(() => checkBoxCheckStateChanged()); }

        private void checkBoxCheckStateChanged()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            DataTable dtEmpty = new DataTable();
            EmployeeFull emptyPerson = new EmployeeFull();
            SeekAnualDays(ref dtEmpty, ref emptyPerson, false,
                _dateTimePickerReturnArray(dateTimePickerStart),
                _dateTimePickerReturnArray(dateTimePickerEnd),
                ref myBoldedDates, ref workSelectedDays);

            dtEmpty.Dispose();

            CheckBoxesFiltersAll_Enable(false);
            _ControlVisible(dataGridView1, false);

            string nameGroup = _ReturnTextOfControl(textBoxGroup);

            //todo dubble
            // check need - DataTable dtTempIntermediate
            DataTable dtTempIntermediate = dtPeople.Clone();
            dtPersonTempAllColumns = dtPeople.Clone();
            EmployeeFull person = new EmployeeFull()
            {
                fio = _ReturnTextOfControl(textBoxFIO),
                code = _ReturnTextOfControl(textBoxNav),
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

                if (_CheckboxCheckedStateReturn(checkBoxReEnter))
                {
                    foreach (DataRow row in dtPeopleGroup.Rows)
                    {
                        if (row[Names.FIO]?.ToString()?.Length > 0 && (row[Names.GROUP]?.ToString() == nameGroup || (@"@" + row[Names.DEPARTMENT_ID]?.ToString()) == nameGroup))
                        {
                            person = new EmployeeFull
                            {
                                fio = row[Names.FIO].ToString(),
                                code = row[Names.CODE].ToString(),
                                GroupPerson = row[Names.GROUP].ToString(),
                                Department = row[Names.DEPARTMENT].ToString(),
                                PositionInDepartment = row[Names.EMPLOYEE_POSITION].ToString(),
                                City = row[Names.PLACE_EMPLOYEE]?.ToString(),
                                DepartmentId = row[Names.DEPARTMENT_ID].ToString(),
                                ControlInSeconds = ConvertStringTimeHHMMSSToSeconds(row[Names.DESIRED_TIME_IN].ToString()),
                                ControlOutSeconds = ConvertStringTimeHHMMSSToSeconds(row[Names.DESIRED_TIME_OUT].ToString()),
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
                if (!_CheckboxCheckedStateReturn(checkBoxReEnter))
                { dtTempIntermediate = dtPersonRegistrationsFullList.Copy(); }
                else
                { FilterRegistrationsOfPerson(ref person, dtPersonRegistrationsFullList, ref dtTempIntermediate); }
            }

            //Table with all columns
            dtPersonTempAllColumns = dtTempIntermediate.Copy();
            dtPersonTemp = LeaveAndOrderColumnsOfDataTable(dtTempIntermediate, Names.orderColumnsRegistrations);
            dtTempIntermediate = null;

            //show selected data  within the selected collumns   
            ShowDatatableOnDatagridview(dtPersonTemp, "PeopleGroup");

            //change enabling of checkboxes
            if (_CheckboxCheckedStateReturn(checkBoxReEnter))// if (checkBoxReEnter.Checked)
            {
                _ControlEnable(checkBoxTimeViolations, true);
                _ControlEnable(checkBoxWeekend, true);
                _ControlEnable(checkBoxCelebrate, true);

                if (_CheckboxCheckedStateReturn(checkBoxTimeViolations))  // if (checkBoxStartWorkInTime.Checked)
                { _MenuItemBackColorChange(LoadDataItem, SystemColors.Control); }
            }
            else if (!_CheckboxCheckedStateReturn(checkBoxReEnter))
            {
                _CheckboxCheckedSet(checkBoxTimeViolations, false);
                _CheckboxCheckedSet(checkBoxWeekend, false);
                _CheckboxCheckedSet(checkBoxCelebrate, false);
                _ControlEnable(checkBoxTimeViolations, false);
                _ControlEnable(checkBoxWeekend, false);
                _ControlEnable(checkBoxCelebrate, false);
            }

            panelViewResize(numberPeopleInLoading);
            _ControlVisible(dataGridView1, true);
            _ControlEnable(checkBoxReEnter, true);
        }

        private DataTable LeaveAndOrderColumnsOfDataTable(DataTable dt, string[] columns)
        {
            DataTable dtUniqRecords = new DataTable();
            dtUniqRecords = dt.DefaultView.ToTable(true, columns);

            dtUniqRecords.SetColumnsOrder(columns);
            return dtUniqRecords;
        }

        private void FilterRegistrationsOfPerson(ref EmployeeFull person, DataTable dataTableSource, ref DataTable dataTableForStoring, string typeReport = "Полный")
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            logger.Trace("FilterRegistrationsOfPerson: " + person.code + "| dataTableSource: " + dataTableSource.Rows.Count, "| typeReport: " + typeReport);
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
                var allWorkedDaysPerson = dataTableSource.Select("[NAV-код] = '" + person.code + "'");

                if (_CheckboxCheckedStateReturn(checkBoxReEnter) || currentAction == "sendEmail") //checkBoxReEnter.Checked
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
                        rowDtStoring[Names.REAL_TIME_IN] = ConvertSecondsToStringHHMMSS(firstRegistrationInDay);  //("Фактич. время прихода ЧЧ:ММ:СС", typeof(string)),//24

                        // rowDtStoring[@"Реальное время ухода"] = lastRegistrationInDay;                 //("Реальное время ухода", typeof(decimal)), //18
                        rowDtStoring[Names.REAL_TIME_OUT] = ConvertSecondsToStringHHMMSS(lastRegistrationInDay);     //("Фактич. время ухода ЧЧ:ММ", typeof(string)), //25

                        //worked out times
                        workedSeconds = lastRegistrationInDay - firstRegistrationInDay;
                        rowDtStoring[Names.EMPLOYEE_TIME_SPENT] = workedSeconds;                                  // ("Реальное отработанное время", typeof(decimal)), //26
                        //  rowDtStoring[@"Отработанное время ЧЧ:ММ"] = ConvertSecondsToStringHHMMSS(workedSeconds);  //("Отработанное время ЧЧ:ММ", typeof(string)), //27
                        logger.Trace("FilterRegistrationsOfPerson: " + person.code + "| " + rowDtStoring[Names.DATE_REGISTRATION]?.ToString() + " " + firstRegistrationInDay + " - " + lastRegistrationInDay);

                        //todo 
                        //will calculate if day of week different
                        if (firstRegistrationInDay > (person.ControlInSeconds + offsetTimeIn) && firstRegistrationInDay != 0) // "Опоздание ЧЧ:ММ", typeof(bool)),           //28
                        {
                            if (typeReport == "Полный")
                            { rowDtStoring[Names.EMPLOYEE_BEING_LATE] = ConvertSecondsToStringHHMMSS(firstRegistrationInDay - person.ControlInSeconds); }
                            else if (typeReport == "Упрощенный")
                            { rowDtStoring[Names.EMPLOYEE_BEING_LATE] = "1"; }
                        }

                        if (lastRegistrationInDay < (person.ControlOutSeconds - offsetTimeOut) && lastRegistrationInDay != 0)  // "Ранний уход ЧЧ:ММ", typeof(bool)),                 //29
                        {
                            if (typeReport == "Полный")
                            { rowDtStoring[Names.EMPLOYEE_EARLY_DEPARTURE] = ConvertSecondsToStringHHMMSS(person.ControlOutSeconds - lastRegistrationInDay); }
                            else if (typeReport == "Упрощенный")
                            { rowDtStoring[Names.EMPLOYEE_EARLY_DEPARTURE] = "1"; }
                        }

                        if (rowDtStoring[Names.EMPLOYEE_ABSENCE]?.ToString() == "1" && typeReport == "Полный")  // "Ранний уход ЧЧ:ММ", typeof(bool)),                 //29
                        {
                            rowDtStoring[Names.EMPLOYEE_ABSENCE] = "Да";
                        }

                        exceptReason = rowDtStoring[Names.EMPLOYEE_SHIFT_COMMENT]?.ToString();

                        rowDtStoring[Names.EMPLOYEE_SHIFT_COMMENT] = outResons.Find((x) => x._id == exceptReason)?._name;

                        switch (exceptReason)
                        {
                            case "2": //Отпуск
                            case "10": //Отпуск по беременности и родам
                            case "11": //Отпуск по уходу за ребёнком
                                rowDtStoring[Names.EMPLOYEE_EARLY_DEPARTURE] = "";
                                rowDtStoring[Names.EMPLOYEE_BEING_LATE] = "";
                                if (typeReport == "Полный")
                                { rowDtStoring[@"Отпуск"] = outResons.Find((x) => x._id == exceptReason)?._name; }
                                else if (typeReport == "Упрощенный")
                                { rowDtStoring[@"Отпуск"] = "1"; }
                                break;
                            case "3": //Больничный
                            case "21": //Больничный 0,5 (менее < 5 часов)
                                rowDtStoring[Names.EMPLOYEE_EARLY_DEPARTURE] = "";
                                rowDtStoring[Names.EMPLOYEE_BEING_LATE] = "";
                                if (typeReport == "Полный")
                                { rowDtStoring[Names.EMPLOYEE_SICK_LEAVE] = outResons.Find((x) => x._id == exceptReason)?._name; }
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
                                { rowDtStoring[Names.EMPLOYEE_HOOKY] = outResons.Find((x) => x._id == exceptReason)?._name; }
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

                                rowDtStoring[Names.EMPLOYEE_SHIFT_COMMENT] = outResons.Find((x) => x._id == exceptReason)?._name;
                                rowDtStoring[Names.EMPLOYEE_EARLY_DEPARTURE] = "";
                                rowDtStoring[Names.EMPLOYEE_BEING_LATE] = "";
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
                { QueryDeleteDataFromDataTable(ref dtTemp, "[Опоздание ЧЧ:ММ]='' AND [Ранний уход ЧЧ:ММ]=''", person.code); }

                foreach (DataRow dr in dtTemp.AsEnumerable())
                { dataTableForStoring.ImportRow(dr); }

                allWorkedDaysPerson = null;
            }
            catch (Exception err) { MessageBox.Show(err.ToString()); }

            hsDays = null;
            rowDtStoring = null;
            dtTemp = null;
            exceptReason = null;
            dtAllRegistrationsInSelectedDay = null;
        }


        private void SeekAnualDays(ref DataTable dt, ref EmployeeFull person, bool delRow, int[] startOfPeriod, int[] endOfPeriod, ref string[] boldedDays, ref string[] workDays)//   //Exclude Anual Days from the table "PersonTemp" DB
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            if (person == null)
            { person = new EmployeeFull(); }
            if (person.code == null)
            { person.code = "0"; }

            List<string> daysBolded = new List<string>();
            List<string> daysListBolded = new List<string>();
            List<string> daysListWorkedInDB = new List<string>();

            var oneDay = TimeSpan.FromDays(1);
            var twoDays = TimeSpan.FromDays(2);

            var mySelectedStartDay = new DateTime(startOfPeriod[0], startOfPeriod[1], startOfPeriod[2]);
            var mySelectedEndDay = new DateTime(endOfPeriod[0], endOfPeriod[1], endOfPeriod[2]);
            var myMonthCalendar = new MonthCalendar();

            myMonthCalendar.MaxSelectionCount = 60;
            myMonthCalendar.SelectionRange = new SelectionRange(mySelectedStartDay, mySelectedEndDay);

            //whole range of days in the selection period
            List<string> wholeSelectedDays = new List<string>();
            for (var myDate = myMonthCalendar.SelectionStart; myDate <= myMonthCalendar.SelectionEnd; myDate += oneDay)
            {
                wholeSelectedDays.Add(myDate.ToYYYYMMDD());
            }
            wholeSelectedDays.Sort();

            logger.Trace("SeekAnualDays,start-end: " + person.code + " - " +
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

            List<string> days = ReturnBoldedDaysFromDB(person.code, @"Выходной");
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
            days = ReturnBoldedDaysFromDB(person.code, @"Рабочий");
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
                    QueryDeleteDataFromDataTable(ref dt, "[Дата регистрации]='" + day + "'", person.code); // ("Дата регистрации",typeof(string)),//12
                }
            }

            if (dt != null)
            { dt.AcceptChanges(); }

            daysBolded.Sort();

            if (person == null || person.code == "0")
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
            string query = null;
            if (nav.Length == 6)
            {
                query = "SELECT DayBolded FROM BoldedDates WHERE (NAV LIKE '" + nav + "' OR  NAV LIKE '0') AND DayType LIKE '" + dayType + "';";
            }
            else
            {
                query = "SELECT DayBolded FROM BoldedDates WHERE (NAV LIKE '0') AND DayType LIKE '" + dayType + "';";
            }

            using (SqLiteDbReader dbReader = new SqLiteDbReader(sqLiteLocalConnectionString, dbApplication))
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
                logger.Trace("ReturnBoldedDaysFromDB: query: " + query + "\n" + boldedDays.Count + " rows loaded");
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
            rows = null;
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
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

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
                        logger.Trace(dirPath + " - " + err.ToString());
                    }
                }
            }
            foreach (var file in files)
            {
                try
                {
                    System.IO.File.Copy(file, appFolderTempPath + @"\" + file, true);
                }
                catch (Exception err) { logger.Trace(file + " - " + err.ToString()); }
            }
            System.IO.Compression.ZipFile.CreateFromDirectory(appFolderTempPath, localAppFolderPath + @"\" + fullNameZip, System.IO.Compression.CompressionLevel.Optimal, false);
            LoggerAddInfo("Архив создан: " + localAppFolderPath + @"\" + fullNameZip);
        }

        private void MakeZip(string filePath, string fullNameZip)
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            if (filePath.Contains(@"\"))
            {
                try { System.IO.Directory.CreateDirectory(filePath.Replace(filePath, appFolderTempPath + @"\" + filePath.Remove(filePath.IndexOf('\\')))); }
                catch (Exception err) { logger.Trace(filePath + " - " + err.ToString()); }
            }
            try
            {
                System.IO.File.Copy(filePath, appFolderTempPath + @"\" + filePath, true);
            }
            catch (Exception err) { logger.Trace(filePath + " - " + err.ToString()); }

            System.IO.Compression.ZipFile.CreateFromDirectory(appFolderTempPath, localAppFolderPath + @"\" + fullNameZip, System.IO.Compression.CompressionLevel.Optimal, false);
            LoggerAddInfo("Made archive: " + localAppFolderPath + @"\" + fullNameZip);
        }

        //----- Clearing. Start ---------//
        private void ClearItemsInApplicationFolders(string maskFiles)
        {
            System.IO.FileInfo[] filesPath = null;
            bool filesInFolderExist = false;
            string folder = System.IO.Path.GetDirectoryName(maskFiles);
            if (System.IO.Directory.Exists(maskFiles))
            {
                filesInFolderExist = true;
            }

                if (filesInFolderExist)
                {
                    //   filesPath = new System.IO.DirectoryInfo(maskFiles).GetFiles(@"*.*", System.IO.SearchOption.AllDirectories);
                    logger.Trace("Files '" + maskFiles + "' in folder exist:" + filesInFolderExist);
                    var dir = new System.IO.DirectoryInfo(maskFiles);
                    try { dir.Delete(true); }
                    catch (Exception e) { logger.Warn("Папка не удалена: " + maskFiles + " " + e.Message); }
                    try { System.IO.Directory.CreateDirectory(maskFiles); }
                    catch (Exception e) { logger.Warn("Папка не создана: " + maskFiles + " " + e.Message); }
                }
                else
                {
                    filesPath = new System.IO.DirectoryInfo(folder).GetFiles(
                        maskFiles.Remove(0, maskFiles.LastIndexOf(@"\") + 1), 
                        System.IO.SearchOption.AllDirectories);
                }
                foreach (var file in filesPath)
                {
                    try
                    {
                        file.Delete();
                        logger.Trace("Удален файл: " + file.FullName );
                    }
                    catch (Exception e) { logger.Warn("Файл '" + file.Name + "'не удален из-за ошибки: " + e.Message); }
                }
        }

        private async void ClearReportItem_Click(object sender, EventArgs e) //ReCreatePersonTables()
        {
            await Task.Run(() => CleaReportsRecreateTables());
        }

        private async void CleaReportsRecreateTables()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            logger.Info("-= Очистика от отчетов =-");

            ClearItemsInApplicationFolders(@"*.xlsx");

            _SetTextOfControl(textBoxFIO, "");
            _SetTextOfControl(textBoxGroup, "");
            _SetTextOfControl(textBoxGroupDescription, "");
            _SetTextOfControl(textBoxNav, "");

            GC.Collect();

            TryMakeLocalDB();

            DataTable dt = new DataTable();
            dgvo.ShowData(dataGridView1, dt);

            _toolStripStatusLabelSetText(StatusLabel2, @"Временные таблицы удалены");
        }

        private async void ClearDataItem_Click(object sender, EventArgs e) //ReCreateAllPeopleTables()
        {
            await Task.Run(() => ClearGotReportsRecreateTables());
        }

        private async void ClearGotReportsRecreateTables()
        {
            logger.Trace("-= ClearGotReportsRecreateTables =-");
            logger.Info("-= Очистика от отчетов и полученных данных =-");

            ExecuteQueryOnLocalDB(dbApplication, "DELETE FROM 'LastTakenPeopleComboList';").GetAwaiter().GetResult();

            ClearItemsInApplicationFolders(@"*.xlsx");
            ClearItemsInApplicationFolders(@"*.log");

            _SetTextOfControl(textBoxFIO, "");
            _SetTextOfControl(textBoxGroup, "");
            _SetTextOfControl(textBoxGroupDescription, "");
            _SetTextOfControl(textBoxNav, "");

            _comboBoxClr(comboBoxFio);

            TryMakeLocalDB();

            using (DataTable dt = new DataTable())
            {
                dgvo.ShowData(dataGridView1, dt);
            }

            _toolStripStatusLabelSetText(StatusLabel2, @"База очищена от загруженных данных. Остались только созданные группы");
        }

        private async void ClearAllItem_Click(object sender, EventArgs e) //ReCreate DB
        {
            await Task.Run(() => ReCreateDB());
        }

        private void VacuumDB(string dbPath)
        {
            SQLiteConnectionStringBuilder builder =
                new SQLiteConnectionStringBuilder();
            builder.DataSource = dbPath;
            builder.PageSize = 4096;
            builder.UseUTF16Encoding = true;

            using (SQLiteConnection conn = new SQLiteConnection(builder.ConnectionString))
            {
                conn.Open();

                SQLiteCommand vacuum = new SQLiteCommand(@"VACUUM", conn);
                vacuum.ExecuteNonQuery();
            }
        }

        private async void ReCreateDB()
        {
            logger.Trace("-= ReCreateDB =-");

            if (dbApplication.Exists)
            {
                logger.Info("-= Очистика локальной базы от всех полученных, сгенерированных, сохраненных и введенных данных =-");

                GC.Collect();

                ExecuteQueryOnLocalDB(dbApplication, "DROP Table if exists 'PeopleGroup';").GetAwaiter().GetResult();
                ExecuteQueryOnLocalDB(dbApplication, "DROP Table if exists 'PeopleGroupDescription';").GetAwaiter().GetResult();
                ExecuteQueryOnLocalDB(dbApplication, "DROP Table if exists 'TechnicalInfo';").GetAwaiter().GetResult();
                ExecuteQueryOnLocalDB(dbApplication, "DROP Table if exists 'BoldedDates';").GetAwaiter().GetResult();
                ExecuteQueryOnLocalDB(dbApplication, "DROP Table if exists 'ConfigDB';").GetAwaiter().GetResult();
                ExecuteQueryOnLocalDB(dbApplication, "DROP Table if exists 'LastTakenPeopleComboList';").GetAwaiter().GetResult();

                VacuumDB(sqLiteLocalConnectionString);

                ClearItemsInApplicationFolders(@"*.xlsx");
                ClearItemsInApplicationFolders(@"*.log");

                _SetTextOfControl(textBoxFIO, "");
                _SetTextOfControl(textBoxGroup, "");
                _SetTextOfControl(textBoxGroupDescription, "");
                _SetTextOfControl(textBoxNav, "");

                _comboBoxClr(comboBoxFio);

                TryMakeLocalDB();
            }
            else
            {
                TryMakeLocalDB();
            }

            dgvo.ShowData(dataGridView1, null);

            _toolStripStatusLabelSetText(StatusLabel2, @"Все таблицы очищены");
        }

        private void ClearRegistryItem_Click(object sender, EventArgs e) //ClearRegistryData()
        { ClearRegistryData(); }

        private void ClearRegistryData()
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
                _toolStripStatusLabelSetText(StatusLabel2,
                    "Ошибки с доступом у реестру на запись. Данные не удалены.",
                    true, "| ClearRegistryData: " + err.ToString());
            }
        }
        //----- Clearing. End ---------//



        //gathering a person's features from textboxes and other controls
        private void SelectPersonFromControls(ref EmployeeFull personSelected)
        {
            personSelected.fio = _ReturnTextOfControl(textBoxFIO);
            personSelected.code = _ReturnTextOfControl(textBoxNav);
            personSelected.GroupPerson = _ReturnTextOfControl(textBoxGroup);

            personSelected.ControlInHHMM = ConvertDecimalTimeToStringHHMM(_numUpDownReturn(numUpDownHourStart), _numUpDownReturn(numUpDownMinuteStart));
            personSelected.ControlOutHHMM = ConvertDecimalTimeToStringHHMM(_numUpDownReturn(numUpDownHourEnd), _numUpDownReturn(numUpDownMinuteEnd));
        }



        //---- Start. Drawing ---//

        private void VisualItem_Click(object sender, EventArgs e) //FindWorkDaysInSelected() , DrawFullWorkedPeriodRegistration()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            EmployeeFull personVisual = new EmployeeFull();

            decimal[] timeIn = new decimal[5];
            decimal[] timeOut = new decimal[5];

            try
            {
                dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
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

                if (dgvo.RowsCount(dataGridView1) > -1)
                {
                    personVisual.fio = dgvo.cellValue[0];
                    personVisual.code = dgvo.cellValue[1]; //Take the name of selected group
                    personVisual.ControlInHHMM = dgvo.cellValue[3]; //Take the name of selected group

                    timeIn = ConvertStringTimeHHMMToDecimalArray(personVisual.ControlInHHMM);
                    personVisual.ControlInSeconds = (int)timeIn[4];

                    personVisual.ControlOutHHMM = dgvo.cellValue[4]; //Take the name of selected group
                    timeOut = ConvertStringTimeHHMMToDecimalArray(personVisual.ControlOutHHMM);

                    personVisual.ControlOutSeconds = (int)timeOut[4];
                    _numUpDownSet(numUpDownHourStart, timeIn[0]);
                    _numUpDownSet(numUpDownMinuteStart, timeIn[1]);

                    _numUpDownSet(numUpDownHourEnd, timeOut[0]);
                    _numUpDownSet(numUpDownMinuteEnd, timeOut[1]);

                    personVisual.Department = dgvo.cellValue[5];
                    personVisual.PositionInDepartment = dgvo.cellValue[6];
                    personVisual.Shift = dgvo.cellValue[7];
                    personVisual.DepartmentId = dgvo.cellValue[8];
                    personVisual.Comment = dgvo.cellValue[9];

                    if (nameOfLastTable == "PeopleGroup" || nameOfLastTable == "ListFIO")
                    {
                        personVisual.GroupPerson = dgvo.cellValue[2]; //Take the name of selected group
                        StatusLabel2.Text = @"Выбрана группа: " + personVisual.GroupPerson + @" | Курсор на: " + personVisual.fio;
                        groupBoxPeriod.BackColor = Color.PaleGreen;
                    }
                    else if (nameOfLastTable == "PersonRegistrationsList")
                    {
                        personVisual.GroupPerson = _ReturnTextOfControl(textBoxGroup);
                        StatusLabel2.Text = @"Выбран: " + personVisual.fio;
                    }
                }
            }
            catch (Exception err) { logger.Info("VisualItem_Click: " + err.ToString()); }

            if (personVisual.fio.Length == 0)
            {
                SelectPersonFromControls(ref personVisual);
            }

            _ControlVisible(dataGridView1, false);

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

            _ControlVisible(panelView, true);
            _ControlVisible(pictureBox1, true);
        }

        private void DrawRegistration(ref EmployeeFull personDraw)  // Visualisation of registration
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            //  int iPanelBorder = 2;
            int iShiftStart = 300;
            int iShiftHeightAll = 36;

            int iOffsetBetweenHorizontalLines = 19; //смещение между горизонтальными линиями
            int iOffsetBetweenVerticalLines = 60; //смещение между "часовыми" линиями
            int iNumbersOfHoursInDay = 24;        //количество часов в сутках(количество вертикальных часовых линий)

            int iHeightLineWork = 4; //толщина линии рабочего времени на графике
            int pointDrawYfor_rects = 44; //начальное смещение линии рабочего графика

            int iHeightLineRealWork = 14; //толщина линии фактически отработанного веремени на графике
            int pointDrawYfor_rectsReal = 39; // начальное смещение линии отработанного графика

            int iWidthRects = 2; // ширина прямоугольников = время нахождение в рабочей зоне(минимальное)
            int iLenghtRect = 0; //количество  входов-выходов в рабочие дни для всех отобранных людей для  анализа регистраций входа-выхода

            //constant for a person
            string fio = personDraw.fio;
            string nav = personDraw.code;
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
            { hsNAV.Add(row[Names.CODE].ToString()); }
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
            _panelResume(panelView);

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
                var myBrushWorkHour = new SolidBrush(Color.Gray);
                var myBrushRealWorkHour = new SolidBrush(clrRealRegistration);
                var myBrushAxis = new SolidBrush(Color.Black);
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
                            workDay + " (" + ShortFIO(fio) + ")",
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
        }

        private void DrawFullWorkedPeriodRegistration(ref EmployeeFull personDraw)  // Draw the whole period registration
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            //  int iPanelBorder = 2;
            int iShiftStart = 300;
            int iShiftHeightAll = 36;

            int iOffsetBetweenHorizontalLines = 19; //смещение между горизонтальными линиями
            int iOffsetBetweenVerticalLines = 60; //смещение между "часовыми" линиями
            int iNumbersOfHoursInDay = 24;        //количество часов в сутках(количество вертикальных часовых линий)

            int iHeightLineWork = 4; //толщина линии рабочего времени на графике
            int pointDrawYfor_rects = 44; //начальное смещение линии рабочего графика

            int iHeightLineRealWork = 14; //толщина линии фактически отработанного веремени на графике
            int pointDrawYfor_rectsReal = 39; // начальное смещение линии отработанного графика

            int iLenghtRect = 0; //количество  входов-выходов в рабочие дни для всех отобранных людей для  анализа регистраций входа-выхода

            //constant for a person
            string fio = personDraw.fio;
            string nav = personDraw.code;
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
            { hsNAV.Add(row[Names.CODE].ToString()); }
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
            _panelResume(panelView);

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
                var myBrushWorkHour = new SolidBrush(Color.Gray);
                var myBrushRealWorkHour = new SolidBrush(clrRealRegistration);
                var myBrushAxis = new SolidBrush(Color.Black);
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
            _ControlVisible(pictureBox1, false);

            try
            {
                if (panelView?.Controls?.Count > 1) panelView.Controls.RemoveAt(1);
                bmp?.Dispose();
                pictureBox1?.Dispose();
            }
            catch { }

            _ControlVisible(dataGridView1, true);

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
            int iOffsetBetweenHorizontalLines = 19;
            int iShiftHeightAll = 36;
            switch (sLastSelectedElement)
            {
                case "DrawFullWorkedPeriodRegistration":
                    _panelSetHeight(panelView, iShiftHeightAll + iOffsetBetweenHorizontalLines * workSelectedDays.Length * numberPeople); //Fixed size of Picture. If need autosize - disable this row
                    break;
                case "DrawRegistration":
                    _panelSetHeight(panelView, iShiftHeightAll + iOffsetBetweenHorizontalLines * workSelectedDays.Length * numberPeople); //Fixed size of Picture. If need autosize - disable this row
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

            if (dbApplication.Exists)
            {
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
            DialogResult result = MessageBox.Show(
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
                 @"\n\nДата и время локального ПК: " + _dateTimePickerReturnString(dateTimePickerEnd),

                "Информация о программе",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1);

        }

        private void PrepareForMakingFormMailing(object sender, EventArgs e) //MailingItem()
        {
            nameOfLastTable = "Mailing";
            _MenuItemEnabled(SettingsMenuItem, false);
            _MenuItemEnabled(FunctionMenuItem, false);
            _MenuItemEnabled(GroupsMenuItem, false);
            CheckBoxesFiltersAll_Enable(false);
            _ControlVisible(panelView, false);

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
            using (var sqlConnection = new SQLiteConnection(sqLiteLocalConnectionString))
            {
                sqlConnection.Open();
                string group;
                using (var sqlCommand = new SQLiteCommand(
                    "SELECT GroupPerson FROM PeopleGroupDescription;", sqlConnection))
                {
                    using (var sqlReader = sqlCommand.ExecuteReader())
                    {
                        foreach (DbDataRecord record in sqlReader)
                        {
                            group = record["GroupPerson"]?.ToString();
                            if (group?.Length > 0)
                            {
                                listComboParameters.Add(group);
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
                    "День подготовки отчета", "", "День, в который выполнять подготовку и отправку данного отчета./nНачало, Средина, Конец"
                    );
        }

        private void SaveMailing(string recipientEmail, string senderEmail, string groupsReport, string nameReport, string descriptionReport,
            string periodPreparing, string status, string dateCreatingMailing, string SendingDate, string typeReport, string daySendingReport)
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            bool recipientValid = false;
            bool senderValid = false;

            if (recipientEmail?.Length > 0 && recipientEmail.Contains('.') && recipientEmail.Contains('@') && recipientEmail?.Split('.').Count() > 1)
            { recipientValid = true; }

            if (senderEmail?.Length > 0 && senderEmail.Contains('.') && senderEmail.Contains('@') && senderEmail?.Split('.').Count() > 1)
            { senderValid = true; }

            if (dbApplication.Exists && nameReport?.Length > 0 && senderValid && recipientValid)
            {
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
                _toolStripStatusLabelSetText(StatusLabel2, "Добавлена рассылка: " + nameReport + "| Всего рассылок: " + dgvo.RowsCount(dataGridView1));
            }
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
            _ControlVisible(panelView, false);

            btnPropertiesSave.Text = "Сохранить настройки";
            RemoveClickEvent(btnPropertiesSave);
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
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

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
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

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
            _ControlVisible(panelView, true);
        }

        private void DisposeTemporaryControls()
        {
            _ControlVisible(groupBoxProperties, false);
            _ControlDispose(labelServer1);
            _ControlDispose(labelServer1UserName);
            _ControlDispose(labelServer1UserPassword);
            _ControlDispose(labelMailServerName);
            _ControlDispose(labelMailServerUserName);
            _ControlDispose(labelMailServerUserPassword);
            _ControlDispose(labelmysqlServer);
            _ControlDispose(labelmysqlServerUserName);
            _ControlDispose(labelmysqlServerUserPassword);

            _ControlDispose(textBoxServer1);
            _ControlDispose(textBoxServer1UserName);
            _ControlDispose(textBoxServer1UserPassword);
            _ControlDispose(textBoxMailServerName);
            _ControlDispose(textBoxMailServerUserName);
            _ControlDispose(textBoxMailServerUserPassword);
            _ControlDispose(textBoxmysqlServer);
            _ControlDispose(textBoxmysqlServerUserName);
            _ControlDispose(textBoxmysqlServerUserPassword);

            _ControlDispose(listComboLabel);
            _ControlDispose(periodComboLabel);
            _ControlDispose(labelSettings9);

            _ControlDispose(listCombo);
            _ControlDispose(periodCombo);
            _ControlDispose(comboSettings9);

            _ControlDispose(labelSettings15);
            _ControlDispose(comboSettings15);

            _ControlDispose(labelSettings16);
            _ControlDispose(textBoxSettings16);
        }

        private void buttonPropertiesSave_Click(object sender, EventArgs e) //SaveProperties()
        {
            SaveProperties(); //btnPropertiesSave 

            DisposeTemporaryControls();
            EnableMainMenuItems(true);
            _ControlVisible(panelView, true);
        }

        private void ButtonPropertiesSave_MailingSave(object sender, EventArgs e)
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            string recipientEmail = _ReturnTextOfControl(textBoxServer1UserName);
            string senderEmail = mailSenderAddress;
            if (mailSenderAddress.Length == 0)
            { senderEmail = _ReturnTextOfControl(textBoxServer1); }
            string nameReport = _ReturnTextOfControl(textBoxMailServerName);
            string description = _ReturnTextOfControl(textBoxMailServerUserName);
            string report = _comboBoxReturnSelected(listCombo);
            string period = _listBoxReturnSelected(periodCombo);
            string status = _comboBoxReturnSelected(comboSettings9);
            string typeReport = _comboBoxReturnSelected(comboSettings15);
            string dayReport = _ReturnTextOfControl(textBoxSettings16);

            if (recipientEmail.Length > 5 && nameReport.Length > 0)
            {
                SaveMailing(recipientEmail, senderEmail,
                    report, nameReport, description, period, status,
                    DateTime.Now.ToYYYYMMDDHHMM(), "", typeReport, dayReport);
            }

            ShowDataTableDbQuery(dbApplication, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
            "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
            "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
            " ORDER BY RecipientEmail asc, DateCreated desc; ");

            DisposeTemporaryControls();
            EnableMainMenuItems(true);
            _ControlVisible(panelView, true);
        }

        private void SaveProperties() //Save Parameters into Registry and variables
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            string server = _ReturnTextOfControl(textBoxServer1);
            string user = _ReturnTextOfControl(textBoxServer1UserName);
            string password = _ReturnTextOfControl(textBoxServer1UserPassword);

            string sMailServer = _ReturnTextOfControl(textBoxMailServerName);
            string sMailUser = _ReturnTextOfControl(textBoxMailServerUserName);
            string sMailUserPassword = _ReturnTextOfControl(textBoxMailServerUserPassword);

            string sMySqlServer = _ReturnTextOfControl(textBoxmysqlServer);
            string sMySqlServerUser = _ReturnTextOfControl(textBoxmysqlServerUserName);
            string sMySqlServerUserPassword = _ReturnTextOfControl(textBoxmysqlServerUserPassword);

            CheckAliveIntellectServer(server, user, password).GetAwaiter().GetResult();

            if (bServer1Exist)
            {
                _ControlVisible(groupBoxProperties, false);
                _MenuItemEnabled(GetFioItem, true);

                sServer1 = server;
                sServer1UserName = user;
                sServer1UserPassword = password;

                mailServer = sMailServer;
                mailSenderAddress = sMailUser;
                mailsOfSenderOfPassword = sMailUserPassword;

                mysqlServer = sMySqlServer;
                mysqlServerUserName = sMySqlServerUser;
                mysqlServerUserPassword = sMySqlServerUserPassword;

                try
                {
                    using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(appRegistryKey))
                    {
                        try { EvUserKey.SetValue("SKDServer", sServer1, Microsoft.Win32.RegistryValueKind.String); } catch { }
                        try { EvUserKey.SetValue("SKDUser", EncryptionDecryptionCriticalData.EncryptStringToBase64Text(sServer1UserName, keyEncryption, keyDencryption), Microsoft.Win32.RegistryValueKind.String); } catch { }
                        try { EvUserKey.SetValue("SKDUserPassword", EncryptionDecryptionCriticalData.EncryptStringToBase64Text(sServer1UserPassword, keyEncryption, keyDencryption), Microsoft.Win32.RegistryValueKind.String); } catch { }

                        try { EvUserKey.SetValue("MySQLServer", mysqlServer, Microsoft.Win32.RegistryValueKind.String); } catch { }
                        try { EvUserKey.SetValue("MySQLUser", mysqlServerUserName, Microsoft.Win32.RegistryValueKind.String); } catch { }
                        try { EvUserKey.SetValue("MySQLUserPassword", EncryptionDecryptionCriticalData.EncryptStringToBase64Text(mysqlServerUserPassword, keyEncryption, keyDencryption), Microsoft.Win32.RegistryValueKind.String); } catch { }

                        logger.Info("CreateSubKey: Данные в реестре сохранены");
                    }
                }
                catch (Exception err) { logger.Error("CreateSubKey: Ошибки с доступом на запись в реестр. Данные сохранены не корректно. " + err.ToString()); }

                {
                    string resultSaving = "";
                    ParameterOfConfigurationInSQLiteDB parameterSQLite = new ParameterOfConfigurationInSQLiteDB(dbApplication);

                    ParameterOfConfiguration parameterOfConfiguration = new ParameterOfConfigurationBuilder().
                        SetParameterName("SKDServer").
                        SetParameterValue(sServer1).
                        SetParameterDescription("URI SKD-сервера").
                        IsPassword(false).
                        SetIsExample("no");
                    resultSaving += parameterSQLite.SaveParameter(parameterOfConfiguration) + "\n";
                    parameterOfConfiguration = new ParameterOfConfigurationBuilder().
                        SetParameterName("SKDUser").
                        SetParameterValue(sServer1UserName).
                        SetParameterDescription("SKD MSSQL User's Name").
                        IsPassword(true).
                        SetIsExample("no");
                    resultSaving += parameterSQLite.SaveParameter(parameterOfConfiguration) + "\n";
                    parameterOfConfiguration = new ParameterOfConfigurationBuilder().
                        SetParameterName("SKDUserPassword").
                        SetParameterValue(sServer1UserPassword).
                        SetParameterDescription("SKD MSSQL User's Password").
                        IsPassword(true).
                        SetIsExample("no");
                    resultSaving += parameterSQLite.SaveParameter(parameterOfConfiguration) + "\n";


                    parameterOfConfiguration = new ParameterOfConfigurationBuilder().
                        SetParameterName("MailServer").
                        SetParameterValue(mailServer).
                        SetParameterDescription("URI Mail-серверa").
                        IsPassword(false).
                        SetIsExample("no");
                    resultSaving += parameterSQLite.SaveParameter(parameterOfConfiguration) + "\n";
                    parameterOfConfiguration = new ParameterOfConfigurationBuilder().
                        SetParameterName("MailUser").
                        SetParameterValue(mailSenderAddress).
                        SetParameterDescription("Senders E-Mail").
                        IsPassword(false).
                        SetIsExample("no");
                    resultSaving += parameterSQLite.SaveParameter(parameterOfConfiguration) + "\n";
                    parameterOfConfiguration = new ParameterOfConfigurationBuilder().
                        SetParameterName("MailUserPassword").
                        SetParameterValue(mailsOfSenderOfPassword).
                        SetParameterDescription("Password of sender of e-mails").
                        IsPassword(true).
                        SetIsExample("no");
                    resultSaving += parameterSQLite.SaveParameter(parameterOfConfiguration) + "\n";


                    parameterOfConfiguration = new ParameterOfConfigurationBuilder().
                        SetParameterName("MySQLServer").
                        SetParameterValue(mysqlServer).
                        SetParameterDescription("URI MySQL серверa (www)").
                        IsPassword(false).
                        SetIsExample("no");
                    resultSaving += parameterSQLite.SaveParameter(parameterOfConfiguration) + "\n";
                    parameterOfConfiguration = new ParameterOfConfigurationBuilder().
                        SetParameterName("MySQLUser").
                        SetParameterValue(mysqlServerUserName).
                        SetParameterDescription("MySQL User login").
                        IsPassword(false).
                        SetIsExample("no");
                    resultSaving += parameterSQLite.SaveParameter(parameterOfConfiguration) + "\n";
                    parameterOfConfiguration = new ParameterOfConfigurationBuilder().
                        SetParameterName("MySQLUserPassword").
                        SetParameterValue(mysqlServerUserPassword).
                        SetParameterDescription("Password of MySQL User").
                        IsPassword(true).
                        SetIsExample("no");
                    resultSaving += parameterSQLite.SaveParameter(parameterOfConfiguration) + "\n";

                    MessageBox.Show(resultSaving);

                    DisposeTemporaryControls();
                    _ControlVisible(panelView, true);

                    parameterOfConfiguration = null;
                    parameterSQLite = null;
                    resultSaving = null;
                }

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
            }
            else
            {
                GetInfoSetup();
            }
            sqlServerConnectionString = "Data Source=" + sServer1 + "\\SQLEXPRESS;Initial Catalog=intellect;Persist Security Info=True;User ID=" + sServer1UserName + ";Password=" + sServer1UserPassword + "; Connect Timeout=30";
            mysqlServerConnectionStringDB1 = @"server=" + mysqlServer + @";User=" + mysqlServerUserName + @";Password=" + mysqlServerUserPassword + @";database=wwwais;convert zero datetime=True;Connect Timeout=60";
            mysqlServerConnectionStringDB2 = @"server=" + mysqlServer + @";User=" + mysqlServerUserName + @";Password=" + mysqlServerUserPassword + @";database=wwwais;convert zero datetime=True;Connect Timeout=60";

            server = user = password = sMailServer = sMailUser = sMailUserPassword = sMySqlServer = sMySqlServerUser = sMySqlServerUserPassword = null;
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
                SelectFioAndNavFromCombobox();
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

        private void textBoxGroupDescription_TextChanged(object sender, EventArgs e)
        {
            if (textBoxGroupDescription.Text.Trim().Length > 0)
            { StatusLabel2.Text = @"Создать или добавить в группу: " + textBoxGroup.Text.Trim().ToString() + "(" + textBoxGroupDescription.Text.Trim() + ")"; }
            else
            { StatusLabel2.Text = @"Создать или добавить в группу: " + textBoxGroup.Text.Trim().ToString(); }
        }

        private void textBoxGroup_TextChanged(object sender, EventArgs e)
        {
            if (textBoxGroup.Text.Trim().Length > 0)
            {
                AddPersonToGroupItem.Enabled = true;
                CreateGroupItem.Enabled = true;
                if (textBoxGroupDescription.Text.Trim().Length > 0)
                {
                    StatusLabel2.Text = @"Создать или добавить в группу: " + textBoxGroup.Text.Trim().ToString() + "(" + textBoxGroupDescription.Text.Trim() + ")";
                }
                else
                {
                    StatusLabel2.Text = @"Создать или добавить в группу: " + textBoxGroup.Text.Trim().ToString();
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

            SetMenuItemsTextAfterClosingDateTimePicker();
        }

        private void dateTimePickerEnd_CloseUp(object sender, EventArgs e)
        {
            SetMenuItemsTextAfterClosingDateTimePicker();
        }

        private void SetMenuItemsTextAfterClosingDateTimePicker()
        {
            string dayStart = _dateTimePickerReturn(dateTimePickerStart).ToYYYYMMDD();
            string dayEnd = _dateTimePickerReturn(dateTimePickerEnd).ToYYYYMMDD();

            _MenuItemSetText(LoadLastInputsOutputsItem, "Загрузить регистрации пропусков за сегодня (" + DateTime.Now.ToYYYYMMDD() + ")");

            if (dayStart != dayEnd)
                _MenuItemSetText(LoadInputsOutputsItem, "Загрузить регистрации пропусков с " + dayStart + " по " + dayEnd);
            else
                _MenuItemSetText(LoadInputsOutputsItem, "Загрузить регистрации пропусков за " + dayStart);
        }


        private void PersonOrGroupItem_Click(object sender, EventArgs e) //PersonOrGroup()
        { PersonOrGroup(); }

        private void PersonOrGroup()
        {
            string menu = _MenuItemReturnText(PersonOrGroupItem);
            switch (menu)
            {
                case (Names.WORK_WITH_A_GROUP):
                    _MenuItemSetText(PersonOrGroupItem, Names.WORK_WITH_A_PERSON);
                    _ControlEnable(comboBoxFio, false);
                    nameOfLastTable = "PersonRegistrationsList";
                    break;
                case (Names.WORK_WITH_A_PERSON):
                    _MenuItemSetText(PersonOrGroupItem, Names.WORK_WITH_A_GROUP);
                    _ControlEnable(comboBoxFio, true);
                    nameOfLastTable = "PeopleGroup";
                    break;
                default:
                    _MenuItemSetText(PersonOrGroupItem, Names.WORK_WITH_A_GROUP);
                    _ControlEnable(comboBoxFio, true);
                    nameOfLastTable = "PeopleGroup";
                    break;
            }
        }
        //--- End. Behaviour Controls ---//




        //---  Start.  DatagridView functions ---//

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e) //dataGridView1CellClick()
        { dataGridView1CellClick(); }

        private void dataGridView1CellClick()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            int IndexCurrentRow = dgvo.CurrentRowIndex(dataGridView1);

            if (0 < dgvo.RowsCount(dataGridView1) && IndexCurrentRow < dgvo.RowsCount(dataGridView1))
            {
                try
                {
                    //   dgvo = new DataGridViewOperations();

                    if (nameOfLastTable == "PeopleGroupDescription")
                    {
                        dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            Names.GROUP,
                            Names.GROUP_DECRIPTION
                            });

                        textBoxGroup.Text = dgvo?.cellValue[0]; //Take the name of selected group
                        textBoxGroupDescription.Text = dgvo?.cellValue[1]; //Take the name of selected group
                        groupBoxPeriod.BackColor = Color.PaleGreen;
                        groupBoxFilterReport.BackColor = SystemColors.Control;
                        StatusLabel2.Text = @"Выбрана группа: " + dgvo?.cellValue[0];
                        if (textBoxFIO.TextLength > 3)
                        {
                            comboBoxFio.SelectedIndex = comboBoxFio.FindString(textBoxFIO.Text);
                        }
                    }
                    else if (
                        nameOfLastTable == "ListFIO" ||
                        nameOfLastTable == "PeopleGroup" ||
                        nameOfLastTable == "PersonRegistrationsList")
                    {
                        dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            Names.GROUP,
                            Names.FIO,
                            Names.CODE,
                            Names.DESIRED_TIME_IN,
                            Names.DESIRED_TIME_OUT
                            });

                        textBoxGroup.Text = dgvo?.cellValue[0];
                        textBoxFIO.Text = dgvo?.cellValue[1];
                        textBoxNav.Text = dgvo?.cellValue[2];

                        StatusLabel2.Text = @"Выбрана группа: " + dgvo?.cellValue[0] +
                            @" |Курсор на: " + ShortFIO(dgvo?.cellValue[1]);

                        groupBoxPeriod.BackColor = Color.PaleGreen;
                        groupBoxTimeStart.BackColor = Color.PaleGreen;
                        groupBoxTimeEnd.BackColor = Color.PaleGreen;
                        groupBoxFilterReport.BackColor = SystemColors.Control;

                        decimal[] timeIn = new decimal[4];
                        decimal[] timeOut = new decimal[4];
                        try
                        {
                            timeIn = ConvertStringTimeHHMMToDecimalArray(dgvo?.cellValue[3]);
                            timeOut = ConvertStringTimeHHMMToDecimalArray(dgvo?.cellValue[4]);
                            _numUpDownSet(numUpDownHourStart, timeIn[0]);
                            _numUpDownSet(numUpDownMinuteStart, timeIn[1]);
                            _numUpDownSet(numUpDownHourEnd, timeOut[0]);
                            _numUpDownSet(numUpDownMinuteEnd, timeOut[1]);
                        }
                        catch { logger.Warn("dataGridView1CellClick: " + timeIn[0]); }

                        if (dgvo?.cellValue[1]?.Length > 3)
                        {
                            try { comboBoxFio.SelectedIndex = comboBoxFio.FindString(dgvo?.cellValue[1]); }
                            catch
                            {
                                logger.Warn("dataGridView1CellClick: " + dgvo?.cellValue[1] + " not found");
                            }
                        }
                    }
                    else if (nameOfLastTable == "LastIputsOutputs")
                    {
                        dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            Names.N_ID_STRING,
                            Names.FIO,
                            Names.CHECKPOINT_ACTION
                            });

                        textBoxFIO.Text = dgvo?.cellValue[1];
                        textBoxNav.Text = "";

                        StatusLabel2.Text = @" |Курсор на: " + ShortFIO(dgvo?.cellValue[1]);
                    }
                }
                catch (Exception err)
                {
                    logger.Warn("dataGridView1CellClick," + nameOfLastTable + ":" + err.ToString());
                }
            }
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e) //SearchMembersSelectedGroup()
        { SearchMembersSelectedGroup(); }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e) //DataGridView1CellEndEdit()
        { DataGridView1CellEndEdit(); }

        private void DataGridView1CellEndEdit()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            string fio = "";
            string nav = "";
            string group = "";
            int currRow = dgvo.RowsCount(dataGridView1);

            if (currRow > -1)
            {
                try
                {
                    DataGridViewOperations dgvo = new DataGridViewOperations();

                    string currColumn = dgvo.ColumnName(dataGridView1, dgvo.CurrentColumnIndex(dataGridView1));
                    string currCellValue = dgvo.CurrentCellValue(dataGridView1);
                    string editedCell = "";

                    if (nameOfLastTable == @"BoldedDates")
                    {
                        dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
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
                        if (dgvo.cellValue[1]?.Length != 6)
                        { navD = "всех"; }
                        else { navD = dgvo.cellValue[1]; }

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
                    else if (nameOfLastTable == @"PeopleGroup" || nameOfLastTable == @"ListFIO")
                    {
                        dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
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

                        fio = dgvo.cellValue[0];
                        textBoxFIO.Text = fio;

                        nav = dgvo.cellValue[1];
                        textBoxNav.Text = nav;

                        group = dgvo.cellValue[2];
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
                                sqlCommand.Parameters.Add("@Department", DbType.String).Value = dgvo.cellValue[5];
                                sqlCommand.Parameters.Add("@PositionInDepartment", DbType.String).Value = dgvo.cellValue[6];
                                sqlCommand.Parameters.Add("@DepartmentId", DbType.String).Value = dgvo.cellValue[9];

                                sqlCommand.Parameters.Add("@ControllingHHMM", DbType.String).Value = dgvo.cellValue[3];
                                sqlCommand.Parameters.Add("@ControllingOUTHHMM", DbType.String).Value = dgvo.cellValue[4];

                                sqlCommand.Parameters.Add("@Shift", DbType.String).Value = dgvo.cellValue[7];
                                sqlCommand.Parameters.Add("@Comment", DbType.String).Value = dgvo.cellValue[8];

                                try { sqlCommand.ExecuteNonQuery(); } catch { }
                            }
                        }
                        SeekAndShowMembersOfGroup(group);
                        nameOfLastTable = "PeopleGroup";
                        StatusLabel2.Text = @"Обновлено время прихода " + ShortFIO(fio) + " в группе: " + group;
                    }
                    else if (nameOfLastTable == @"PeopleGroupDescription")
                    {
                        dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            Names.GROUP,
                            Names.GROUP_DECRIPTION });

                        textBoxGroup.Text = dgvo.cellValue[0]; //Take the name of selected group
                        textBoxGroupDescription.Text = dgvo.cellValue[1]; //Take the name of selected group
                        groupBoxPeriod.BackColor = Color.PaleGreen;
                        groupBoxFilterReport.BackColor = SystemColors.Control;
                        StatusLabel2.Text = @"Выбрана группа: " + dgvo.cellValue[0];
                    }
                    else if (nameOfLastTable == @"Mailing")
                    {
                        dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            @"Получатель", @"Отчет по группам", @"Наименование", @"Описание",
                            @"Период", @"Статус", @"Тип отчета", @"День отправки отчета" });

                        switch (currColumn)
                        {
                            case "День отправки отчета":
                                editedCell = ReturnStrongNameDayOfSendingReports(dgvo.cellValue[7]);
                                ExecuteSqlAsync("UPDATE 'Mailing' SET DayReport='" + editedCell + "' WHERE RecipientEmail='" + dgvo.cellValue[0]
                                  + "' AND NameReport='" + dgvo.cellValue[2] + "' AND GroupsReport ='" + dgvo.cellValue[1]
                                  + "' AND Period='" + dgvo.cellValue[4] + "' AND TypeReport ='" + dgvo.cellValue[6]
                                  + "' AND Status ='" + dgvo.cellValue[5] + "' AND Description ='" + dgvo.cellValue[3] + "';")
                                  .GetAwaiter().GetResult();
                                break;

                            case "Тип отчета":
                                if (currCellValue == "Полный") { editedCell = "Полный"; }
                                else { editedCell = "Упрощенный"; }

                                ExecuteSqlAsync("UPDATE 'Mailing' SET TypeReport='" + editedCell + "' WHERE RecipientEmail='" + dgvo.cellValue[0]
                                  + "' AND NameReport='" + dgvo.cellValue[2] + "' AND GroupsReport ='" + dgvo.cellValue[1]
                                  + "' AND Period='" + dgvo.cellValue[4] + "' AND DayReport='" + dgvo.cellValue[7]
                                  + "' AND Status ='" + dgvo.cellValue[5] + "' AND Description ='" + dgvo.cellValue[3] + "';")
                                  .GetAwaiter().GetResult();
                                break;

                            case "Статус":
                                if (currCellValue == "Активная") { editedCell = "Активная"; }
                                else { editedCell = "Неактивная"; }

                                ExecuteSqlAsync("UPDATE 'Mailing' SET Status='" + editedCell + "' WHERE RecipientEmail='" + dgvo.cellValue[0]
                                  + "' AND NameReport='" + dgvo.cellValue[2] + "' AND GroupsReport ='" + dgvo.cellValue[1]
                                  + "' AND Period='" + dgvo.cellValue[4] + "' AND DayReport='" + dgvo.cellValue[7]
                                  + "' AND TypeReport ='" + dgvo.cellValue[6] + "' AND Description ='" + dgvo.cellValue[3] + "';")
                                  .GetAwaiter().GetResult();
                                break;

                            case "Период":
                                if (currCellValue == "Текущий месяц") { editedCell = "Текущий месяц"; }
                                else { editedCell = "Предыдущий месяц"; }

                                ExecuteSqlAsync("UPDATE 'Mailing' SET Period='" + editedCell + "' WHERE RecipientEmail='" + dgvo.cellValue[0]
                                   + "' AND NameReport='" + dgvo.cellValue[2] + "' AND GroupsReport ='" + dgvo.cellValue[1]
                                   + "' AND TypeReport ='" + dgvo.cellValue[6] + "' AND DayReport='" + dgvo.cellValue[7]
                                   + "' AND Status ='" + dgvo.cellValue[5] + "' AND Description ='" + dgvo.cellValue[3] + "';")
                                   .GetAwaiter().GetResult();
                                break;

                            case "Описание":
                                editedCell = currCellValue;

                                ExecuteSqlAsync("UPDATE 'Mailing' SET Description='" + editedCell + "' WHERE RecipientEmail='" + dgvo.cellValue[0]
                                  + "' AND NameReport='" + dgvo.cellValue[2] + "' AND GroupsReport ='" + dgvo.cellValue[1]
                                  + "' AND TypeReport ='" + dgvo.cellValue[6] + "' AND DayReport='" + dgvo.cellValue[7]
                                  + "' AND Status ='" + dgvo.cellValue[5] + "' AND Period='" + dgvo.cellValue[4] + "';")
                                  .GetAwaiter().GetResult();
                                break;

                            case "Отчет по группам":
                                editedCell = currCellValue;

                                ExecuteSqlAsync("UPDATE 'Mailing' SET GroupsReport ='" + editedCell + "' WHERE RecipientEmail='" + dgvo.cellValue[0]
                                  + "' AND NameReport='" + dgvo.cellValue[2] + "' AND Description ='" + dgvo.cellValue[3]
                                  + "' AND Status ='" + dgvo.cellValue[5] + "' AND Period='" + dgvo.cellValue[4]
                                  + "' AND TypeReport ='" + dgvo.cellValue[6] + "' AND DayReport='" + dgvo.cellValue[7] + "';")
                                  .GetAwaiter().GetResult();
                                break;

                            case "Получатель":
                                if (currCellValue.Contains('@') && currCellValue.Contains('.'))
                                {
                                    editedCell = currCellValue;

                                    ExecuteSqlAsync("UPDATE 'Mailing' SET RecipientEmail ='" + editedCell + "' WHERE TypeReport ='" + dgvo.cellValue[6]
                                      + "' AND NameReport='" + dgvo.cellValue[2] + "' AND GroupsReport ='" + dgvo.cellValue[1]
                                      + "' AND DayReport='" + dgvo.cellValue[7] + "' AND Period='" + dgvo.cellValue[4]
                                      + "' AND Status ='" + dgvo.cellValue[5] + "' AND Description ='" + dgvo.cellValue[3] + "';")
                                      .GetAwaiter().GetResult();
                                }
                                break;

                            case "Наименование":
                                editedCell = currCellValue;

                                ExecuteSqlAsync("UPDATE 'Mailing' SET NameReport ='" + editedCell + "' WHERE RecipientEmail='" + dgvo.cellValue[0]
                                  + "' AND Description='" + dgvo.cellValue[3] + "' AND GroupsReport ='" + dgvo.cellValue[1]
                                  + "' AND DayReport='" + dgvo.cellValue[7] + "' AND TypeReport ='" + dgvo.cellValue[6]
                                  + "' AND Period ='" + dgvo.cellValue[4] + "' AND Status ='" + dgvo.cellValue[5] + "';")
                                  .GetAwaiter().GetResult();
                                break;

                            default:
                                break;
                        }

                        ShowDataTableDbQuery(dbApplication, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
                        "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
                        "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
                        " ORDER BY RecipientEmail asc, DateCreated desc; ");
                    }
                    else if (nameOfLastTable == @"MailingException")
                    {
                        dgvo.FindValuesInCurrentRow(dataGridView1, new string[] { @"Получатель", @"Описание" });

                        switch (currColumn)
                        {
                            case "Получатель":
                                ExecuteSqlAsync("UPDATE 'MailingException' SET RecipientEmail='" + currCellValue +
                                    "' WHERE Description='" + dgvo.cellValue[1] + "';").GetAwaiter().GetResult();
                                break;

                            case "Описание":
                                ExecuteSqlAsync("UPDATE 'MailingException' SET Description='" + currCellValue +
                                    "' WHERE RecipientEmail='" + dgvo.cellValue[0] + "';").GetAwaiter().GetResult();
                                break;
                            default:
                                break;
                        }

                        ShowDataTableDbQuery(dbApplication, "MailingException", "SELECT RecipientEmail AS 'Получатель', " +
                        "NameReport AS 'Наименование', Description AS 'Описание', DateCreated AS 'Дата создания/модификации', " +
                        " DayReport AS 'День отправки отчета'", " ORDER BY RecipientEmail asc, DateCreated desc; ");
                    }
                    else if (nameOfLastTable == @"SelectedCityToLoadFromWeb")
                    {
                        dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            Names.PLACE_EMPLOYEE,
                            @"Дата создания" });

                        switch (currColumn)
                        {
                            case "Местонахождение сотрудника":
                                ExecuteSqlAsync("UPDATE 'SelectedCityToLoadFromWeb' SET City='" + dgvo.cellValue[0] +
                                                    "' WHERE DateCreated='" + dgvo.cellValue[1] + "';").GetAwaiter().GetResult();
                                break;
                            default:
                                break;
                        }

                        ShowDataTableDbQuery(dbApplication, "SelectedCityToLoadFromWeb", "SELECT City AS 'Местонахождение сотрудника', DateCreated AS 'Дата создания'",
                        " ORDER BY City asc, DateCreated desc; ");
                    }
                }
                catch { }
            }
        }

        private async Task ExecuteSqlAsync(string SqlQuery) //Prepare DB and execute of SQL Query
        {
            string result = string.Empty;
            if (dbApplication.Exists)
            {
                using (SqLiteDbWriter dbWriter = new SqLiteDbWriter(sqLiteLocalConnectionString, dbApplication))
                {
                    dbWriter.ExecuteQuery(SqlQuery);
                    result += dbWriter.Status;
                }
            }
            logger.Trace("ExecuteSqlAsync: query: " + SqlQuery + "\nresult - " + result);
        }


        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (dgvo.ColumnsCount(dataGridView1) > 0 && dgvo.RowsCount(dataGridView1) > 0)
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
        private void dataGridView1_MouseRightClick(object sender, MouseEventArgs e)
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            int currentMouseOverRow = dataGridView1.HitTest(e.X, e.Y).RowIndex;

            if (e.Button == MouseButtons.Right && currentMouseOverRow > -1)
            {
                string txtboxGroup = _ReturnTextOfControl(textBoxGroup);
                string txtboxGroupDescription = _ReturnTextOfControl(textBoxGroupDescription);

                ContextMenu mRightClick = new ContextMenu();
                DataGridViewOperations dgvo = new DataGridViewOperations();
                string recepient = "";
                if (nameOfLastTable == @"PeopleGroupDescription")
                {
                    dgvo.FindValuesInCurrentRow(dataGridView1, new string[]
                    {
                        Names.GROUP,
                        Names.GROUP_DECRIPTION,
                        Names.RECEPIENTS_OF_REPORTS
                    });

                    if (dgvo.cellValue[2]?.Length > 0)
                    {
                        recepient = dgvo.cellValue[2];
                    }
                    else if (mailSenderAddress?.Length > 0)
                    {
                        recepient = mailSenderAddress;
                    }

                    mRightClick.MenuItems.Add(new MenuItem(
                        text: "&Загрузить регистрации пропусков сотрудников группы: '" +
                                dgvo.cellValue[1] + "' за " + _dateTimePickerStartReturnMonth(),
                        onClick: GetDataOfGroup_Click));
                    mRightClick.MenuItems.Add(new MenuItem(
                        text: "Загрузить  входы-выходы сотрудников группы: '" +
                                dgvo.cellValue[1] + "' за " + _dateTimePickerStartReturnMonth() + " и &подготовить отчет",
                        onClick: DoReportByRightClick));
                    mRightClick.MenuItems.Add(new MenuItem(
                        text: "Загрузить регистрации пропусков сотрудников группы: '" +
                                dgvo.cellValue[1] + "' за " + _dateTimePickerStartReturnMonth() + " и &отправить: " + recepient,
                        onClick: DoReportAndEmailByRightClick));

                    mRightClick.MenuItems.Add("-");
                    mRightClick.MenuItems.Add(new MenuItem(
                        text: "&Удалить группу: '" + dgvo.cellValue[0] + "'(" + dgvo.cellValue[1] + ")",
                        onClick: DeleteCurrentRow));
                    mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                }
                else if (nameOfLastTable == @"LastIputsOutputs")
                {
                    dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                        Names.FIO,
                        Names.N_ID_STRING,
                        Names.CHECKPOINT_ACTION
                    });

                    mRightClick.MenuItems.Add(new MenuItem(
                        text: "&Обновить данные о регистрации входов-выходов сотрудников",
                       onClick: LoadLastIputsOutputs_Update_Click));

                    if (nameOfLastTable != "LastIputsOutputs")//Visitor selected=null
                    {
                        checkInputsOutputs = false;
                    }
                    mRightClick.MenuItems.Add(new MenuItem(
                        text: "&Подсветить все входы-выходы '" + dgvo.cellValue[0] + "'",
                       onClick: PaintRowsFioItem_Click));
                    mRightClick.MenuItems.Add(new MenuItem(
                        text: "&Сбросить фильтр",
                       onClick: ResetFilterLoadLastIputsOutput_Click));
                    mRightClick.MenuItems.Add(new MenuItem(
                        text: "Подсветить все состояния '" + dgvo.cellValue[2] + "'",
                       onClick: PaintRowsActionItem_Click));
                    mRightClick.MenuItems.Add("-");
                    mRightClick.MenuItems.Add(new MenuItem(
                        text: "&Загрузить данные регистраций входов-выходов '" +
                                dgvo.cellValue[0] + "' за " + _dateTimePickerStartReturnMonth(),
                        onClick: GetDataOfPerson_Click));
                    mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                }
                else if (nameOfLastTable == @"Mailing")
                {
                    dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                        @"Наименование", @"Описание", @"День отправки отчета", @"Период", @"Тип отчета", @"Получатель"});

                    mRightClick.MenuItems.Add(new MenuItem(
                        text: @"Выполнить активные рассылки по всем у кого: тип отчета - " +
                                dgvo.cellValue[4] + " за " + dgvo.cellValue[3] + " на " + dgvo.cellValue[2],
                        onClick: SendAllReportsInSelectedPeriod));
                    mRightClick.MenuItems.Add("-");
                    mRightClick.MenuItems.Add(new MenuItem(
                        text: @"Выполнить рассылку:   " + dgvo.cellValue[0] + "(" + dgvo.cellValue[1] + ") для " + dgvo.cellValue[5],
                        onClick: DoMainAction));
                    //mRightClick.MenuItems.Add("-");
                    //mRightClick.MenuItems.Add(new MenuItem(text: @"Загрузить регистрации пропусков сотрудников группы: '" + dgvo.cellValue[1] +
                    //     "' за " + _dateTimePickerStartReturnMonth() + " и &отправить: " + dgvo.cellValue[5], 
                    //     onClick: DoReportAndEmailByRightClick));
                    mRightClick.MenuItems.Add("-");
                    mRightClick.MenuItems.Add(new MenuItem(
                        text: @"Создать новую рассылку",
                        onClick: PrepareForMakingFormMailing));
                    mRightClick.MenuItems.Add(new MenuItem(
                        text: @"Клонировать рассылку:   " + dgvo.cellValue[0] + "(" + dgvo.cellValue[1] + ")",
                        onClick: MakeCloneMailing));
                    mRightClick.MenuItems.Add("-");
                    mRightClick.MenuItems.Add(new MenuItem(
                        text: @"Состав рассылки:   " + dgvo.cellValue[0] + "(" + dgvo.cellValue[1] + ")",
                        onClick: MembersGroupItem_Click));
                    mRightClick.MenuItems.Add("-");
                    mRightClick.MenuItems.Add(new MenuItem(
                        text: @"Удалить рассылку:   " + dgvo.cellValue[0] + "(" + dgvo.cellValue[1] + ")",
                        onClick: DeleteCurrentRow));
                    mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                }
                else if (nameOfLastTable == @"MailingException")
                {
                    dgvo.FindValuesInCurrentRow(dataGridView1, new string[] { @"Получатель" });

                    mRightClick.MenuItems.Add(new MenuItem(
                        text: @"Добавить новый адрес 'для исключения из рассылок'",
                        onClick: MakeNewRecepientExcept));
                    mRightClick.MenuItems.Add("-");
                    mRightClick.MenuItems.Add(new MenuItem(
                        text: @"Удалить адрес, ранее внесенный как 'исключеный из рассылок':   " + dgvo.cellValue[0],
                        onClick: DeleteCurrentRow));
                    mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                }
                else if (nameOfLastTable == @"PeopleGroup" || nameOfLastTable == @"ListFIO")
                {
                    dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
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

                    if (string.Compare(dgvo.cellValue[8], txtboxGroup) != 0 && txtboxGroup?.Length > 0) //добавить пункт меню если в текстбоксе группа другая
                    {
                        mRightClick.MenuItems.Add(new MenuItem(
                            text: "&Добавить '" + dgvo.cellValue[0] + "' в группу '" + txtboxGroup + "'",
                            onClick: AddPersonToGroupItem_Click));
                        mRightClick.MenuItems.Add("-");
                    }

                    mRightClick.MenuItems.Add(new MenuItem(
                        text: "Загрузить регистрации пропусков на входе в офис &группы сотрудников '" + dgvo.cellValue[8] + "' за " + _dateTimePickerStartReturnMonth(),
                        onClick: GetDataOfGroup_Click));

                    mRightClick.MenuItems.Add(new MenuItem(
                        text: "Загрузить регистрации пропусков на входе в офис &сотрудника: '" + dgvo.cellValue[0] + "' за " + _dateTimePickerStartReturnMonth(),
                        onClick: GetDataOfPerson_Click));
                    mRightClick.MenuItems.Add("-");
                    mRightClick.MenuItems.Add(new MenuItem(
                        text: "&Удалить '" + dgvo.cellValue[0] + "' из группы '" + txtboxGroup + "'",
                        onClick: DeleteCurrentRow));
                    mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                }
                else if (nameOfLastTable == @"BoldedDates")
                {
                    dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                        Names.DAYOFF_DATE,
                        Names.DAYOFF_USED_BY,
                        Names.DAYOFF_TYPE
                    });

                    string dayType = "";
                    if (txtboxGroup?.Length == 0 || txtboxGroup?.ToLower() == "выходной")
                    { dayType = "Выходной"; }
                    else { dayType = "Рабочий"; }

                    string nav = "";
                    if (textBoxNav?.Text?.Trim()?.Length != 6)
                    { nav = "для всех"; }
                    else { nav = textBoxNav.Text.Trim(); }

                    string navD = "";
                    if (dgvo.cellValue[1]?.Length != 6)
                    { navD = "всех"; }
                    else { navD = dgvo.cellValue[1]; }

                    mRightClick.MenuItems.Add(new MenuItem(
                        text: @"Сохранить для " + nav + @" как '" + dayType + @"' " + monthCalendar.SelectionStart.ToYYYYMMDD(),
                        onClick: AddAnualDateItem_Click));
                    mRightClick.MenuItems.Add("-");
                    mRightClick.MenuItems.Add(new MenuItem(
                        text: @"Удалить из сохранненых '" + dgvo.cellValue[2] + @"'  '" + dgvo.cellValue[0] + @"' для " + navD,
                        onClick: DeleteAnualDateItem_Click));
                    mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                }
                else if (nameOfLastTable == @"SelectedCityToLoadFromWeb")
                {
                    mRightClick.MenuItems.Add(new MenuItem(
                        text: @"Добавить новый город",
                        onClick: AddNewCityToLoadByRightClick));
                    mRightClick.MenuItems.Add("-");
                    mRightClick.MenuItems.Add(new MenuItem(
                        text: @"Удалить выбранный город",
                        onClick: DeleteCityToLoadByRightClick));
                    mRightClick.Show(dataGridView1, new Point(e.X, e.Y));
                }
            }
        }


        private void SelectedToLoadCityItem_Click(object sender, EventArgs e) //SelectedToLoadCity()
        { SelectedToLoadCity(); }

        private void SelectedToLoadCity()
        {
            ShowDataTableDbQuery(dbApplication, "SelectedCityToLoadFromWeb", "SELECT City AS 'Местонахождение сотрудника', DateCreated AS 'Дата создания'",
            " ORDER BY DateCreated desc; ");

            if (dgvo.RowsCount(dataGridView1) < 2)
            {
                AddNewCityToLoad();
                ShowDataTableDbQuery(dbApplication, "SelectedCityToLoadFromWeb", "SELECT City AS 'Местонахождение сотрудника', DateCreated AS 'Дата создания'",
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
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            if (dbApplication.Exists)
            {
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
        }

        private void DeleteCityToLoadByRightClick(object sender, EventArgs e)
        {
            DeleteCityToLoad();
            SelectedToLoadCity();
        }

        private void DeleteCityToLoad()
        {
            DataGridViewOperations dgvo = new DataGridViewOperations();
            dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                Names.PLACE_EMPLOYEE,
                @"Дата создания" });

            DeleteDataTableQueryParameters(dbApplication, "SelectedCityToLoadFromWeb",
                            "City", dgvo.cellValue[0]).GetAwaiter().GetResult();
        }


        private async void DoReportAndEmailByRightClick(object sender, EventArgs e)
        {
            await Task.Run(() => DoReportAndEmailByRightClick());
        }

        private void DoReportAndEmailByRightClick()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            DataGridViewOperations dgvo = new DataGridViewOperations();
            dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                Names.GROUP,
                Names.GROUP_DECRIPTION,
                Names.RECEPIENTS_OF_REPORTS
            });
            resultOfSendingReports = new List<Mailing>();
            logger.Trace("DoReportAndEmailByRightClick");

            _toolStripStatusLabelSetText(StatusLabel2, "Готовлю отчет по группе" + dgvo.cellValue[0]);

            if (dgvo.cellValue[2]?.Length > 0)
            {
                MailingAction(
                    "sendEmail", dgvo.cellValue[2].Trim(), mailSenderAddress,
                    dgvo.cellValue[0], dgvo.cellValue[0], dgvo.cellValue[1], SelectedDatetimePickersPeriodMonth(), "Активная", "Упрощенный", DateTime.Now.ToYYYYMMDDHHMM());
            }
            else if (mailSenderAddress?.Length > 0)
            {
                MailingAction("sendEmail", mailSenderAddress, mailSenderAddress,
             dgvo.cellValue[0], dgvo.cellValue[0], dgvo.cellValue[1], SelectedDatetimePickersPeriodMonth(), "Активная", "Упрощенный", DateTime.Now.ToYYYYMMDDHHMM());
            }
            else
            {
                _toolStripStatusLabelSetText(
                    StatusLabel2,
                    "Попытка отправить отчет " + dgvo.cellValue[0] + " не существующему получателю",
                    true,
                    "DoReportAndEmailByRightClick, the report was attempted to send to non existent user: " + dgvo.cellValue[0]);
            }

            _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);

            if (resultOfSendingReports.Count > 0)
            { SendAdminReport(); }

            _ProgressBar1Stop();
            nameOfLastTable = "PeopleGroupDescription";
        }

        private void DoReportByRightClick(object sender, EventArgs e)
        { DoReportByRightClick(); }

        private void DoReportByRightClick()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            DataGridViewOperations dgvo = new DataGridViewOperations();
            dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                Names.GROUP,
                Names.GROUP_DECRIPTION
            });

            _toolStripStatusLabelSetText(StatusLabel2, "Готовлю отчет по группе" + dgvo.cellValue[0]);
            logger.Trace("DoReportByRightClick: " + dgvo.cellValue[0]);

            resultOfSendingReports = new List<Mailing>();

            GetRegistrationAndSendReport(dgvo.cellValue[0], dgvo.cellValue[0], dgvo.cellValue[1], SelectedDatetimePickersPeriodMonth(), "Активная", "Полный", DateTime.Now.ToYYYYMMDDHHMM(), false, "", "");

            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("explorer.exe", " /select, " + filePathExcelReport + @".xlsx")); // //System.Reflection.Assembly.GetExecutingAssembly().Location)

            _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);

            if (resultOfSendingReports.Count > 0)
            { SendAdminReport(); }

            _ProgressBar1Stop();
            nameOfLastTable = "PeopleGroupDescription";
        }

        private void MakeCloneMailing(object sender, EventArgs e) //MakeCloneMailing()
        { MakeCloneMailing(); }

        private void MakeCloneMailing()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            DataGridViewOperations dgvo = new DataGridViewOperations();
            dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                @"Получатель", @"Отчет по группам", @"Наименование", @"Описание", @"Период"});

            SaveMailing(
               dgvo.cellValue[0], mailSenderAddress, dgvo.cellValue[1], dgvo.cellValue[2] + "_1",
               dgvo.cellValue[3] + "_1", dgvo.cellValue[4], "Неактивная", DateTime.Now.ToYYYYMMDDHHMM(), "", "Копия", DEFAULT_DAY_OF_SENDING_REPORT);

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
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            if (dbApplication.Exists)
            {
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
        }

        private void MailingsExceptItem_Click(object sender, EventArgs e)
        {
            _ControlEnable(comboBoxFio, false);
            dataGridView1.Select();

            ShowDataTableDbQuery(dbApplication, "MailingException", "SELECT RecipientEmail AS 'Получатель', " +
            "NameReport AS 'Наименование', Description AS 'Описание', DateCreated AS 'Дата создания/модификации', " +
            " DayReport AS 'День отправки отчета'", " ORDER BY RecipientEmail asc, DateCreated desc; ");
        }

        private void MailingsShowItem_Click(object sender, EventArgs e)
        {
            _ControlEnable(comboBoxFio, false);
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
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            _ProgressBar1Start();

            switch (nameOfLastTable)
            {
                case "PeopleGroupDescription":
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
                        resultOfSendingReports = new List<Mailing>();

                        DataGridViewOperations dgvo = new DataGridViewOperations();
                        dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            @"Получатель", @"Отчет по группам", @"Наименование", @"Описание",
                            @"Период", @"Статус", @"Тип отчета", @"День отправки отчета" });
                        _toolStripStatusLabelSetText(StatusLabel2, "Готовлю отчет " + dgvo.cellValue[2]);

                        ExecuteSqlAsync("UPDATE 'Mailing' SET SendingLastDate='" + DateTime.Now.ToYYYYMMDDHHMM()
                            + "' WHERE RecipientEmail='" + dgvo.cellValue[0] + "' AND GroupsReport ='" + dgvo.cellValue[1]
                            + "' AND NameReport='" + dgvo.cellValue[2] + "' AND Description ='" + dgvo.cellValue[3]
                            + "' AND Period='" + dgvo.cellValue[4] + "' AND Status='" + dgvo.cellValue[5]
                            + "' AND TypeReport='" + dgvo.cellValue[6] + "' AND DayReport ='" + dgvo.cellValue[7]
                            + "';").GetAwaiter().GetResult();

                        MailingAction("sendEmail", dgvo.cellValue[0], mailSenderAddress,
                            dgvo.cellValue[1], dgvo.cellValue[2], dgvo.cellValue[3], dgvo.cellValue[4],
                            dgvo.cellValue[5], dgvo.cellValue[6], dgvo.cellValue[7]);

                        logger.Info("DoMainAction, sendEmail: Получатель: " +
                            dgvo.cellValue[0] + "|" + dgvo.cellValue[1] + "|Наименование: " + dgvo.cellValue[2] + "|" +
                            dgvo.cellValue[3] + "|Период: " + dgvo.cellValue[4] + "|" + dgvo.cellValue[5] + "|" +
                            dgvo.cellValue[6] + "|" + dgvo.cellValue[7]);

                        ShowDataTableDbQuery(dbApplication, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
                        "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
                        "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
                        " ORDER BY RecipientEmail asc, DateCreated desc; ");

                        _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);

                        if (resultOfSendingReports.Count > 0)
                        { SendAdminReport(); }

                        break;
                    }
                default:
                    break;
            }

            _ProgressBar1Stop();
        }

        private void SendAllReportsInSelectedPeriod(object sender, EventArgs e) //SendAllReportsInSelectedPeriod()
        { SendAllReportsInSelectedPeriod(); }

        private void SendAllReportsInSelectedPeriod()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            _ProgressBar1Start();


            resultOfSendingReports = new List<Mailing>();

            DataGridViewOperations dgvo = new DataGridViewOperations();
            dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            @"Получатель", @"Отчет по группам", @"Наименование", @"Описание",
                            @"Период", @"Статус", @"Тип отчета", @"День отправки отчета" });
            _toolStripStatusLabelSetText(StatusLabel2, "Готовлю все активные рассылки с отчетами " + dgvo.cellValue[6] + " за " + dgvo.cellValue[4] + " на " + dgvo.cellValue[7]);

            currentAction = "sendEmail";
            DoListsFioGroupsMailings().GetAwaiter().GetResult();

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

            logger.Trace("SendAllReportsInSelectedPeriod: активные отчеты " + dgvo.cellValue[6] + " за " + dgvo.cellValue[4] +
                DateTime.Now.FirstDayOfMonth().ToYYYYMMDD() + " - " + DateTime.Now.LastDayOfMonth().ToYYYYMMDD() +
                " на дату - " + dgvo.cellValue[7]
                );

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
                                record["DayReport"]?.ToString()?.Trim()?.ToUpper() == dgvo.cellValue[7]
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

                                logger.Trace("DayReport: " + dayReportInDB + " " + dayReport + " " + dayToSendReport);

                                if (
                                        status == "Активная" &&
                                        typeReport == dgvo.cellValue[6] &&
                                        period == dgvo.cellValue[4] &&
                                        dayReportInDB == dgvo.cellValue[7]
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
                _toolStripStatusLabelSetText(StatusLabel2, "Готовлю отчет " + mailng._nameReport);

                str = "UPDATE 'Mailing' SET SendingLastDate='" + DateTime.Now.ToYYYYMMDDHHMM() +
                    "' WHERE RecipientEmail='" + mailng._recipient +
                    "' AND NameReport='" + mailng._nameReport +
                    "' AND Period='" + mailng._period +
                    "' AND Status='" + mailng._status +
                    "' AND TypeReport='" + mailng._typeReport +
                    "' AND GroupsReport ='" + mailng._groupsReport + "';";
                logger.Trace(str);
                ExecuteSqlAsync(str).GetAwaiter().GetResult();
                GetRegistrationAndSendReport(
                    mailng._groupsReport, mailng._nameReport, mailng._descriptionReport, mailng._period, mailng._status,
                    mailng._typeReport, mailng._dayReport, true, mailng._recipient, mailSenderAddress);

                _ProgressWork1Step();
            }

            logger.Info(method + ": Перечень задач по подготовке и отправке отчетов завершен...");

            ShowDataTableDbQuery(dbApplication, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
            "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
            "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
            " ORDER BY RecipientEmail asc, DateCreated desc; ");

            _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);

            if (resultOfSendingReports.Count > 0)
            { SendAdminReport(); }

            _ProgressBar1Stop();
        }

        private void DeleteCurrentRow(object sender, EventArgs e) //DeleteCurrentRow()
        {
            if (dgvo.RowsCount(dataGridView1) > -1)
            { DeleteCurrentRow(); }
        }

        private void DeleteCurrentRow()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            string group = _ReturnTextOfControl(textBoxGroup);
            DataGridViewOperations dgvo = new DataGridViewOperations();

            switch (nameOfLastTable)
            {
                case "PeopleGroupDescription":
                    {
                        dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            Names.GROUP
                        });

                        DeleteDataTableQueryParameters(dbApplication, "PeopleGroup", "GroupPerson", dgvo.cellValue[0], "", "", "", "").GetAwaiter().GetResult();
                        DeleteDataTableQueryParameters(dbApplication, "PeopleGroupDescription", "GroupPerson", dgvo.cellValue[0], "", "", "", "").GetAwaiter().GetResult();

                        UpdateAmountAndRecepientOfPeopleGroupDescription();
                        ShowDataTableDbQuery(dbApplication, "PeopleGroupDescription", "SELECT GroupPerson AS 'Группа', GroupPersonDescription AS 'Описание группы', AmountStaffInDepartment AS 'Колличество сотрудников в группе' ", " group by GroupPerson ORDER BY GroupPerson asc; ");
                        _toolStripStatusLabelSetText(StatusLabel2, "Удалена группа: " + dgvo.cellValue[0] + "| Всего групп: " + dgvo.RowsCount(dataGridView1));
                        MembersGroupItem.Enabled = true;
                        break;
                    }
                case "PeopleGroup" when group.Length > 0:
                    {
                        int indexCurrentRow = dgvo.RowsCount(dataGridView1);

                        dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            Names.CODE,
                            Names.GROUP });
                        DeleteDataTableQueryParameters(dbApplication, "PeopleGroup", "GroupPerson", dgvo.cellValue[1], "NAV", dgvo.cellValue[0], "", "").GetAwaiter().GetResult();

                        if (indexCurrentRow > 2)
                        { SeekAndShowMembersOfGroup(group); }
                        else
                        {
                            DeleteDataTableQueryParameters(dbApplication, "PeopleGroupDescription", "GroupPerson", dgvo.cellValue[1], "", "", "", "").GetAwaiter().GetResult();

                            UpdateAmountAndRecepientOfPeopleGroupDescription();
                            ShowDataTableDbQuery(dbApplication, "PeopleGroupDescription", "SELECT GroupPerson AS 'Группа', GroupPersonDescription AS 'Описание группы', AmountStaffInDepartment AS 'Колличество сотрудников в группе' ", " group by GroupPerson ORDER BY GroupPerson asc; ");
                        }

                        textBoxGroup.BackColor = Color.White;
                        MembersGroupItem.Enabled = true;
                        break;
                    }
                case "Mailing":
                    {
                        dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            @"Получатель", @"Наименование", @"Дата создания/модификации",
                            @"Отчет по группам", @"Период", @"Тип отчета" });
                        DeleteDataTableQueryParameters(dbApplication, "Mailing",
                            "RecipientEmail", dgvo.cellValue[0],
                            "NameReport", dgvo.cellValue[1],
                            "DateCreated", dgvo.cellValue[2],
                            "GroupsReport", dgvo.cellValue[3],
                            "TypeReport", dgvo.cellValue[5],
                            "Period", dgvo.cellValue[4]).GetAwaiter().GetResult();

                        ShowDataTableDbQuery(dbApplication, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
                        "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
                        "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
                        " ORDER BY RecipientEmail asc, DateCreated desc; ");
                        _toolStripStatusLabelSetText(StatusLabel2, "Удалена рассылка отчета " + dgvo.cellValue[1] + "| Всего рассылок: " + dgvo.RowsCount(dataGridView1));
                        break;
                    }
                case "MailingException":
                    {
                        dgvo.FindValuesInCurrentRow(dataGridView1, new string[] {
                            @"Получатель"});
                        DeleteDataTableQueryParameters(dbApplication, "MailingException",
                            "RecipientEmail", dgvo.cellValue[0]).GetAwaiter().GetResult();

                        ShowDataTableDbQuery(dbApplication, "MailingException", "SELECT RecipientEmail AS 'Получатель', " +
                        "NameReport AS 'Наименование', Description AS 'Описание', DateCreated AS 'Дата создания/модификации', " +
                        "DayReport AS 'День отправки отчета'", " ORDER BY RecipientEmail asc, DateCreated desc; ");
                        _toolStripStatusLabelSetText(StatusLabel2, "Удален из исключений " + dgvo.cellValue[0] + "| Всего исключений: " + dgvo.RowsCount(dataGridView1));
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
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            if (currentModeAppManual)
            {
                _MenuItemSetText(ModeItem, "Выключить режим e-mail рассылок");
                _MenuItemSetTooltip(ModeItem, "Включен автоматический режим. Выполняются Активные рассылки из БД.");
                _MenuItemBackColorChange(ModeItem, Color.DarkOrange);

                _toolStripStatusLabelSetText(StatusLabel2, "Включен режим рассылки отчетов по почте");
                _toolStripStatusLabelBackColor(StatusLabel2, Color.PaleGreen); //Color.DarkOrange

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
                _MenuItemSetText(ModeItem, "Включить режим автоматических e-mail рассылок");
                _MenuItemSetTooltip(ModeItem, "Включен интерактивный режим. Все рассылки остановлены.");
                _MenuItemBackColorChange(ModeItem, SystemColors.Control);

                _toolStripStatusLabelSetText(StatusLabel2, "Интерактивный режим");
                _toolStripStatusLabelBackColor(StatusLabel2, SystemColors.Control);

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
        {
            Task.Run(() => InitScheduleTask(manualMode));
        }

        private void InitScheduleTask(bool manualMode) //ScheduleTask()
        {
            long interval = 60 * 1000; //60 seconds
            if (manualMode)
            {
                _MenuItemSetText(ModeItem, "Выключить режим e-mail рассылок");
                _MenuItemSetTooltip(ModeItem, "Включен автоматический режим. Выполняются Активные рассылки из БД.");
                _MenuItemBackColorChange(ModeItem, Color.DarkOrange);
                _toolStripStatusLabelSetText(StatusLabel2, "Включен режим авторассылки отчетов");
                _toolStripStatusLabelBackColor(StatusLabel2, Color.PaleGreen);

                timer?.Dispose();
                currentAction = "sendEmail";

                timer = new System.Threading.Timer(new System.Threading.TimerCallback(ScheduleTask), null, 0, interval);
            }
            else
            {
                _MenuItemSetText(ModeItem, "Включить режим автоматических e-mail рассылок");
                _MenuItemSetTooltip(ModeItem, "Включен интерактивный режим. Все рассылки остановлены.");
                _MenuItemBackColorChange(ModeItem, SystemColors.Control);

                _toolStripStatusLabelSetText(StatusLabel2, "Интерактивный режим");
                _toolStripStatusLabelBackColor(StatusLabel2, SystemColors.Control);

                timer?.Dispose();
            }
        }

        private void ScheduleTask(object obj) //SelectMailingDoAction()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            lock (synclock)
            {
                DateTime dd = DateTime.Now;
                if (dd.Hour == 4 && dd.Minute == 10 && sent == false) //do something at Hour 2 and 5 minute //dd.Day == 1 && 
                {
                    _toolStripStatusLabelSetText(StatusLabel2, "Ведется работа по подготовке отчетов " + DateTime.Now.ToYYYYMMDDHHMM() + " ...");
                    _toolStripStatusLabelBackColor(StatusLabel2, Color.LightPink);
                    CheckAliveIntellectServer(sServer1, sServer1UserName, sServer1UserPassword).GetAwaiter().GetResult();
                    SelectMailingDoAction();
                    sent = true;
                    _toolStripStatusLabelSetText(StatusLabel2, "Все задачи по подготовке и отправке отчетов завершены.");
                    logger.Info("");
                    logger.Info("---/  " + DateTime.Now.ToYYYYMMDDHHMMSS() + "  /---");
                }
                else
                {
                    sent = false;
                }

                if (dd.Hour == 7 && dd.Minute == 1)
                {
                    _toolStripStatusLabelSetText(StatusLabel2, "Режим почтовых рассылок. " + DateTime.Now.ToYYYYMMDDHHMM());
                    _toolStripStatusLabelBackColor(StatusLabel2, Color.LightCyan);
                    ClearItemsInApplicationFolders(@"*.xlsx");
                }
            }
        }

        private void TestToSendAllMailingsItem_Click(object sender, EventArgs e) //SelectMailingDoAction()
        {
            TestToSendAllMailings().GetAwaiter().GetResult();
        }

        private async Task TestToSendAllMailings()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            await Task.Run(() => CheckAliveIntellectServer(sServer1, sServer1UserName, sServer1UserPassword).GetAwaiter().GetResult());
            await Task.Run(() => UpdateMailingInDB());
            await Task.Run(() => SelectMailingDoAction());
        }

        private string ReturnStrongNameDayOfSendingReports(string inputDate)
        {
            string result = "END_OF_MONTH";

            if (inputDate.Equals("1") || inputDate.Equals("01") || inputDate.Contains("ПЕРВ") || inputDate.Contains("НАЧАЛ") || inputDate.Contains("START"))
            {
                result = "START_OF_MONTH";
            }
            else if (inputDate.Equals("16") || inputDate.Equals("15") || inputDate.Equals("14") || inputDate.Contains("СЕРЕД") || inputDate.Contains("СРЕД") || inputDate.Contains("MIDDLE"))
            {
                result = "MIDDLE_OF_MONTH";
            }

            return result;
        }

        private int ReturnNumberStrongNameDayOfSendingReports(string inputDate, DaysOfSendingMail daysOfSendingMail)
        {
            int result = 0;

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

        private string ReturnStrongNameDayOfSendingReports(int inputDate)
        {
            string result = "";

            if (inputDate == 1 || inputDate == 2 || inputDate == 3)
            {
                result = "START_OF_MONTH";
            }
            else if (inputDate == 14 || inputDate == 15 || inputDate == 16)
            {
                result = "MIDDLE_OF_MONTH";
            }
            else
            {
                result = "END_OF_MONTH";
            }

            return result;
        }

        private void SelectMailingDoAction()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            _ProgressBar1Start();

            currentAction = "sendEmail";

            resultOfSendingReports = new List<Mailing>();
            HashSet<Mailing> mailingList = new HashSet<Mailing>();

            DoListsFioGroupsMailings().GetAwaiter().GetResult();

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
                    _toolStripStatusLabelSetText(StatusLabel2, "Готовлю отчет " + mailng._nameReport);

                    str = "UPDATE 'Mailing' SET SendingLastDate='" + DateTime.Now.ToYYYYMMDDHHMM() +
                        "' WHERE RecipientEmail='" + mailng._recipient +
                        "' AND NameReport='" + mailng._nameReport +
                        "' AND Period='" + mailng._period +
                        "' AND Status='" + mailng._status +
                        "' AND TypeReport='" + mailng._typeReport +
                        "' AND GroupsReport ='" + mailng._groupsReport + "';";
                    logger.Trace(str);
                    ExecuteSqlAsync(str).GetAwaiter().GetResult();
                    GetRegistrationAndSendReport(
                        mailng._groupsReport, mailng._nameReport, mailng._descriptionReport, mailng._period, mailng._status,
                        mailng._typeReport, mailng._dayReport, true, mailng._recipient, mailSenderAddress);

                    _ProgressWork1Step();
                }
                logger.Info("SelectMailingDoAction: Перечень задач по подготовке и отправке отчетов завершен...");
            }
            ShowDataTableDbQuery(dbApplication, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
            "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
            "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
            " ORDER BY RecipientEmail asc, DateCreated desc; ");

            _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);

            if (resultOfSendingReports.Count > 0)
            { SendAdminReport(); }

            _ProgressBar1Stop();
        }


        private void UpdateMailingInDB()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            _ProgressBar1Start();

            string recipient = "";
            string gproupsReport = "";
            string nameReport = "";
            string descriptionReport = "";
            string period = "";
            string status = "";
            string typeReport = "";
            string dayReport = "";
            string str = "";

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
                ExecuteSqlAsync(str).GetAwaiter().GetResult();

                _ProgressWork1Step();
            }

            logger.Info("UpdateMailingInDB: Перечень задач по подготовке и отправке отчетов завершен...");

            ShowDataTableDbQuery(dbApplication, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
            "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
            "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
            " ORDER BY RecipientEmail asc, DateCreated desc; ");

            mailingList = null;

            _ProgressBar1Stop();
        }

        private string GetSafeFilename(string filename, string splitter = "_")
        {
            return string.Join(splitter, filename.Split(System.IO.Path.GetInvalidFileNameChars()));
        }


        private string SelectedDatetimePickersPeriodMonth() //format of result: "1971-01-01 00:00:00|1971-01-31 23:59:59" // 'yyyy-MM-dd HH:mm:SS'
        {
            return _dateTimePickerStartReturn() + "|" + _dateTimePickerEndReturn();
        }

        private void MailingAction(string mainAction, string recipientEmail, string senderEmail, string groupsReport, string nameReport, string description, string period, string status, string typeReport, string dayReport)
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            _toolStripStatusLabelBackColor(StatusLabel2, SystemColors.Control);

            switch (mainAction)
            {
                case "saveEmail":
                    {
                        SaveMailing(mailSenderAddress, senderEmail, groupsReport, nameReport, description, period, status, DateTime.Now.ToYYYYMMDDHHMM(), "", typeReport, dayReport);

                        ShowDataTableDbQuery(dbApplication, "Mailing", "SELECT RecipientEmail AS 'Получатель', GroupsReport AS 'Отчет по группам', NameReport AS 'Наименование', " +
                        "Description AS 'Описание', Period AS 'Период', TypeReport AS 'Тип отчета', DayReport AS 'День отправки отчета', " +
                        "SendingLastDate AS 'Дата последней отправки отчета', Status AS 'Статус', DateCreated AS 'Дата создания/модификации'",
                        " ORDER BY RecipientEmail asc, DateCreated desc; ");
                        break;
                    }
                case "sendEmail":
                    {
                        CheckAliveIntellectServer(sServer1, sServer1UserName, sServer1UserPassword).GetAwaiter().GetResult();

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
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            DataTable dtTempIntermediate = dtPeople.Clone();
            DateTime dtCurrentDate = DateTime.Now;
            EmployeeFull person = new EmployeeFull();
            DateTime selectedDate = DateTime.Now;

            GetNamesOfPassagePoints();  //Get names of the Passage Points

            if (period.ToLower().Contains("текущ"))
            {
                selectedDate = dtCurrentDate;
            }
            else if (period.ToLower().Contains("предыдущ"))
            {
                selectedDate = new DateTime(dtCurrentDate.Year, dtCurrentDate.Month, 1).AddDays(-1);
            }

            reportStartDay = selectedDate.FirstDayOfMonth().ToYYYYMMDD() + " 00:00:00";
            reportLastDay = selectedDate.LastDayOfMonth().ToYYYYMMDD() + " 23:59:59";

            if (!period.ToLower().Contains("предыдущ") && !period.ToLower().Contains("текущ"))
            {
                reportStartDay = SelectedDatetimePickersPeriodMonth().Split('|')[0];
                reportLastDay = SelectedDatetimePickersPeriodMonth().Split('|')[1];
            }

            DateTime start = DateTime.Parse(reportStartDay);
            DateTime end = DateTime.Parse(reportLastDay);

            SeekAnualDays(ref dtTempIntermediate, ref person, false, start.ToIntYYYYMMDD(), end.ToIntYYYYMMDD(), ref myBoldedDates, ref workSelectedDays);
            logger.Trace(reportStartDay + " - " + reportLastDay);

            string nameGroup = "";

            string titleOfbodyMail = "";
            string[] groups = groupsReport.Split('+');

            foreach (string groupName in groups)
            {
                _toolStripStatusLabelBackColor(StatusLabel2, SystemColors.Control);
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
                                code = row[Names.CODE].ToString(),

                                GroupPerson = row[Names.GROUP].ToString(),
                                Department = row[Names.DEPARTMENT].ToString(),
                                PositionInDepartment = row[Names.EMPLOYEE_POSITION].ToString(),
                                DepartmentId = row[Names.DEPARTMENT_ID].ToString(),
                                City = row[Names.PLACE_EMPLOYEE].ToString(),

                                ControlInSeconds = ConvertStringTimeHHMMSSToSeconds(row[Names.DESIRED_TIME_IN].ToString()),
                                ControlOutSeconds = ConvertStringTimeHHMMSSToSeconds(row[Names.DESIRED_TIME_OUT].ToString()),
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
                        ExportDatatableSelectedColumnsToExcel(dtPersonTemp, nameReport, filePathExcelReport).GetAwaiter().GetResult();

                        if (sendReport)
                        {
                            if (reportExcelReady)
                            {
                                titleOfbodyMail = "с " + reportStartDay.Split(' ')[0] + " по " + reportLastDay.Split(' ')[0];
                                _toolStripStatusLabelSetText(StatusLabel2, "Выполняю отправку отчета адресату: " + recipientEmail);

                                foreach (var oneAddress in recipientEmail.Split(','))
                                {
                                    if (oneAddress.Contains('@'))
                                    {
                                        SendStandartReport(oneAddress.Trim(), titleOfbodyMail, description, filePathExcelReport + @".xlsx", appName);
                                        logger.Trace(method + ", SendEmail succesfull: From:" +
                                            mailSenderAddress + "| To: " + oneAddress + "| Subject: " + titleOfbodyMail + "| " +
                                            description + "| attached: " + filePathExcelReport + @".xlsx"
                                            );
                                    }
                                }

                                _toolStripStatusLabelSetText(StatusLabel2, DateTime.Now.ToYYYYMMDDHHMM() + " Отчет '" + nameReport + "'(" + groupName + ") отправлен " + recipientEmail);
                                _toolStripStatusLabelBackColor(StatusLabel2, Color.PaleGreen);
                            }
                            else
                            {
                                _toolStripStatusLabelSetText(
                                    StatusLabel2,
                                    DateTime.Now.ToYYYYMMDDHHMM() + " Ошибка экспорта в файл отчета: " + nameReport + "(" + groupName + ")",
                                    true
                                    );
                            }
                        }
                    }
                    else
                    {
                        _toolStripStatusLabelSetText(
                            StatusLabel2,
                            DateTime.Now.ToYYYYMMDDHHMM() + "Ошибка получения данных для отчета: " + nameReport,
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
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            MailSender mailSender = new MailSender(server);
            mailSender.SetFrom(from);
            mailSender.SetTo(to);

            mailSender.SendEmailAsync(_subject, builder).GetAwaiter().GetResult();
            return mailSender.Status;
        }

        //Compose Standart Report and send e-mail to recepient
        private static void SendStandartReport(string to, string period, string department, string pathToFile, string messageAfterPicture)
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            MailUser _to = new MailUser(to.Split('@')[0], to);
            string subject = "Отчет по посещаемости за период: " + period;

            BodyBuilder builder = MessageBodyStandartReport(period, department, messageAfterPicture);
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

        private static BodyBuilder MessageBodyStandartReport(string period, string department, string messageAfterPicture)
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
            builder.HtmlBody = string.Format(
            @"<p><font size='3' color='black' face='Arial'>Здравствуйте,</p>
            Во вложении «Отчет по учету рабочего времени сотрудников».<p>
            <b>Период: </b>{0}
            <br/><b>Подразделение: </b>'{1}'<br/><p>Уважаемые руководители,</p>
            <p>согласно Приказу ГК АИС «О функционировании процессов кадрового делопроизводства»,<br/><br/>
            <b>Внесите,</b> пожалуйста, <b>до конца текущего месяца</b> по сотрудникам подразделения 
            информацию о командировках, больничных, отпусках, прогулах и т.п. <b>на сайт</b> www.ais .<br/><br/>
            Руководители <b>подразделений</b> ЦОК, <b>не отображающихся на сайте,<br/>вышлите, </b>пожалуйста, 
            <b>Табель</b> учета рабочего времени<br/>
            <b>в отдел компенсаций и льгот до последнего рабочего дня месяца.</b><br/></p>
            <font size='3' color='black' face='Arial'>С, Уважением,<br/>{2}
            </font><br/><br/><font size='2' color='black' face='Arial'><i>
            Данное сообщение и отчет созданы автоматически<br/>программой по учету рабочего времени сотрудников.
            </i></font><br/><font size='1' color='red' face='Arial'><br/>{3}
            </font></p><hr><img alt='ASTA' src='cid:{4}'/><br/><a href='mailto:ryik.yuri@gmail.com'>_</a>"
            , period, department, NAME_OF_SENDER_REPORTS, DateTime.Now.ToYYYYMMDDHHMM(), image.ContentId);

            // We may also want to attach a calendar event for Yuri's party...
            //  builder.Attachments.Add(@"C:\Users\Yuri\Documents\party.ics");
            return builder;
        }

        //Compose Admin Report and send e-mail to Administrator
        private static void SendAdminReport()
        {
            method = System.Reflection.MethodBase.GetCurrentMethod().Name;
            logger.Trace("-= " + method + " =-");

            string period = DateTime.Now.ToYYYYMMDD();
            string subject = "Результат отправки отчетов за " + period;
            MailUser _to = new MailUser(mailJobReportsOfNameOfReceiver.Split('@')[0], mailJobReportsOfNameOfReceiver);

            BodyBuilder builder = MessageBodyAdminReport(period, resultOfSendingReports);
            string statusOfSentEmail = SendEmailAsync(_mailServer, _mailUser, _to, subject, builder);

            logger.Trace("SendAdminReport: Try to send From:" + mailSenderAddress + "| To:" + mailJobReportsOfNameOfReceiver + "| " + period);
            logger.Info("SendAdminReport: " + statusOfSentEmail);
            stimerPrev = "Административный отчет отправлен " + statusOfSentEmail;
        }

        private static BodyBuilder MessageBodyAdminReport(string period, List<Mailing> reportOfResultSending)
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






        /*
        private static string GetFunctionName(Action method)
        {
            return method.Method.Name;
        }*/

        // public static System.Reflection.MethodBase.GetCurrentMethod().Name s;

        private void ListBox_DrawItem(object sender, DrawItemEventArgs e) //Colorize the Listbox
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

        private void ComboBox_DrawItem(object sender, DrawItemEventArgs e) //Colorize the Combobox
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


        //clear all registered Click events on the selected button
        private void RemoveClickEvent(Button b)
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
        private string _ReturnTextOfControl(Control control) //add string into  from other threads
        {
            string tBox = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { tBox = control?.Text?.Trim(); }));
            else
                tBox = control?.Text?.Trim();
            return tBox;
        }

        private void _SetTextOfControl(Control control, string s) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { control.Text = s?.Trim(); }));
            else
                control.Text = s?.Trim();
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
                Invoke(new MethodInvoker(delegate { try { comboBx.SelectedIndex = i; } catch { } }));
            else
                try { comboBx.SelectedIndex = i; } catch { }
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

        private void _dateTimePickerSet(DateTimePicker dateTimePicker, int year, int month, int day) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    dateTimePicker.Value = new DateTime(year, month, day);
                }
                ));
            else
            {
                dateTimePicker.Value = new DateTime(year, month, day);
            }
        }

        private void _dateTimePickerSet(DateTimePicker dateTimePicker, int[] date) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    dateTimePicker.Value = new DateTime(date[0], date[1], date[2]);
                }
                ));
            else
            {
                dateTimePicker.Value = new DateTime(date[0], date[1], date[2]);
            }
        }

        private DateTime _dateTimePickerReturn(DateTimePicker dateTimePicker) //add string into  from other threads
        {
            DateTime result = today;
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

        private string _dateTimePickerReturnString(DateTimePicker dateTimePicker) //add string into  from other threads
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

        private string _dateTimePickerStartReturn() //add string into  from other threads
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

        private string _dateTimePickerEndReturn() //add string into  from other threads
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



        private void _toolStripStatusLabelSetText(ToolStripStatusLabel statusLabel, string s, bool error = false, string errorText = null) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    statusLabel.Text = s;
                    if (error) statusLabel.BackColor = Color.DarkOrange;
                }));
            else
            {
                statusLabel.Text = s;
                if (error) statusLabel.BackColor = Color.DarkOrange;
            }

            stimerPrev = s;

            if (error)
            { logger.Warn(s + "\nОшибка: " + errorText); }
            else
            { logger.Info(s); }
        }

        private string _toolStripStatusLabelReturnText(ToolStripStatusLabel statusLabel)
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
                Invoke(new MethodInvoker(delegate { statusLabel.BackColor = s; }));
            else
                statusLabel.BackColor = s;
        }

        private Color _toolStripStatusLabelReturnBackColor(ToolStripStatusLabel statusLabel) //add string into  from other threads
        {
            Color s = SystemColors.ControlText;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { s = statusLabel.BackColor; }));
            else
                s = statusLabel.BackColor;

            return s;
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
                    panel?.ResumeLayout();
                }));
            else
            {
                panel?.ResumeLayout();
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
                    if (panel?.Parent?.Height > 0)
                        height = panel.Parent.Height;
                }));
            else
            {
                if (panel?.Parent?.Height > 0)
                    height = panel.Parent.Height;
            }
            return height;
        }

        private int _panelHeightReturn(Panel panel) //access from other threads
        {
            int height = 0;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    if (panel?.Height > 0)
                        height = panel.Height;
                }));
            else
            {
                if (panel?.Height > 0)
                    height = panel.Height;
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
                     if (panel?.Width > 0)
                         width = panel.Width;
                 }));
            }
            else
            {
                if (panel?.Width > 0)
                    width = panel.Width;
            }
            return width;
        }

        private int _panelControlsCountReturn(Panel panel) //access from other threads
        {
            int count = 0;
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    if (panel?.Controls?.Count > 0)
                        count = panel.Controls.Count;
                }));
            else
            {
                if (panel?.Controls?.Count > 0)
                    count = panel.Controls.Count;
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

        private void _MenuItemSetText(ToolStripMenuItem menuItem, string text) //access from other threads
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
        }

        private void _MenuItemSetTooltip(ToolStripMenuItem menuItem, string text) //access from other threads
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
        }


        private void _ControlVisible(Control control, bool state) //access from other threads
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

        private void _ControlEnable(Control control, bool state) //access from other threads
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

        private void _ControlDispose(Control control) //access from other threads
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

        private void _ControlChangeBackColor(Control control, Color color) //add string into  from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { control.BackColor = color; }));
            else
                control.BackColor = color; ;
        }

        private void _ControlSetText(Control control, string text)
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
        }

        private void _ControlSetToolTip(Control control, string text)
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

        private void _ProgressWork1Step() //add into progressBar Value 2 from other threads
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

        private void _ProgressBar1Start() //Set progressBar Value into 0 from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    timer1.Enabled = true;
                    ProgressBar1.Value = 0;
                    StatusLabel2.BackColor = SystemColors.Control;
                }));
            else
            {
                timer1.Enabled = true;
                ProgressBar1.Value = 0;
                StatusLabel2.BackColor = SystemColors.Control;
            }
        }

        private void _ProgressBar1Stop() //Set progressBar Value into 100 from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    timer1.Stop();
                    StatusLabel2.Text = stimerPrev;

                    ProgressBar1.Value = 100;
                    StatusLabel2.ForeColor = Color.Black;
                }));
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

        private decimal TryParseStringToDecimal(string str)  //string -> decimal. if error it will return 0
        {
            decimal result = 0;
            if (!decimal.TryParse(str, out result))
            { result = 0; }
            return result;
        }

        private int TryParseStringToInt(string str)  //string -> decimal. if error it will return 0
        {
            int result = 0;
            bool convertOk = int.TryParse(str, out result);
            return result;
        }

        private decimal ConvertDecimalSeparatedTimeToDecimal(decimal decimalHour, decimal decimalMinute)
        {
            decimal result = decimalHour + TryParseStringToDecimal(TimeSpan.FromMinutes((double)decimalMinute).TotalHours.ToString());
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

        private string ConvertSecondsToStringHHMM(int seconds)
        {
            string result;
            int hours = seconds / 3600;
            int minutes = (seconds % 3600) / 60;

            result = string.Format("{0:d2}:{1:d2}", hours, minutes);
            return result;
        }

        private string ConvertSecondsToStringHHMMSS(int seconds)
        {
            string result;
            int hours = seconds / 3600;
            int minutes = (seconds % 3600) / 60;
            int sec = seconds - hours * 3600 - minutes * 60;

            result = string.Format("{0:d2}:{1:d2}:{2:d2}", hours, minutes, sec);
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

        private string ConvertDecimalTimeToStringHHMM(decimal hours, decimal minutes)
        {
            string result = string.Format("{0:d2}:{1:d2}", (int)hours, (int)minutes);
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

        //todo -  remove
        //it already is exists in the  class - DateTimeConvertor
        private string ConvertStringsTimeToStringHHMMSS(string time)
        {
            int h = 0;
            int m = 0;
            int s = 0;

            if (time.Contains(':'))
            {
                int.TryParse(time.Split(':')[0], out h);

                if (time.Split(':').Length > 1)
                {
                    int.TryParse(time.Split(':')[1], out m);

                    if (time.Split(':').Length > 2)
                    {
                        int.TryParse(time.Split(':')[2], out s);
                    }
                }
                return String.Format("{0:d2}:{1:d2}:{2:d2}", h, m, s);
            }
            else
            {
                return time;
            }
        }

        private int ConvertStringsTimeToSeconds(string hour, string minute)
        {
            int h = 0;
            int m = 0;
            bool hourOk = int.TryParse(hour, out h);
            bool minuteOk = int.TryParse(minute, out m);
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

        private int ConvertStringTimeHHMMSSToSeconds(string timeInHHMMSS) //time HH:MM:SS converted to decimal value
        {
            string hours = "0";
            string minutes = "0";
            string seconds = "0";
            int length = timeInHHMMSS.Split(':').Length;

            if (length > 2)
            {
                string[] time = timeInHHMMSS.Split(':');
                hours = time[0];
                minutes = time[1];
                seconds = time[2];
            }
            else if (length == 2)
            {
                string[] time = timeInHHMMSS.Split(':');
                hours = time[0];
                minutes = time[1];
            }
            else if (length == 1)
            {
                hours = timeInHHMMSS;
            }

            return (60 * 60 * Convert.ToInt32(hours) + 60 * Convert.ToInt32(minutes) + Convert.ToInt32(seconds));
        }

        private int[] ConvertStringDateToIntArray(string dateYYYYmmDD) //date "YYYY-MM-DD HH:MM" to  int[] { 1970, 1, 1 }
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

        //---- End. Convertors of data types ----//



        private void StatusLabelAddInfo(object sender, AccountEventArgs e)
        {
            _toolStripStatusLabelSetText(StatusLabel2, e.Message);
        }
        private void StatusLabelSetBackColor(object sender, AccountEventBoolArgs e)
        {
            _toolStripStatusLabelBackColor(StatusLabel2, e.Color);
        }

        private void StatusLabelAddInfo(string message)
        {
            _toolStripStatusLabelSetText(StatusLabel2, message);
        }

        private void LoggerAddInfo(object sender, AccountEventArgs e)
        {
            logger.Info(e.Message);
        }

        private void LoggerAddInfo(string message)
        {
            logger.Info(message);
        }



        private void CreateDBItem_Click(object sender, EventArgs e)
        {
            TryMakeLocalDB();
        }

        private void GetCurrentSchemeItem_Click(object sender, EventArgs e)
        {
            GetSQLiteDbScheme();
        }

        private void GetSQLiteDbScheme()
        {
            StringBuilder sb = new StringBuilder();
            string fpath = SelectFileOpenFileDialog("Выберите файл", "SQL файлы (*.sql)|*.sql|Все files (*.*)|*.*");

            if (fpath == null)
                return;

            fpath = dbApplication.FullName.ToString();

            Cursor = Cursors.WaitCursor;
            System.Threading.Thread worker = new System.Threading.Thread(new System.Threading.ThreadStart(delegate
            {
                try
                {
                    SQLiteConnectionStringBuilder builder = new SQLiteConnectionStringBuilder();
                    builder.DataSource = fpath;
                    builder.PageSize = 4096;
                    builder.UseUTF16Encoding = true;
                    using (SQLiteConnection conn = new SQLiteConnection(builder.ConnectionString))
                    {
                        conn.Open();

                        SQLiteCommand count = new SQLiteCommand(
                            @"SELECT COUNT(*) FROM SQLITE_MASTER", conn);
                        long num = (long)count.ExecuteScalar();

                        int step = 0;
                        SQLiteCommand query = new SQLiteCommand(
                            @"SELECT * FROM SQLITE_MASTER", conn);
                        using (SQLiteDataReader reader = query.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                step++;

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
                                {
                                    sb.Append(sql + ";\r\n\r\n");
                                };
                                this.Invoke(mi);
                            } // while
                        } // using
                    } // using

                    MethodInvoker mi3 = delegate
                    {
                        MessageBox.Show(this,
                            sb.ToString(),
                            "SQLite Scheme: " + fpath,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        logger.Info("SQLite Scheme: " + fpath);
                        logger.Info(sb.ToString());
                    };
                    this.Invoke(mi3);
                }
                catch (Exception ex)
                {
                    MethodInvoker mi1 = delegate
                    {
                        MessageBox.Show(this,
                            ex.Message,
                            "Extraction Failed",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
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




        /// <summary>
        /// ///////////////////////////////////
        //  Update - Upload
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AutoupdatItem_Click(object sender, EventArgs e)
        {
            if (!uploadingUpdate)
            {
                UpdatingParameters parameters = MakeStartParametersOfUpdating();

                AutoUpdater.CheckForUpdateEvent += AutoUpdaterOnAutoCheckForUpdateEvent; //write errors if had no access to the folder
                AutoUpdater.ApplicationExitEvent += AutoUpdater_ApplicationExitEvent;    //https://archive.codeplex.com/?p=autoupdaterdotnet
                AutoUpdater.RunUpdateAsAdmin = false;
                AutoUpdater.Mandatory = true;
                AutoUpdater.UpdateMode = Mode.ForcedDownload;

                AutoUpdater.LetUserSelectRemindLater = false;
                AutoUpdater.RemindLaterTimeSpan = RemindLaterFormat.Days;
                AutoUpdater.RemindLaterAt = 2;

                AutoUpdater.Start(parameters.appUpdateURL, System.Reflection.Assembly.GetEntryAssembly());
                AutoUpdater.CheckForUpdateEvent -= AutoUpdaterOnAutoCheckForUpdateEvent; //write errors if had no access to the folder
            }
            else
            { StatusLabelAddInfo(@"Ждите! На сервер загружается новая версия ПО"); }
        }

        UpdatingParameters MakeStartParametersOfUpdating()
        {
            UpdatingParameters parameters = new UpdatingParameters
            {
                localFolderUpdatingURL = localAppFolderPath,
                remoteFolderUpdatingURL = remoteFolderUpdateURL,
                appVersion = appVersionAssembly,
                appFileXml = appNameXML,
                appUpdateMD5 = appFileMD5,
                appFileZip = appFileZip
            };

            MakerOfLinks makerLinks = new MakerOfLinks();
            makerLinks.SetParameters(parameters);
            makerLinks.Make();

            return makerLinks.GetParameters();
        }

        private async Task AutoUpdate()
        {
            //Check updates frequently
            System.Timers.Timer timer = new System.Timers.Timer
            {
                Interval = 1 * 60 * 60 * 1000,       // 1 * 60 * 1000 // it sets the interval of checking equal at a minute
                SynchronizingObject = this
            };
            timer.Elapsed += delegate
            {
                //https://github.com/ravibpatel/AutoUpdater.NET
                //http://www.cyberforum.ru/csharp-beginners/thread2169711.html

                //Basic Authetication for XML, Update file and Change Log
                // BasicAuthentication basicAuthentication = new BasicAuthentication("myUserName", "myPassword");
                // AutoUpdater.BasicAuthXML = AutoUpdater.BasicAuthDownload = AutoUpdater.BasicAuthChangeLog = basicAuthentication;

                // AutoUpdater.CheckForUpdateEvent += AutoUpdaterOnCheckForUpdateEvent; //check manualy only
                // AutoUpdater.ReportErrors = true; // will show error message, if there is no update available or if it can't get to the XML file from web server.
                // AutoUpdater.CheckForUpdateEvent -= AutoUpdaterOnAutoCheckForUpdateEvent;
                // AutoUpdater.RemindLaterTimeSpan = RemindLaterFormat.Minutes;
                // AutoUpdater.RemindLaterAt = 1;
                // AutoUpdater.ApplicationExitEvent += ApplicationExit;
                if (!uploadingUpdate)
                {
                    UpdatingParameters parameters = MakeStartParametersOfUpdating();
                    logger.Trace(@"Update URL: " + parameters.appUpdateURL);
                    appUpdateURL = parameters.appUpdateURL;

                    AutoUpdater.CheckForUpdateEvent += AutoUpdaterOnAutoCheckForUpdateEvent; //write errors if had no access to the folder
                    AutoUpdater.ApplicationExitEvent += AutoUpdater_ApplicationExitEvent;    //https://archive.codeplex.com/?p=autoupdaterdotnet
                    AutoUpdater.RunUpdateAsAdmin = false;
                    AutoUpdater.Mandatory = true;
                    AutoUpdater.UpdateMode = Mode.ForcedDownload;

                    AutoUpdater.LetUserSelectRemindLater = false;
                    AutoUpdater.RemindLaterTimeSpan = RemindLaterFormat.Days;
                    AutoUpdater.RemindLaterAt = 2;

                    // AutoUpdater.ReportErrors = false;
                    // AutoUpdater.AppCastURL = appUpdateURL;
                    AutoUpdater.DownloadPath = appFolderUpdatePath;

                    AutoUpdater.Start(parameters.appUpdateURL, System.Reflection.Assembly.GetEntryAssembly());
                    //AutoUpdater.Start("ftp://kv-sb-server.corp.ais/Common/ASTA/ASTA.xml", new NetworkCredential("FtpUserName", "FtpPassword")); //download from FTP
                    AutoUpdater.CheckForUpdateEvent -= AutoUpdaterOnAutoCheckForUpdateEvent; //write errors if had no access to the folder
                }
                else
                { logger.Trace(@"Обновление приостановлено. На сервер сейчас загружается новая версия ПО"); }
            };
            timer.Start();
        }

        private void AutoUpdaterOnAutoCheckForUpdateEvent(UpdateInfoEventArgs args)
        {
            if (args != null)
            {
                if (args.IsUpdateAvailable)
                {
                    try
                    {
                        if (AutoUpdater.DownloadUpdate())
                        {
                            _toolStripStatusLabelBackColor(StatusLabel2, Color.DarkOrange);
                            System.Xml.XmlDocument xmldoc = new System.Xml.XmlDocument();
                            System.Xml.XmlNodeList xmlnode;
                            xmldoc.Load(appUpdateURL);
                            xmlnode = xmldoc.GetElementsByTagName("version");
                            string foundNewVersionApp = xmlnode[0].InnerText;
                            logger.Info("-----------------------------------------");
                            logger.Info("");
                            logger.Trace("-= Update =-");
                            logger.Trace("...");
                            _toolStripStatusLabelSetText(
                                StatusLabel2,
                                @" обнаружена новая версия " + appName + " ver." + foundNewVersionApp,
                                true,
                                "Запущено обновление ПО. Старая версия - " + appVersionAssembly + " -> " + foundNewVersionApp
                                );
                            logger.Trace("...");
                            logger.Trace("-= Update =-");
                            logger.Info("");
                            logger.Info("-----------------------------------------");

                            ApplicationExit();
                        }
                    }
                    catch (Exception exception)
                    { logger.Warn(@"Update's check was failed: " + exception.Message + "| " + exception.GetType().ToString()); }
                    // Uncomment the following line if you want to show standard update dialog instead.
                    // AutoUpdater.ShowUpdateForm();
                }
                else
                {
                    _toolStripStatusLabelBackColor(StatusLabel2, Color.PaleGreen);
                    _toolStripStatusLabelSetText(StatusLabel2, @"Новых версий ПО '" + appName + "' не обнаружено");
                }
            }
            else
            {
                logger.Warn(@"Update check failed: There is a problem reaching update server URL.");
            }
        }
        private void AutoUpdater_ApplicationExitEvent()
        {
            Text = @"Closing application...";
            System.Threading.Thread.Sleep(500);
            Application.Exit();
        }
        //Calculate MD5 checksum of local file
        private string CalculateMD5OfFile(string file)
        {
            CalculatingHash calculatedHash = new CalculatingHash(file);
            return calculatedHash.Calculate();
        }

        private UpdatingParameters PrepareUpdating()
        {
            LoggerAddInfo(remoteFolderUpdateURL);

            //Make an archive with the currrent app's version 
            MakeZip(appAllFiles, appFileZip);

            //Make MD5 of ZIP archive
            appFileMD5 = CalculateMD5OfFile(appFileZip);

            UpdatingParameters parameters = new UpdatingParameters
            {
                localFolderUpdatingURL = localAppFolderPath,
                remoteFolderUpdatingURL = remoteFolderUpdateURL,
                appVersion = appVersionAssembly,
                appFileXml = appNameXML,
                appUpdateMD5 = appFileMD5,
                appFileZip = appFileZip
            };

            MakerOfLinks makerLinks = new MakerOfLinks();
            makerLinks.status += StatusLabelAddInfo;

            MakerOfUpdateXmlFile makerXML = new MakerOfUpdateXmlFile();
            makerXML.status += StatusLabelAddInfo;

            UpdatePreparing preparing = new UpdatePreparing(makerLinks, makerXML, parameters);
            preparing.status += StatusLabelAddInfo;

            preparing.Do();

            makerXML.status -= StatusLabelAddInfo;
            makerLinks.status -= StatusLabelAddInfo;
            preparing.status -= StatusLabelAddInfo;

            return preparing.GetParameters();
        }

        private void CalculateHashItem_Click(object sender, EventArgs e) //Selectfiles()
        {
            SelectfilesForCalculatingHash();
            //  TestHash();
        }


        private void SelectfilesForCalculatingHash() //SelectFileOpenFileDialog() CalculateFileHash()
        {
            string result = null;
            string filePath;
            CalculatingHash calculatedHash;

            DialogResult selectTwoFiles = MessageBox.Show("Выбрать 2 файла для сравнения?", "Сравнение файлов",
                MessageBoxButtons.YesNo,
                      MessageBoxIcon.Exclamation,
                      MessageBoxDefaultButton.Button1);

            filePath = SelectFileOpenFileDialog("Выберите первый файл для вычисления хэша");
            calculatedHash = new CalculatingHash(filePath);
            result += calculatedHash.Calculate() + "\n";

            if (selectTwoFiles == DialogResult.Yes)
            {
                filePath = SelectFileOpenFileDialog("Выберите следующий файл для вычисления хэша");
                calculatedHash = new CalculatingHash(filePath);
                result += "\n" + calculatedHash.Calculate() + "\n";
            }
            MessageBox.Show(result, "Результат вычисления хэша");
        }

        private string SelectFileOpenFileDialog(string titleWindowDialog = null, string maskFiles = "Все файлы (*.*)|*.*")
        {
            string filePath = null;
            MethodInvoker mi = delegate
            {
                using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
                {
                    openFileDialog1.FileName = "";

                    if (titleWindowDialog != null)
                    { openFileDialog1.Title = titleWindowDialog; }

                    openFileDialog1.Filter = maskFiles;
                    DialogResult res = openFileDialog1.ShowDialog();
                    if (res == DialogResult.Cancel)
                        return;

                    filePath = openFileDialog1.FileName;
                }
            };

            this.Invoke(mi);
            return filePath;
        }


        //Upload App's files to Server
        private void UploadApplicationItem_Click(object sender, EventArgs e) //Uploading()
        {
            Task.Run(() => Uploading());
        }

        private async void Uploading() //UploadApplicationToShare()
        {
            _toolStripStatusLabelBackColor(StatusLabel2, SystemColors.Control);

            uploadingUpdate = true;

            UpdatingParameters parameters = PrepareUpdating();

            string[] source =  {
                parameters.localFolderUpdatingURL + @"\" + parameters.appFileXml,
                parameters.localFolderUpdatingURL + @"\" + parameters.appFileZip
            };

            string[] target = {
                parameters.appUpdateFolderURI + parameters.appFileXml,
                parameters.appUpdateFolderURI + parameters.appFileZip
            };
            
            using (Uploader uploader = new Uploader(parameters, source, target))
            {
                uploader.status += StatusLabelAddInfo;
                uploader.uploaded += StatusLabelSetBackColor;
                uploader.Upload();
                uploader.status -= StatusLabelAddInfo;
                uploader.uploaded -= StatusLabelSetBackColor;
            }

            foreach (var file in source)
            {
                if (System.IO.File.Exists(file))
                {
                    try { System.IO.File.Delete(file); } catch { }
                }

            }
            parameters = null;
            uploadingUpdate = false;
        }
    }
}
