using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASTA
{
    class AAA
    {
        //todo
        //will make Update local DB structure

        /*
-= -= update scheme =-
1    Start a transaction.
2    Run PRAGMA schema_version to determine the current schema version number. This number will be needed for step 6 below.
3    Activate schema editing using PRAGMA writable_schema=ON.
4    Run an UPDATE statement to change the definition of table X in the sqlite_master table: 
	UPDATE sqlite_master SET sql=... WHERE type='table' AND name='X';
5    Caution: Making a change to the sqlite_master table like this will render the database corrupt and unreadable 
	if the change contains a syntax error. It is suggested that careful testing of 
	the UPDATE statement be done on a separate blank database prior to using it on a database containing important data.
6    If the change to table X also affects other tables or indexes or triggers are views within schema, then 
	run UPDATE statements to modify those other tables indexes and views too. For example, 
	if the name of a column changes, all FOREIGN KEY constraints, triggers, indexes, and views 
	that refer to that column must be modified.
7    Caution: Once again, making changes to the sqlite_master table like 
	this will render the database corrupt and unreadable if the change contains an error. 
	Carefully test this entire procedure on a separate test database 
	prior to using it on a database containing important data and/or 
	make backup copies of important databases prior to running this procedure.
8    Increment the schema version number using PRAGMA schema_version=X 
	where X is one more than the old schema version number found in step 2 above.
9    Disable schema editing using PRAGMA writable_schema=OFF.
10    (Optional) Run PRAGMA integrity_check to verify that the schema changes did not damage the database.
11    Commit the transaction started on step 1 above. 

-= If some future version of SQLite adds new ALTER TABLE capabilities, 
	those capabilities will very likely use one of the two procedures outlined above.


==================
-= rename table =-

ALTER TABLE existing_table
RENAME TO new_table;


=======================
-= Adding a new column  //new column cannot have a UNIQUE or PRIMARY KEY constraint

ALTER TABLE table
ADD COLUMN column_definition;


======================


-=  ALTER TABLE – Other actions =-

PRAGMA foreign_keys=off; 
BEGIN TRANSACTION; 

ALTER TABLE table RENAME TO temp_table; 
CREATE TABLE table
( 
   column_definition,
   ...
);
 
INSERT INTO table (column_list)
  SELECT column_list
  FROM temp_table;
 
DROP TABLE temp_table;
 
COMMIT;
 
PRAGMA foreign_keys=on;
 

         
-= DROP COLUMN example =-

BEGIN TRANSACTION;
 
ALTER TABLE equipment RENAME TO temp_equipment;
 
CREATE TABLE equipment (
 name text NOT NULL,
 model text NOT NULL,
 serial integer NOT NULL UNIQUE
);
 
INSERT INTO equipment 
SELECT
 name, model, serial
FROM
 temp_equipment;
 
DROP TABLE temp_equipment;
 
COMMIT;

         */




    /*

     using in a method:
     {
     DataTable dt = DbSqlJob.GetTable();

     var category_Data = dt.AsEnumerable()
         .GroupBy(row => row.Field<string>("Категория"))
         .Select(cat => new {
             Category_Name = cat.Key,
             Category_Count = cat.Count()
         });
     }

     public static class DbSqlJob
 {
     private static string DB_PATH = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "testbase.mdf");
     private static string CONNECT_STRING = string.Format(
                           @"Data Source=.\SQLEXPRESS;AttachDbFilename={0};" +
                           "Integrated Security=True;Connect Timeout=30;User Instance=True", DB_PATH);
     public static DataTable GetTable()
     {
         DataTable dt = new DataTable();
         using (System.Data.SqlClient.SqlDataAdapter adapter = new System.Data.SqlClient.SqlDataAdapter("SELECT * FROM Table_1", CONNECT_STRING))
         {
             adapter.Fill(dt);
         }
         return dt;
     }
 }*/


    //todo
    /*
     const string connStr = "server=localhost;user=root; database=test;password=;";
using (MySqlConnection conn = new MySqlConnection(connStr))
{
string sql = "SELECT Text FROM Kody WHERE Kod=@Kod";
MySqlCommand comand = new MySqlCommand(sql, conn);
command.Parameters.AddWithValue("@Kod", textBox1.Text);
conn.Open();
string name = comand.ExecuteScalar().ToString();
label1.Text = name;
}
*/
    /*
        interface IDBParameter
        {
            string DBName { get; set; }
            string TableName { get; set; }
            string DBConnectionString { get; set; }
            string DBQuery { get; set; }
            void Execute();
        }
        */

    //todo
    /*
     const string connStr = "server=localhost;user=root; database=test;password=;";
using (MySqlConnection conn = new MySqlConnection(connStr))
{
string sql = "SELECT Text FROM Kody WHERE Kod=@Kod";
MySqlCommand comand = new MySqlCommand(sql, conn);
command.Parameters.AddWithValue("@Kod", textBox1.Text);
conn.Open();
string name = comand.ExecuteScalar().ToString();
label1.Text = name;
}
*/


    /// <summary>
    /// ///////////////////////Export to Excel///////////////////
    /// </summary>

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
    /*

    // подключить библиотеку Microsoft.Office.Interop.Excel
    // создаем псевдоним для работы с Excel:
    using Excel = Microsoft.Office.Interop.Excel;

        //Объявляем приложение
        Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();
        //Отобразить Excel
        ex.Visible = true;
        //Количество листов в рабочей книге
        ex.SheetsInNewWorkbook = 2;
        //Добавить рабочую книгу
        Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);
        //Отключить отображение окон с сообщениями
        ex.DisplayAlerts = false;                                       
        //Получаем первый лист документа (счет начинается с 1)
        Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);
        //Название листа (вкладки снизу)
        sheet.Name = "Отчет за 13.12.2017";
        //Пример заполнения ячеек
        for (int i = 1; i <= 9; i++)
        {
          for (int j = 1; j < 9; j++)
          sheet.Cells[i, j] = String.Format("Boom {0} {1}", i, j);
        }
        //Захватываем диапазон ячеек
        Excel.Range range1 = sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[9, 9]);
        //Шрифт для диапазона
        range1.Cells.Font.Name = "Tahoma";
        //Размер шрифта для диапазона
        range1.Cells.Font.Size = 10;
        //Захватываем другой диапазон ячеек
        Excel.Range range2 = sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[9, 2]);
        range2.Cells.Font.Name = "Times New Roman";
        //Задаем цвет этого диапазона. Необходимо подключить System.Drawing
        range2.Cells.Font.Color = ColorTranslator.ToOle(Color.Green);
        //Фоновый цвет
        range2.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(0xFF, 0xFF, 0xCC));

    
    //Расстановка рамок.
   // Расставляем рамки со всех сторон:
        range2.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
        range2.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
        range2.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
        range2.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
        range2.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;

    //Цвет рамки можно установить так:
    range2.Borders.Color = ColorTranslator.ToOle(Color.Red);

    //Выравнивания в диапазоне задаются так:
        rangeDate.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
        rangeDate.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;


    //Определим задачу: получить сумму диапазона ячеек A4:A10.
    //Для начала снова получим диапазон ячеек:
    Excel.Range formulaRange = sheet.get_Range(sheet.Cells[4, 1], sheet.Cells[9, 1]);

    //Далее получим диапазон вида A4:A10 по адресу ячейки ( [4,1]; [9;1] ) описанному выше:
    string addr = formulaRange.get_Address(1, 1, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

    //Теперь в переменной addr у нас хранится строковое значение диапазона ( [4,1]; [9;1] ), то есть A4:A10.
    //Вычисляем формулу:
        //Одна ячейка как диапазон
        Excel.Range r = sheet.Cells[10, 1] as Excel.Range;
        //Оформления
        r.Font.Name = "Times New Roman";
        r.Font.Bold = true;
        r.Font.Color = ColorTranslator.ToOle(Color.Blue);
        //Задаем формулу суммы
        r.Formula = String.Format("=СУММ({0}", addr);

    // Выделение ячейки или диапазона ячеек
    //выделить ячейку или диапазон, как если бы мы выделили их мышкой:
        sheet.get_Range("J3", "J8").Activate();
        //или
        sheet.get_Range("J3", "J8").Select();
        //Можно вписать одну и ту же ячейку, тогда будет выделена одна ячейка.
        sheet.get_Range("J3", "J3").Activate();
        sheet.get_Range("J3", "J3").Select();

        //Чтобы настроить авто ширину и высоту для диапазона, используем такие команды:
        range.EntireColumn.AutoFit(); 
        range.EntireRow.AutoFit();

    //Чтобы получить значение из ячейки, используем такой код:
        //Получение одной ячейки как ранга
        Excel.Range forYach = sheet.Cells[ob + 1, 1] as Excel.Range;
        //Получаем значение из ячейки и преобразуем в строку
        string yach = forYach.Value2.ToString();

    // Чтобы добавить лист и дать ему заголовок, используем следующее:
        var sh = workBook.Sheets;
        Excel.Worksheet sheetPivot = (Excel.Worksheet)sh.Add(Type.Missing, sh[1], Type.Missing, Type.Missing);
        sheetPivot.Name = "Сводная таблица";

   // Добавление разрыва страницы
        //Ячейка, с которой будет разрыв
        Excel.Range razr = sheet.Cells[n, m] as Excel.Range;
        //Добавить горизонтальный разрыв (sheet - текущий лист)
        sheet.HPageBreaks.Add(razr); 
        //VPageBreaks - Добавить вертикальный разрыв

    //Сохраняем документ
        ex.Application.ActiveWorkbook.SaveAs("doc.xlsx", Type.Missing,
          Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
          Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

    //Как открыть существующий документ Excel
        ex.Workbooks.Open(@"C:\Users\Myuser\Documents\Excel.xlsx",
          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
          Type.Missing, Type.Missing);

    //Складываем значения предыдущих 12 ячеек слева
        rang.Formula = "=СУММ(RC[-12]:RC[-1])";

    //Так же во время работы может возникнуть ошибка: метод завершен неверно. Это может означать, что не выбран лист, с которым идет работа.

    //Чтобы выбрать лист, выполните,  где sheetData это нужный лист.
    sheetData.Select(Type.Missing); 

         */



        /*
        class ActiveDirectoryDataBak
        {
            static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
            UserADAuthorization UserADAuthorization;

            public ObservableCollection<ADUser> ADUsersCollection = new ObservableCollection<ADUser>();

            public ActiveDirectoryDataBak(string _user, string _domain, string _password, string _domainPath)
            {
                UserADAuthorization = new UserADAuthorization()
                {
                    Name = _user,
                    Password = _password,
                    Domain = _domain,
                    DomainPath = _domainPath
                };

                // isValid = ValidateCredentials(UserADAuthorization);
                //  logger.Trace("!Test only!  "+"Доступ к домену '" + UserADAuthorization.Domain + "' предоставлен: " + isValid);
                ADUsersCollection = new ObservableCollection<ADUser>();
            }

            public ObservableCollection<ADUser> GetADUsers()
            {
                logger.Trace(UserADAuthorization.DomainPath);
                int userCount = 0;
                // sometimes doesn't work correctly
                // if (isValid)
                {
                    using (var context = new PrincipalContext(
                        ContextType.Domain,
                        UserADAuthorization.DomainPath,

                        //Get users from 'OU=Domain Users' only should be uncommented next string 
                        //if wanted to get the whole list objects of the domain next string should be commented

                        //"OU=Domain Users,DC=" + UserADAuthorization.Domain.Split('.')[0] + ",DC=" + UserADAuthorization.Domain.Split('.')[1],

                        UserADAuthorization.Name,
                        UserADAuthorization.Password))
                    {
                        /* DirectoryEntry dir = new DirectoryEntry(context); //DirectoryEntry dir = new DirectoryEntry(" LDAP://intra.vostok.ru/OU=vostok_users,DC=intra,DC=vostok,DC=ru");
                        DirectorySearcher search = new DirectorySearcher(dir);

    search.Filter = "(&(objectCategory=person)(objectClass=user)(!userAccountControl:1.2.840.113556.1.4.803:=2))";
    //Все пользователи	(&(objectCategory=person)(objectClass=user))
    //Все компьютеры (objectCategory=computer)
    //Все контакты	(objectClass=contact)    
    //Все группы	(objectCategory=group)
    //Все организационные подразделения	(objectCategory=organizationalUnit)
    //Пользователи с cn начитающимися на "Вас"	(&(objectCategory=person)(objectClass=user)(cn=Вас*))
    //пользователи с установленным параметром "Срок действия пароля не ограничен"	(&(objectCategory=person)(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=65536))
    //Все отключенные пользователи 	(&(objectCategory=person)(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=2))
    //Все включенные пользователи  	(&(objectCategory=person)(objectClass=user)(!userAccountControl:1.2.840.113556.1.4.803:=2))
    //Пользователи, не требующие паролей 	(&(objectCategory=person)(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=32))

    search.PageSize=1000;
    search.SearchScope = SearchScope.Subtree;
    try
    {
    foreach (SearchResult result in search.FindAll())
    {
    var entry = result.GetDirectoryEntry();
                                string _mail = null, _login = null, _fio = null, _code = null,
                                    _decription = null, _lastLogon = null,
                                    _mailNickName = null, _mailServer = null, _department = null,
                                    _stateAccount = null, stateUAC = null, _sid = null, _guid = null;
                                UACAccountState statesUACOfAccount;
                                int sumOfUACStatesOfPerson = 0;

    res.Add(new main_OrgStruct()
    {
        ADName = entry.Properties["cn"].Value != null ? entry.Properties["cn"].Value.ToString() : "NoN",
        DisplayName = entry.Properties["displayName"].Value != null ? entry.Properties["displayName"].Value.ToString() : "NoN",
        I = entry.Properties["givenName"].Value != null ? entry.Properties["givenName"].Value.ToString() : "NoN",
        F = entry.Properties["sn"].Value != null ? entry.Properties["sn"].Value.ToString() : "NoN",
        EMail = entry.Properties["mail"].Value != null ? entry.Properties["mail"].Value.ToString() : "NoN",
        MobileNumber = entry.Properties["mobile"].Value != null ? long.Parse(entry.Properties["mobile"].Value.ToString()) : 0,
        NumberFull = entry.Properties["telephoneNumber"].Value != null ? long.Parse(entry.Properties["telephoneNumber"].Value.ToString()) : 0,
    });
    }
    }
    catch (Exception e)
    {
    var ee = e.Message;
    }

    *//*
                        using (var UserExt = new UserPrincipalExtended(context))
                        {
                            UserPrincipalExtended foundUser = null;
                            using (var searcher = new PrincipalSearcher(UserExt))
                            {
                                string _mail = null, _login = null, _fio = null, _code = null,
                                    _decription = null, _lastLogon = null,
                                    _mailNickName = null, _mailServer = null, _department = null,
                                    _stateAccount = null, stateUAC = null, _sid = null, _guid = null;
                                UACAccountState statesUACOfAccount;
                                int sumOfUACStatesOfPerson = 0;
                                foreach (var result in searcher.FindAll())
                                {
                                    using (DirectoryEntry de = result.GetUnderlyingObject() as DirectoryEntry)
                                    {
                                        _mail = de?.Properties["mail"]?.Value?.ToString();
                                        _code = de?.Properties["extensionAttribute1"]?.Value?.ToString();
                                        _decription = de?.Properties["description"]?.Value?.ToString()?.Trim()?.ToLower();
                                        _login = de?.Properties["sAMAccountName"]?.Value?.ToString();

                                        // get all logins
                                        if (_login?.Length > 0)
                                        // get only alive logins
                                        //  if (_login?.Length > 0 && _mail != null && _mail.Contains("@") && _code?.Length > 0 &&
                                        //    (!object.Equals(_decription, "dismiss") | !object.Equals(_decription, "fwd")))
                                        {
                                            foundUser = UserPrincipalExtended.FindByIdentity(context, IdentityType.SamAccountName, _login);

                                            _fio = foundUser?.DisplayName?.ToString();

                                            DateTime dt = DateTime.Parse("1970-01-01");
                                            DateTime.TryParse(foundUser?.LastLogon?.ToString(), out dt);
                                            _lastLogon = dt.ToString("yyyy-MM-dd HH:mm:ss");

                                            dt = DateTime.Parse("2200-01-01");

                                            _mailNickName = foundUser?.MailNickname;
                                            _department = foundUser?.Department;
                                            _mailServer = foundUser?.MailServerName;

                                            stateUAC = foundUser?.StateAccount.ToString();
                                            sumOfUACStatesOfPerson = 0;
                                            int.TryParse(stateUAC, out sumOfUACStatesOfPerson);
                                            statesUACOfAccount = new UACAccountState(sumOfUACStatesOfPerson);
                                            _stateAccount = "uac: " + statesUACOfAccount.GetUACStatesOfAccount();

                                            _sid = de?.Properties["sAMAccountName"]?.Value?.ToString();
                                            _guid = foundUser?.Guid?.ToString();
                                            userCount += 1;
                                            ADUsersCollection.Add(new ADUser
                                            {
                                                id = userCount,
                                                login = _login,
                                                stateAccount = _stateAccount,
                                                sid = _sid,
                                                guid = _guid,
                                                mail = _mail,
                                                mailNickName = _mailNickName,
                                                mailServer = _mailServer,
                                                fio = _fio,
                                                code = _code,
                                                description = _decription,
                                                department = _department,
                                                lastLogon = _lastLogon
                                            });
                                        }
                                    }
                                }
                                logger.Trace("ActiveDirectoryGetData,GetDataAD from -= finished =-");
                                statesUACOfAccount = null;
                            }
                            foundUser = null;
                        }
                    }
                }
                logger.Info("ActiveDirectoryGetData, counted users: " + userCount);
                foreach (var user in ADUsersCollection)
                {
                    logger.Trace(
                       user.mailNickName + "| " + user.mail + "| " + user.mailServer + "| " +
                       user.login + "| " + user.code + "| " + user.fio + "| " + user.department + "| " + user.description + "| " +
                       user.lastLogon + "| " + user.stateAccount);
                }
                //   logger.Trace("ActiveDirectoryGetData: User: '" + UserADAuthorization.Name + "' |Password: '" + UserADAuthorization.Password + "' |Domain: '" + UserADAuthorization.Domain + "' |DomainURI: '" + UserADAuthorization.DomainPath + "'");
                return ADUsersCollection;
            }
    */

            /*class NativeMethods*/
            /*
            // it sometimes doesn't work correctly
            static bool ValidateCredentials(UserADAuthorization userADAuthorization)
            {
                IntPtr token;
                bool success = NativeMethods.LogonUser(
                    userADAuthorization.Name,
                    userADAuthorization.Domain,
                    userADAuthorization.Password,
                    NativeMethods.LogonType.Interactive,
                    NativeMethods.LogonProvider.Default,
                    out token);
                if (token != IntPtr.Zero) NativeMethods.CloseHandle(token);
                return success;
            }*/
       /* }
       */
        /*class NativeMethods*/
        /*
        // it sometimes doesn't work correctly
        /// <summary>
        /// Implements P/Invoke Interop calls to the operating system.
        /// </summary>
        internal static class NativeMethods
        {
            /// <summary>
            /// The type of logon operation to perform.
            /// </summary>
            internal enum LogonType :int
            {
                /// <summary>
                /// This logon type is intended for users who will be interactively using the computer, such as a user being logged on by a terminal server, remote shell, or similar process. This logon type has the additional expense of caching logon information for disconnected operations; therefore, it is inappropriate for some client/server applications, such as a mail server.
                /// </summary>
                Interactive = 2,

                /// <summary>
                /// This logon type is intended for high performance servers to authenticate plaintext passwords. The LogonUser function does not cache credentials for this logon type.
                /// </summary>
                Network = 3,

                /// <summary>
                /// This logon type is intended for batch servers, where processes may be executing on behalf of a user without their direct intervention. This type is also for higher performance servers that process many plaintext authentication attempts at a time, such as mail or web servers.
                /// </summary>
                Batch = 4,

                /// <summary>
                /// Indicates a service-type logon. The account provided must have the service privilege enabled.
                /// </summary>
                Service = 5,

                /// <summary>
                /// This logon type is for GINA DLLs that log on users who will be
                /// interactively using the computer.
                /// This logon type can generate a unique audit record that shows
                /// when the workstation was unlocked.
                /// </summary>
                Unlock = 7,

                /// <summary>
                /// This logon type preserves the name and password in the authentication package, which allows the server to make connections to other network servers while impersonating the client. A server can accept plaintext credentials from a client, call LogonUser, verify that the user can access the system across the network, and still communicate with other servers.
                /// </summary>
                NetworkCleartext = 8,

                /// <summary>
                /// This logon type allows the caller to clone its current token and specify new credentials for outbound connections. The new logon session has the same local identifier but uses different credentials for other network connections.
                /// This logon type is supported only by the LOGON32_PROVIDER_WINNT50 logon provider.
                /// </summary>
                NewCredentials = 9
            }

            /// <summary>
            /// Specifies the logon provider.
            /// </summary>
            internal enum LogonProvider :int
            {
                /// <summary>
                /// Use the standard logon provider for the system.
                /// The default security provider is negotiate, unless you pass
                /// NULL for the domain name and the user name is not in UPN format.
                /// In this case, the default provider is NTLM.
                /// NOTE: Windows 2000/NT:   The default security provider is NTLM.
                /// </summary>
                Default = 0,

                /// <summary>
                /// Use this provider if you'll be authenticating against a Windows
                /// NT 3.51 domain controller (uses the NT 3.51 logon provider).
                /// </summary>
                WinNT35 = 1,

                /// <summary>
                /// Use the NTLM logon provider.
                /// </summary>
                WinNT40 = 2,

                /// <summary>
                /// Use the negotiate logon provider.
                /// </summary>
                WinNT50 = 3
            }

            /// <summary>
            /// Logs on the user.
            /// </summary>
            /// <param name="userName">Name of the user.</param>
            /// <param name="domain">The domain.</param>
            /// <param name="password">The password.</param>
            /// <param name="logonType">Type of the logon.</param>
            /// <param name="logonProvider">The logon provider.</param>
            /// <param name="token">The token.</param>
            /// <returns>True if the function succeeds, false if the function fails.
            /// To get extended error information, call GetLastError.</returns>
            [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool LogonUser(
                string userName,
                string domain,
                string password,
                LogonType logonType,
                LogonProvider logonProvider,
                out IntPtr token);

            /// <summary>
            /// Closes the handle.
            /// </summary>
            /// <param name="handle">The handle.</param>
            /// <returns>True if the function succeeds, false if the function fails.
            /// To get extended error information, call GetLastError.</returns>
            [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool CloseHandle(IntPtr handle);
        }
       */




        /*
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
        }*/



    }
}
