//Reference: System.DirectoryServices.AccountManagement
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System;
using System.Linq;
//using System.Runtime.InteropServices;
using System.Collections.ObjectModel;
//using System.Collections.Specialized;
using System.DirectoryServices;

namespace ASTA
{
    class UserADAuthorization
    {
        public string Name { get; set; }       // имя
        public string Domain { get; set; }      // домен
        public string Password { get; set; }    // пароль
        public string DomainPath { get; set; }    // URI сервера

        public override string ToString()
        {
            return Name + "\t" + Domain + "\t" + Password + "\t" + DomainPath;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            UserADAuthorization df = obj as UserADAuthorization;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    class ActiveDirectoryData
    {
        static NLog.Logger logger;
        UserADAuthorization UserADAuthorization;

        public ObservableCollection<ADUser> ADUsersCollection;

        public ActiveDirectoryData(string _user, string _domain, string _password, string _domainPath)
        {
            logger = NLog.LogManager.GetCurrentClassLogger();
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

                    /*1. look starting for users from 'OU=Domain Users' */
                    //"OU=Domain Users,DC=" + UserADAuthorization.Domain.Split('.')[0] + ",DC=" + UserADAuthorization.Domain.Split('.')[1],
                    /* 2. if need to start from the root of the domain  - previous string should be commented */

                    UserADAuthorization.Name,
                    UserADAuthorization.Password))
                {
                    using (var UserExt = new UserPrincipalExtended(context))
                    {
                        UserPrincipalExtended foundUser = null;
                        using (var searcher = new PrincipalSearcher(UserExt))
                        {
                            string _mail = null, _login = null, _fio = null, _code = null,
                                _decription = null, _lastLogon = null,
                                _mailNickName = null, _mailServer = null, _department = null,
                                _stateAccount = null, stateUAC = null;
                            UACAccountState statesUACOfAccount;
                            int sumOfUACStatesOfPerson = 0;
                            foreach (var result in searcher.FindAll())
                            {
                                using (DirectoryEntry de = result.GetUnderlyingObject() as DirectoryEntry)
                                {
                                    _login = de?.Properties["sAMAccountName"]?.Value?.ToString() ?? string.Empty;
                                    _mail = de?.Properties["mail"]?.Value?.ToString() ?? string.Empty;
                                    _code = de?.Properties["extensionAttribute1"]?.Value?.ToString() ?? string.Empty;
                                    _decription = de?.Properties["description"]?.Value?.ToString()?.Trim()?.ToLower() ?? string.Empty;

                                    stateUAC = de?.Properties["userAccountControl"]?.Value?.ToString() ?? string.Empty;
                                    sumOfUACStatesOfPerson = 0;
                                    int.TryParse(stateUAC, out sumOfUACStatesOfPerson);
                                    statesUACOfAccount = new UACAccountState(sumOfUACStatesOfPerson);
                                    _stateAccount = "uac(" + sumOfUACStatesOfPerson + "): " + statesUACOfAccount.GetUACStatesOfAccount();

                                    //look for accounts with alive logins only
                                    // if (_login?.Length > 0)
                                    //look for accounts with alive logins had email and a code only
                                    if (_login?.Length > 0 &&       //user's info must be had a login
                                        _mail.Contains("@") &&      //user's info must be had an email
                                        _code?.Length > 0 &&        //user's info must be had a code
                                        !_stateAccount.Contains("ACCOUNTDISABLE") && //a disabled account do not write in the collection 
                                      (!_decription.Contains("dismiss") | !!_decription.Contains("fwd"))//object.Equals(_decription,
                                      )
                                    {
                                        foundUser = UserPrincipalExtended.FindByIdentity(context, IdentityType.SamAccountName, _login);

                                        _fio = foundUser?.DisplayName?.ToString() ?? string.Empty;

                                        DateTime dt = DateTime.Parse("1970-01-01");
                                        DateTime.TryParse(foundUser?.LastLogon?.ToString(), out dt);
                                        _lastLogon = dt.ToString("yyyy-MM-dd HH:mm:ss");

                                        dt = DateTime.Parse("2200-01-01");

                                        _mailNickName = foundUser?.MailNickname ?? string.Empty;
                                        _department = foundUser?.Department ?? string.Empty;
                                        _mailServer = foundUser?.MailServerName ?? string.Empty;

                                        // stateUAC = foundUser?.StateAccount.ToString() ?? string.Empty;

                                        // _sid = foundUser?.Sid?.ToString();
                                        // _guid = foundUser?.Guid?.ToString();

                                        userCount += 1;
                                        ADUsersCollection.Add(new ADUser
                                        {
                                            id = userCount,
                                            login = _login,
                                            stateAccount = _stateAccount,
                                            // sid = _sid,
                                            // guid = _guid,
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
                        foundUser.Dispose();
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
    }

    class SendToFileShare
    {
        private SendToFileShare( string fileName, ADUser user, string pathNetworkShare)
        {
           // string networkShareLocation = @"\\your\network\share\";

            var path = $"{pathNetworkShare}{fileName}"; //$"{pathNetworkShare}{fileName}.pdf";
            var fileByte= System.IO.File.ReadAllBytes(fileName);
            //Credentials for the account that has write-access. Probably best to store these in a web.config file.
            var domain = user.domain;
            var userID =user.login;
            var password =user.password;


            if (ImpersonateUser(domain, userID, password) == true)
            {
                //write the PDF to the share:
                System.IO.File.WriteAllBytes(path, fileByte);
            }
            else
            {
                //Could not authenticate account. Something is up.
                //Log or something.
            }
        }

        /// <summary>
        /// Impersonates the given user during the session.
        /// </summary>
        /// <param name="domain">The domain.</param>
        /// <param name="userName">Name of the user.</param>
        /// <param name="password">The password.</param>
        /// <returns></returns>
        private bool ImpersonateUser(string domain, string userName, string password)
        {
            System.Security.Principal.WindowsIdentity tempWindowsIdentity;
            IntPtr token = IntPtr.Zero;
            IntPtr tokenDuplicate = IntPtr.Zero;

            if (RevertToSelf())
            {
                if (LogonUserA(userName, domain, password, LOGON32_LOGON_INTERACTIVE,
                    LOGON32_PROVIDER_DEFAULT, ref token) != 0)
                {
                    if (DuplicateToken(token, 2, ref tokenDuplicate) != 0)
                    {
                        tempWindowsIdentity = new System.Security.Principal.WindowsIdentity(tokenDuplicate);
                        impersonationContext = tempWindowsIdentity.Impersonate();
                        if (impersonationContext != null)
                        {
                            CloseHandle(token);
                            CloseHandle(tokenDuplicate);
                            return true;
                        }
                    }
                }
            }
            if (token != IntPtr.Zero)
                CloseHandle(token);
            if (tokenDuplicate != IntPtr.Zero)
                CloseHandle(tokenDuplicate);
            return false;
        }

        /// <summary>
        /// Undoes the current impersonation.
        /// </summary>
        private void undoImpersonation()
        {
            impersonationContext.Undo();
        }


        #region Impersionation global variables
        public const int LOGON32_LOGON_INTERACTIVE = 2;
        public const int LOGON32_PROVIDER_DEFAULT = 0;

        System.Security.Principal.WindowsImpersonationContext impersonationContext;

        [System.Runtime.InteropServices.DllImport("advapi32.dll")]
        public static extern int LogonUserA(String lpszUserName,
            String lpszDomain,
            String lpszPassword,
            int dwLogonType,
            int dwLogonProvider,
            ref IntPtr phToken);
        [System.Runtime.InteropServices.DllImport("advapi32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, SetLastError = true)]
        public static extern int DuplicateToken(IntPtr hToken,
            int impersonationLevel,
            ref IntPtr hNewToken);

        [System.Runtime.InteropServices.DllImport("advapi32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, SetLastError = true)]
        public static extern bool RevertToSelf();

        [System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public static extern bool CloseHandle(IntPtr handle);
        #endregion
    }

    class ADUser : IComparable<ADUser>
    {
        public int id;
        public string domain;
        public string login;
        public string password;
        public string code;
        public string mailNickName;
        public string mail;
        public string mailServer;
        public string description;
        public string lastLogon;
        public string fio;
        public string department;
        public string stateAccount;

        //Для возможности поиска дубляжного значения
        public override string ToString()
        {
            return fio + "\t" + department + "\t" + code + "\t" +
                mail + "\t" + login + "\t" + stateAccount + "\t" 
                + description + "\t" + lastLogon + "\t" 
                + domain + "\t" + password;
        }

        public override bool Equals(object obj)
        {
            if (obj == null || !(obj is ADUser))
                return false;

            ADUser df = obj as ADUser;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        //реализация для выполнения сортировки
        int IComparable<ADUser>.CompareTo(ADUser next)
        {
            return new ADUsersComparer().Compare(this, next);
        }

        public string CompareTo(ADUser next)
        {
            return next.CompareTo(this);
        }

    }

    //additional class для выполнения сортировки
    class ADUsersComparer : IComparer<ADUser>
    {
        public int Compare(ADUser x, ADUser y)
        {
            return this.CompareTwoStaffADs(x, y);
        }

        public int CompareTwoStaffADs(ADUser x, ADUser y)
        {
            string a = x.fio + x.login;
            string b = y.fio + y.login;

            string[] words = { a, b };
            Array.Sort(words);

            if (words[0] != a)
            {
                return 1;
            }
            else if (a == b)
            {
                return 0;
            }
            else
            {
                return -1;
            }
        }
    }



    [DirectoryRdnPrefix("CN")]
    [DirectoryObjectClass("user")]
    class UserPrincipalExtended : UserPrincipal
    {
        public UserPrincipalExtended(PrincipalContext context) : base(context) { }

        public UserPrincipalExtended(PrincipalContext
            context,
            string samAccountName,
            string password,
            bool enabled)
            : base(context, samAccountName, password, enabled) { }

        public static new UserPrincipalExtended FindByIdentity(PrincipalContext context,
                                                       string identityValue)
        {
            return (UserPrincipalExtended)FindByIdentityWithType(context,
                                                         typeof(UserPrincipalExtended),
                                                         identityValue);
        }

        public static new UserPrincipalExtended FindByIdentity(PrincipalContext context,
                                                       IdentityType identityType,
                                                       string identityValue)
        {
            return (UserPrincipalExtended)FindByIdentityWithType(context,
                                                         typeof(UserPrincipalExtended),
                                                         identityType,
                                                         identityValue);
        }

        //additional(Extended) properties
        #region custom attributes

        [DirectoryProperty("RealLastLogon")]
        public DateTime? RealLastLogon
        {
            get
            {
                if (ExtensionGet("LastLogon").Length > 0)
                {
                    var lastLogonDate = ExtensionGet("LastLogon")[0];
                    var lastLogonDateType = lastLogonDate.GetType();

                    var highPart = (Int32)lastLogonDateType.InvokeMember("HighPart",
                        System.Reflection.BindingFlags.GetProperty, null, lastLogonDate, null);
                    var lowPart = (Int32)lastLogonDateType.InvokeMember("LowPart",
                        System.Reflection.BindingFlags.GetProperty |
                        System.Reflection.BindingFlags.Public, null, lastLogonDate, null);

                    var longDate = ((Int64)highPart << 32 | (UInt32)lowPart);

                    return longDate > 0 ? (DateTime?)DateTime.FromFileTime(longDate) : null;
                }

                return null;
            }
        }

        [DirectoryProperty("HomePage")]
        public string HomePage
        {
            get
            {
                if (ExtensionGet("HomePage").Length != 1)
                    return null;
                return (string)ExtensionGet("HomePage")[0];
            }
            set { this.ExtensionSet("HomePage", value); }
        }

        [DirectoryProperty("extensionAttribute1")] //Code
        public string ExtensionAttribute1
        {
            get
            {
                if (ExtensionGet("extensionAttribute1").Length != 1)
                    return null;
                return (string)ExtensionGet("extensionAttribute1")[0];
            }
            set { this.ExtensionSet("extensionAttribute1", value); }
        }

        [DirectoryProperty("department")] //Department
        public string Department
        {
            get
            {
                if (ExtensionGet("department").Length != 1)
                    return null;
                return (string)ExtensionGet("department")[0];
            }
            set { this.ExtensionSet("department", value); }
        }

        [DirectoryProperty("mailNickname")] //Mail User's NickName
        public string MailNickname
        {
            get
            {
                if (ExtensionGet("mailNickname").Length != 1)
                    return null;
                return (string)ExtensionGet("mailNickname")[0];
            }
            set { this.ExtensionSet("mailNickname", value); }
        }

        [DirectoryProperty("msExchHomeServerName")] //Mail Server's Name
        public string MailServerName
        {
            get
            {
                if (ExtensionGet("msExchHomeServerName").Length != 1)
                    return null;
                string exchFullName = (string)ExtensionGet("msExchHomeServerName")[0];
                int lastSeparator = exchFullName.LastIndexOf('=') + 1;
                var result = exchFullName.Substring(lastSeparator);
                return result;
            }
            set { this.ExtensionSet("msExchHomeServerName", value); }
        }

        [DirectoryProperty("company")]
        public string Company
        {
            get
            {
                if (ExtensionGet("company").Length != 1)
                    return null;
                return (string)ExtensionGet("company")[0];
            }
            set
            {
                this.ExtensionSet("company", value);
            }
        }

        [DirectoryProperty("userAccountControl")]
        public int StateAccount
        {
            get
            {
                if (ExtensionGet("userAccountControl").Length != 1)
                    return -1;
                return (int)ExtensionGet("userAccountControl")[0];
            }
            set
            {
                this.ExtensionSet("userAccountControl", value);
            }
        }

        #endregion


    }

    class UACAccountState
    {
        // existed acc.states in the string and digital forms
        const int DONT_EXPIRE_PASSWORD = 65536;
        const int NORMAL_ACCOUNT = 512;
        const int PASSWD_CANT_CHANGE = 64;
        const int PASSWD_NOTREQD = 32;
        const int LOCKOUT = 16;
        const int HOMEDIR_REQUIRED = 8;
        const int ACCOUNTDISABLE = 2;
        const int NOPE = 0;

        Dictionary<int, string> dicOfUACs;//dictionary with all of existed acc.states in the digital and string forms
        static int[] statesUAC; //all of existed acc.states in the digital forms
        int shiftStart; //next position in digital form to calculate acc.states
        int sumOfUACStates; //sum all of existed acc.states
        int numberOfUACStates; //number all of existed acc.states
        int indexOfFoundState; //index of last found acc.state
        int _flag; //sum of UAC states of Person's account in AD
        static string _result; //decomposition information of Account states

        public UACAccountState(int sumOfStates)
        {
            dicOfUACs = new Dictionary<int, string>()
            {
                [DONT_EXPIRE_PASSWORD] = " DONT_EXPIRE_PASSWORD",
                [NORMAL_ACCOUNT] = "NORMAL_ACCOUNT",
                [PASSWD_CANT_CHANGE] = "PASSWD_CANT_CHANGE",
                [PASSWD_NOTREQD] = "PASSWD_NOTREQD",
                [LOCKOUT] = "LOCKOUT",
                [HOMEDIR_REQUIRED] = "HOMEDIR_REQUIRED",
                [ACCOUNTDISABLE] = "ACCOUNTDISABLE",
                [NOPE] = "NOPE"
            };
            statesUAC = new int[] { //length=8
                NOPE, ACCOUNTDISABLE, HOMEDIR_REQUIRED, LOCKOUT, PASSWD_NOTREQD,
             PASSWD_CANT_CHANGE, NORMAL_ACCOUNT, DONT_EXPIRE_PASSWORD
            };
            sumOfUACStates = statesUAC.Sum();
            numberOfUACStates = statesUAC.Length;
            indexOfFoundState = numberOfUACStates - 1;
            shiftStart = NOPE;
            _flag = sumOfStates;
            _result = string.Empty;
        }

        public string GetUACStatesOfAccount()
        {
            return GetStatesOfAccount(_flag);
        }

        private string GetStatesOfAccount(int sumStatesOfPerson) //state: /1st st. state=66050  /2nd st. state=514  /3rd st. state=2
        {
            bool foundUAC = false;

            if (sumStatesOfPerson == 0 || sumStatesOfPerson > sumOfUACStates)
                return _result;

            for (int k = 0; k < numberOfUACStates; k++)//number: /1st st.= 66050  /2nd st. =514  /3rd st. =2
            {
                if (sumStatesOfPerson < statesUAC[k])  //k: /1st st. k>7 /2nd st. k=7 /3rd st. k=1
                {
                    indexOfFoundState = k - 1; //indexOfFoundState: /2nd st.  =6  /3rd st.  =1
                    foundUAC = true;  //foundUAC: /2nd st.,3rd st.   = true
                    break;
                }
            }
            _result += dicOfUACs[statesUAC[indexOfFoundState]];  //_result:  /1st st. ="65536 "  /2nd st. ="65536 512 " /3rd st. ="65536 512 2 "
            shiftStart = sumStatesOfPerson - statesUAC[indexOfFoundState]; //shiftStart: /1st st. 66050-65536=514  /2nd st. 514-512=2

            if (!foundUAC || shiftStart != NOPE)
            {
                _result += "|";  //_result:  /1st st. ="65536 "  /2nd st. ="65536 512 " /3rd st. ="65536 512 2 "
                GetStatesOfAccount(shiftStart);  //shiftStart: /1st st. 514 /2nd stage  =2
                return _result;
            }
            else
                return _result;
        }

    }

}
