//Reference: System.DirectoryServices.AccountManagement
using System.DirectoryServices.AccountManagement;
using System;
using System.Linq;
using System.Collections.ObjectModel;
using System.DirectoryServices;
using System.Collections.Generic;

namespace ASTA.Classes.People
{
    public class ADUsers
    {
        IADUser adUser;
        public List<ADUserFullAccount> ADUsersCollection;

        public delegate void Status<T>(object obj, T data);
        public event Status<TextEventArgs> Trace;
        public event Status<TextEventArgs> Info;

        public ADUsers(IADUser user)
        {
            adUser = user;
             Trace?.Invoke(this, new TextEventArgs(
                $"ADData, Domain Controller: {adUser.DomainControllerPath}, login: { adUser.Login}, password: {adUser.Password}, domain: {adUser.Domain}"));
            // isValid = ValidateCredentials(_ADUserAuthorization);       // it sometimes doesn't work correctly
        }

        public List<ADUserFullAccount> GetADUsersFromDomain()
        {
            Trace?.Invoke(this, new TextEventArgs($"ADData, domain: {adUser.Domain}, login: {adUser.Login}, password: {adUser.Password}"));

            ADUsersCollection = new List<ADUserFullAccount>();

            if (string.IsNullOrEmpty(adUser.Domain))
            { throw new ArgumentNullException("domain can not be null"); }
            else if (string.IsNullOrEmpty(adUser.Login))
            { throw new ArgumentNullException("login can not be null"); }
            else if (string.IsNullOrEmpty(adUser.Password))
            { throw new ArgumentNullException("password can not be null"); }
            else if (string.IsNullOrEmpty(adUser.DomainControllerPath))
            { throw new ArgumentNullException("DomainPath can not be null"); }

            // if (isValid)     //it sometimes doesn't work correctly
            {
                using (var context = new PrincipalContext(
                    ContextType.Domain,
                    adUser.DomainControllerPath,

                   /*1. look starting for users from 'OU=Domain Users' */
                   //"OU=Domain Users,DC=" + _ADUserAuthorization.Domain.Split('.')[0] + ",DC=" + _ADUserAuthorization.Domain.Split('.')[1],
                   /* 2. if need to start from the root of the domain  - previous string should be commented */

                   adUser.Login + @"@" + adUser.Domain,
                    adUser.Password))
                {
                    Trace?.Invoke(this, new TextEventArgs(
                        $"GetADUsers, DC: {adUser.DomainControllerPath}, login: {adUser.Login}@{adUser.Domain}, password: {adUser.Password}"));
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

                            foreach (var result in searcher.FindAll())
                            {
                                using (DirectoryEntry de = result.GetUnderlyingObject() as DirectoryEntry)
                                {
                                    _login = de?.Properties["sAMAccountName"]?.Value?.ToString() ?? string.Empty;
                                    _mail = de?.Properties["mail"]?.Value?.ToString() ?? string.Empty;
                                    _code = de?.Properties["extensionAttribute1"]?.Value?.ToString() ?? string.Empty;
                                    _decription = de?.Properties["description"]?.Value?.ToString()?.Trim()?.ToLower() ?? string.Empty;

                                    stateUAC = de?.Properties["userAccountControl"]?.Value?.ToString() ?? string.Empty;

                                    int.TryParse(stateUAC, out int sumOfUACStatesOfPerson);
                                    statesUACOfAccount = new UACAccountState(sumOfUACStatesOfPerson);
                                    _stateAccount = "uac(" + sumOfUACStatesOfPerson + "): " + statesUACOfAccount.GetUACStatesOfAccount();

                                    //look for accounts with alive logins only
                                    // if (_login?.Length > 0)
                                    //look for accounts with alive logins had email and a code only
                                    if (_login?.Length > 0 &&       //user's info must be had a login
                                        _mail.Contains("@") &&      //user's info must be had an email
                                        _code?.Length > 0 &&        //user's info must be had a code
                                        !_stateAccount.Contains("ACCOUNTDISABLE") && //a disabled account do not write in the collection 
                                      (!_decription.Contains("dismiss") || !_decription.Contains("fwd"))//object.Equals(_decription,
                                      )
                                    {
                                        foundUser = UserPrincipalExtended.FindByIdentity(context, IdentityType.SamAccountName, _login);

                                        DateTime.TryParse(foundUser?.LastLogon?.ToString(), out DateTime dt);
                                        _lastLogon = dt.ToString("yyyy-MM-dd HH:mm:ss");

                                        _fio = foundUser?.DisplayName?.ToString() ?? string.Empty;
                                        _mailNickName = foundUser?.MailNickname ?? string.Empty;
                                        _department = foundUser?.Department ?? string.Empty;
                                        _mailServer = foundUser?.MailServerName ?? string.Empty;
                                        // _sid = foundUser?.Sid?.ToString();
                                        // _guid = foundUser?.Guid?.ToString();

                                        if (!string.IsNullOrEmpty(_fio))
                                        {
                                            ADUsersCollection.Add(new ADUserFullAccount
                                            {
                                                //  id = userCount,
                                                Login = _login,
                                                stateAccount = _stateAccount,
                                                // sid = _sid,
                                                // guid = _guid,
                                                mail = _mail,
                                                mailNickName = _mailNickName,
                                                mailServer = _mailServer,
                                                fio = _fio,
                                                Code = _code,
                                                description = _decription,
                                                department = _department,
                                                lastLogon = _lastLogon
                                            });

                                            if (ADUsersCollection?.Count > 1 && ADUsersCollection?.Count % 5 == 0)
                                            {
                                                Info?.Invoke(this, new TextEventArgs(
                                                    $"Из домена {adUser.Domain} получено {ADUsersCollection?.Count} аккаунтов, последний {_fio.ConvertFullNameToShortForm()}"));
                                            }
                                        }
                                    }
                                }
                            }
                            statesUACOfAccount = null;
                        }
                        foundUser.Dispose();
                    }
                }
            }
            Trace?.Invoke(this, new TextEventArgs($"ActiveDirectoryGetData, collected accounts: {ADUsersCollection?.Count}"));

            foreach (var user in ADUsersCollection)
            {
                Trace?.Invoke(this, new TextEventArgs(
                 $"{user.mailNickName}| {user.mail}| {user.mailServer}| {user.Login}| {user.Code}| { user.fio}| {user.department}| {user.description}| {user.lastLogon}| {user.stateAccount}"));
            }
            return ADUsersCollection;
        }


        /*class NativeMethods*/
        /*
        // it sometimes doesn't work correctly
        static bool ValidateCredentials(_ADUserAuthorization userADAuthorization)
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

    [DirectoryRdnPrefix("CN")]
    [DirectoryObjectClass("user")]
    public class UserPrincipalExtended : UserPrincipal
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

    public class UACAccountState
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

        System.Collections.Generic.Dictionary<int, string> dicOfUACs;//dictionary with all of existed acc.states in the digital and string forms
        static int[] statesUAC; //all of existed acc.states in the digital forms
        int shiftStart; //next position in digital form to calculate acc.states
        int sumOfUACStates; //sum all of existed acc.states
        int numberOfUACStates; //number all of existed acc.states
        int indexOfFoundState; //index of last found acc.state
        int _flag; //sum of UAC states of Person's account in AD
        static string _result; //decomposition information of Account states

        public UACAccountState(int sumOfStates)
        {
            dicOfUACs = new System.Collections.Generic.Dictionary<int, string>()
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
