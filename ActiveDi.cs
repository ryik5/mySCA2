using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASTA
{
    //Reference: System.DirectoryServices.AccountManagement
    using System.Collections.Generic;
    using System.DirectoryServices.AccountManagement;
    using System;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Collections.ObjectModel;
    using System.Collections.Specialized;
    using System.DirectoryServices;

    namespace ASTA
    {
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
    }

}
