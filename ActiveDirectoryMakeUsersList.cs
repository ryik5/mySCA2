//Reference: System.DirectoryServices.AccountManagement
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Security.Principal;
using System.DirectoryServices.AccountManagement;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.DirectoryServices;
using System.Threading;
using NLog;

namespace ASTA
{

    interface IUserADAuthorization
    {
        string Name { get; set; }        // имя
        string Domain { get; set; }      // домен
        string Password { get; set; }    // пароль
    }
    public class UserADAuthorization :IUserADAuthorization
    {
        public string Name { get; set; }       // имя
        public string Domain { get; set; }      // домен
        public string Password { get; set; }    // пароль

        /* public static UserADAuthorizationBuilder CreateBuilder()
         {
             return new UserADAuthorizationBuilder();
         }*/
    }



    public class UserADAuthorizationBuilder
    {
        private UserADAuthorization user;
        public UserADAuthorizationBuilder()
        {
            user = new UserADAuthorization();
        }
        public UserADAuthorizationBuilder SetName(string name)
        {
            user.Name = name;
            return this;
        }
        public UserADAuthorizationBuilder SetDomain(string domain)
        {
            user.Domain = domain;
            return this;
        }
        public UserADAuthorizationBuilder SetPassword(string password)
        {
            user.Password = password;
            return this;
        }

        public UserADAuthorization Build()
        {
            return user;
        }

        public static implicit operator UserADAuthorization(UserADAuthorizationBuilder builder)
        {
            return builder.user;
        }
    }

    public class ActiveDirectoryGetData//: Mediator
    {
        static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        // UserADAuthorization UserADAuthorization;
        //  bool isValid = false;
        string user, domain, password, domainPath;
        PrincipalSearcher principalSearcher;

        public ActiveDirectoryGetData(string _user, string _domain, string _password, string _domainPath)
        {
            user = _user;
            domain = _domain;
            password = _password;
            domainPath = _domainPath;
            principalSearcher = GetDataAD();
        }

        private PrincipalSearcher GetDataAD()
        {
            PrincipalSearcher principalSearcher = new PrincipalSearcher();
            //   logger.Trace(DomainPath);

            //UserADAuthorization = new UserADAuthorizationBuilder().SetName(_user).SetPassword(_password).SetDomain(_domain).Build();
            // isValid = ValidateCredentials(UserADAuthorization); //sometimes doesn't work correctly
            //  logger.Info("Доступ к домену '" + _domain + "' предоставлен: " + isValid);
            
            // if (isValid)
            {
                using (var context = new PrincipalContext(ContextType.Domain, domainPath, "OU=Domain Users,DC=" + domain.Split('.')[0] + ",DC=" + domain.Split('.')[1], user, password))
                {
                    using (var searcher = new PrincipalSearcher(new UserPrincipal(context)))
                    {
                        principalSearcher = searcher;

                        
                    }
                }
            }
            //   else
            //  {
            logger.Trace("ActiveDirectoryGetData: User: '" + user + "' |Password: '" + password + "' |Domain: '" + domain + "' |DomainURI: '" + domainPath + "'");
            //   }
            return principalSearcher;
        }

        public StaffMemento SaveObjects()
        {
            return new StaffMemento(principalSearcher);
        }

        // sometimes doesn't work correctly
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
        }
    }

    // Memento
    public class StaffMemento
    {
        public PrincipalSearcher principalSearcher { get; private set; }

        public StaffMemento(PrincipalSearcher _principalSearcher)
        {
            this.principalSearcher = _principalSearcher;
        }
    }
    // Caretaker
    class StaffStore
    {
        public Stack<StaffMemento> Story { get; private set; }
        public StaffStore()
        {
            Story = new Stack<StaffMemento>();
        }
    }


    class MakeADUsersTable
    {
        static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        PrincipalSearcher searcher;

        public MakeADUsersTable(StaffMemento memento)
        {
            this.searcher = memento.principalSearcher;
            DoWork(this.searcher);
        }

        private void DoWork(PrincipalSearcher searcher)
        {
            string mail, code, decription;
            foreach (var result in searcher.FindAll())
            {
                DirectoryEntry de = result.GetUnderlyingObject() as DirectoryEntry;

                mail = de?.Properties["mail"]?.Value?.ToString()?.Trim();
                code = de?.Properties["extensionAttribute1"]?.Value?.ToString()?.Trim();
                try
                {
                    decription = de.Properties["description"].Value.ToString().ToLower().Trim();
                } catch { decription = ""; }

                try
                {
                    if (code?.Length > 0 && mail.Contains("@") && !decription.Equals("dismiss")) //
                    {
                        //todo 
                        //fill struct-Table
                        logger.Trace(
                             de?.Properties["mail"]?.Value + "| " +
                              // de?.Properties["mailNickname"]?.Value + "| " +
                              de?.Properties["sAMAccountName"]?.Value + "| " +
                              de?.Properties["extensionAttribute1"]?.Value + "| " +
                              de?.Properties["displayName"]?.Value
                             );
                    }
                } catch { }
            }
        }
    }



    // sometimes doesn't work correctly!!!!! Check it.
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
}
