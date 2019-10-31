using System;
using System.Diagnostics.Contracts;
using ASTA.Classes.People; 

namespace ASTA.Updating
{  
    
    /* public class SendToFileShare
    {
        private UserAD _user;
        private string _source;
        private string _target;

        public SendToFileShare(UserAD user, string source, string target)
        {
            // var path = $"{pathNetworkShare}{fileName}";         // path = @"\\your\network\share\fileName"

            Contract.Requires(source.Length > 0);
            Contract.Requires(target.Length > 0);
            Contract.Requires(!source.Equals(target));


            if (ImpersonateUser(user) == true)
            {
                System.IO.File.Copy(source, target, true);  //@"\\server\folder\Myfile.txt"

                // other way todo it
                // var fileByte = System.IO.File.ReadAllBytes(source);
                // System.IO.File.WriteAllBytes(target, fileByte);
            }
            else
            {
                System.IO.File.Copy(source, target, true);  //@"\\server\folder\Myfile.txt"
            }
        }

        public SendToFileShare(UserAD user, string[] source, string[] target)
        {
            // var path = $"{pathNetworkShare}{fileName}";         // path = @"\\your\network\share\fileName"
            if (!(source?.Length > 0))
            {
                throw new ArgumentNullException("Source is null");
            }

            if (source?.Length != target?.Length)
            {
                throw new Exception("Amount elements of source does not equal ones of target");
            }

            for (int index = 0; index < source.Length; index++)
            {
                if (source[index].Equals(target[index]))
                {
                    throw new Exception("There is an attempt to copy the file to the same place");
                }
            }

            if (ImpersonateUser(user) == true)
            {
                for (int index = 0; index < source.Length; index++)
                {
                    System.IO.File.Copy(source[index], target[index], true);
                }
                //@"\\server\folder\Myfile.txt"

                // other way todo it
                // var fileByte = System.IO.File.ReadAllBytes(source);
                // System.IO.File.WriteAllBytes(target, fileByte);
            }
         }

        /// <summary>
        /// Impersonates the given user during the session.
        /// </summary>
        /// <param name="domain">The domain.</param>            //user.domain;
        /// <param name="userName">Name of the user.</param>    //user.login;
        /// <param name="password">The password.</param>        // user.password;
        /// <returns></returns>
        private bool ImpersonateUser(UserAD user)
        {
            System.Security.Principal.WindowsIdentity tempWindowsIdentity;
            IntPtr token = IntPtr.Zero;
            IntPtr tokenDuplicate = IntPtr.Zero;

            if (RevertToSelf())
            {
                if (LogonUserA(user.login, user.domain, user.password, LOGON32_LOGON_INTERACTIVE,
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
        private void UndoImpersonation()
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
*/
}
