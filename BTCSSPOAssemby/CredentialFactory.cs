using System.Net;
using System.Security;
using Microsoft.SharePoint.Client;
using CredentialManagement;
using OfficeDevPnP.Core.Utilities;

namespace Btcs.Credentials
{
    public static class CredentialFactory
    {
        /// <summary>
        /// Describe the authentication type to create the credentials accordingly
        /// </summary>
        public enum CredentialsType
        {
            /// <summary>
            /// SharePointOnlineCredentials
            /// </summary>
            SharePointOnline,
            /// <summary>
            /// NetworkCredentials
            /// </summary>
            SharePointActiveDirectory
        }

        /// <summary>
        /// Create a credentials object according to the authentication type
        /// </summary>
        /// <param name="userName">UserName or AppID</param>
        /// <param name="passWord">Password or AppSecret</param>
        /// <param name="authenticationType">The authentication type</param>
        /// <returns>ICredential implementation object</returns>
        public static ICredentials CreateCredentials(string userName, string passWord, CredentialsType authenticationType)
        {
            ICredentials result = null;

            switch (authenticationType)
            {
                case CredentialsType.SharePointOnline:
                    SecureString securePassword = new SecureString();
                    foreach (char c in passWord.ToCharArray())
                    {
                        securePassword.AppendChar(c);
                    }
                    result = new SharePointOnlineCredentials(userName, securePassword);
                    break;
                case CredentialsType.SharePointActiveDirectory:
                    result = new NetworkCredential(userName, passWord);
                    break;
            }
            return result;
        }

        /// <summary>
        /// Create a credentials object according to the authentication type
        /// </summary>
        /// <param name="userName">UserName</param>
        /// <param name="securePassword">Password</param>
        /// <param name="authenticationType">The authentication Type</param>
        /// <returns>ICredential implementation object</returns>
        public static ICredentials CreateCredentials(string userName, SecureString securePassword, CredentialsType authenticationType)
        {
            ICredentials result = null;

            switch (authenticationType)
            {
                case CredentialsType.SharePointOnline:
                    result = new SharePointOnlineCredentials(userName, securePassword);
                    break;
                case CredentialsType.SharePointActiveDirectory:
                    result = new NetworkCredential(userName, securePassword);
                    break;
            }
            return result;
        }

        /// <summary>
        /// Save credentials in the windows credential manager
        /// </summary>
        /// <param name="Name">Credential label (name)</param>
        /// <param name="UserName">The UserName</param>
        /// <param name="SecurePassWord">The secure password</param>
        /// <returns>A SharepointOnlineCredentials object</returns>
        public static SharePointOnlineCredentials StoreCredentials(string Name, string UserName, SecureString SecurePassWord)
        {
            Credential Cred = new Credential
            {
                Target = Name,
                Username = UserName,
                SecurePassword = SecurePassWord,
                PersistanceType = PersistanceType.LocalComputer,
                Type = CredentialType.Generic,
            };
            Cred.Save();
            SharePointOnlineCredentials cred = CredentialManager.GetSharePointOnlineCredential(Name);
            return cred;
        } 

        /// <summary>
        /// Retrieve credentials from the windows credential manager
        /// </summary>
        /// <param name="Name">Credential label (name)</param>
        /// <param name="authenticationType">The authentication Type</param>
        /// <returns></returns>
        public static ICredentials GetStoredCredentials(string Name, CredentialsType authenticationType)
        {
            ICredentials result = null;

            switch (authenticationType)
            {
                case CredentialsType.SharePointOnline:
                    result = CredentialManager.GetSharePointOnlineCredential(Name);
                    break;
                case CredentialsType.SharePointActiveDirectory:
                    result = CredentialManager.GetCredential(Name);
                    break;
            }
            return result;
        }
    }
}
