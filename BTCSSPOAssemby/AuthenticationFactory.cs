using Microsoft.SharePoint.Client;
using System.Net;
using OfficeDevPnP.Core;

namespace Btcs.Authentication
{
    public static class AuthenticationFactory
    {
        /// <summary>
        /// Retrieve an authenticated ClientContext using ICredentials
        /// </summary>
        /// <param name="Url">The SharePoint Online site Url</param>
        /// <param name="credentials">The credentials associated</param>
        /// <returns>A tested authenticated ClientContext</returns>
        public static ClientContext GetAuthenticatedContext(string Url, ICredentials credentials)
        {
            var context = new ClientContext(Url);
            context.Credentials = credentials;

            // We try to call for the SPO site title
            validateAuthentication(context);
            return context;
        }

        /// <summary>
        /// REtrieve an authenticated ClientContext using AppOnly auth
        /// </summary>
        /// <param name="Url">The SharePoint Online site Url</param>
        /// <param name="AppID">The App id</param>
        /// <param name="AppSecret">The App secret</param>
        /// <returns>A tested authenticated ClientContext</returns>
        public static ClientContext GetAuthenticatedContext(string Url, string AppID, string AppSecret)
        {
            var am = new OfficeDevPnP.Core.AuthenticationManager();
            var context = am.GetAppOnlyAuthenticatedContext(Url, AppID, AppSecret);

            // We try to call for the SPO site title
            validateAuthentication(context);
            return context;
        }

        /// <summary>
        /// Validate the Authentication by calling for the web.title
        /// </summary>
        /// <param name="context">The clientcontext to validate</param>
        private static void validateAuthentication(ClientContext context)
        {
            using (context)
            {
                context.Load(context.Web, w => w.Title);
                context.ExecuteQuery();
            }
        }
    }
}
