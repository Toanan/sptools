using Microsoft.SharePoint.Client;


namespace SPtools
{
    public static class Authentication
    {
        /// <summary>
        /// Retrieve ClientContext from username and password
        /// </summary>
        /// <param name="url">The SharePointOnline site url</param>
        /// <param name="UserName">The account user name</param>
        /// <param name="PassWord">The account password</param>
        /// <returns>An authenticated ClientContext</returns>
        public static ClientContext getSPContext(string url, string UserName, System.Security.SecureString PassWord)
        {
            var ctx = new ClientContext(url);
            ctx.Credentials = new SharePointOnlineCredentials(UserName, PassWord);

                return ctx;
        }


    }
}
