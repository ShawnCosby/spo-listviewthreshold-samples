using Microsoft.SharePoint.Client;
using System.Security;

namespace SpoLvtTest
{
    public static class ClientContextUtility
    {
        public static ClientContext GetClientContext(string url, string username, string password)
        {
            var context = new ClientContext(url)
            {
                Credentials = new SharePointOnlineCredentials(username, ToSecureString(password)),
            };

            //Using best practices to help avoid SPO throttling, ensure the web requests decorate the User-agent HTTP header 
            context.SetTrafficDecorator();

            return context;
        }

        private static SecureString ToSecureString(string s)
        {
            var secure = new SecureString();

            foreach (char c in s.ToCharArray())
            {
                secure.AppendChar(c);
            }

            secure.MakeReadOnly();

            return secure;
        }
    }
}
