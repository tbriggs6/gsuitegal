using Google.Apis.Admin.Directory.directory_v1;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace GSuiteGAL
{
    public class GSuiteDirectory
    {
        private static string[] Scopes = { DirectoryService.Scope.AdminDirectoryUserReadonly };
        private static string ApplicationName = "GSuiteGAL";
        private UserCredential credential;

        public List<Address> entries { get; } = new List<Address>();
        private Config config = new Config();

        public void retrieveAddresses()
        {
            // If modifying these scopes, delete your previously saved credentials
            // at ~/.credentials/admin-directory_v1-dotnet-quickstart.json

            string CredFileName = config.installPathName + "\\" + config.credentialFileName;
            using (var stream = new FileStream(CredFileName, FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string tokenPath = config.installPathName + "\\" + config.tokenFileName;
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(tokenPath, true)).Result;
            }

            // Create Directory API service.
            var service = new DirectoryService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            string page_token = "";
            do
            {

                // Define parameters of request.
                UsersResource.ListRequest request = service.Users.List();
                request.Customer = "my_customer";
                request.MaxResults = 100;
                request.OrderBy = UsersResource.ListRequest.OrderByEnum.Email;
                if (page_token.Length > 0)
                {
                    request.PageToken = page_token;
                }

                // List users.
                var response = request.Execute();
                page_token = response.NextPageToken;
                var users = response.UsersValue;

                if (users != null && users.Count > 0)
                {
                    foreach (var userItem in users)
                    {
                        Address a = new Address(userItem.PrimaryEmail, userItem.Name.FullName);
                        entries.Add(a);
                    }
                }

            } while (page_token != null);
        } // end run
    } // end class
} // end name space

