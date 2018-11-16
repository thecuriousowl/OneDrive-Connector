using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace OneDrive_Connector.Controllers
{
    class GraphHelpers
    {
        public static List<User> GetAllEnabledUsers(List<String> exclusion)
        {
            GraphServiceClient graphClient = Authentication.GetAuthenticatedClient();

            List<User> teneoAll = new List<User>();
            // Get full licensed user list 
            int count = 1;
            System.Console.WriteLine("Requesting Page " + count);
            var users = graphClient.Users.Request().Top(200).Select("DisplayName,GivenName,Id,AssignedLicenses").GetAsync().Result;
            count++;
            var page = users.CurrentPage;
            teneoAll.AddRange(page);
            
            while (users.NextPageRequest != null)
            {
                System.Console.WriteLine("Requesting Page " + count);
                users = users.NextPageRequest.GetAsync().Result;
                count++;
                teneoAll.AddRange(users.CurrentPage);
            }

            List<User> enabledUsers = new List<User>();
            foreach (var user in teneoAll)
            {
                List<AssignedLicense> licensing = new List<AssignedLicense>();
                licensing.AddRange(user.AssignedLicenses);
                if (licensing.Count > 0 && !(exclusion.Contains(user.Id)))
                {
                    enabledUsers.Add(user);
                }
            }

            return enabledUsers;
        }

        public static List<User> UsersUpdatedStatus(List<User> input)
        {
            List<User> result = new List<User>();
            foreach(var user in input)
            {
                var root = user.MySite;
                if(root.Contains("teneo_com"))
                {
                    result.Add(user);
                }
            }

            return result;
        }
    }
}
