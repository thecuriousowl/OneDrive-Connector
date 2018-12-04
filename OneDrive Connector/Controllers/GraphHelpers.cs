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
            var users = graphClient.Users.Request().Top(200).Select("DisplayName,GivenName,Id,Mail,AssignedLicenses").GetAsync().Result;
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

            var excludedUsers = ParseGroup("691f100e-8565-481b-a99b-1fcd4e85eee4");

            List<User> enabledUsers = new List<User>();
            foreach (var user in teneoAll)
            {
                List<AssignedLicense> licensing = new List<AssignedLicense>();
                licensing.AddRange(user.AssignedLicenses);
                if (licensing.Count > 0 && !(exclusion.Contains(user.Id)) && !(IsMember(user, excludedUsers)))
                {
                    enabledUsers.Add(user);
                }
            }

            return enabledUsers;
        }

        public static bool IsMember(User user, List<DirectoryObject> group)
        {
            bool result = false;
            foreach (var member in group)
            {
                if (user.Id == member.Id) { result = true; }
            }
            return result;
        }

        public static List<DirectoryObject> ParseGroup(String groupID)
        {
            GraphServiceClient thisClient = Authentication.GetAuthenticatedClient();

            List<DirectoryObject> result = new List<DirectoryObject>();
            List<DirectoryObject> request = new List<DirectoryObject>();
            var root = thisClient.Groups[groupID].Members.Request().GetAsync().Result;

            // build full member list of member objects
            request.AddRange(root.CurrentPage);
            while (root.NextPageRequest != null)
            {
                root = root.NextPageRequest.GetAsync().Result;
                request.AddRange(root.CurrentPage);
            }

            foreach (var dirObject in request)
            {
                if (dirObject.ODataType == "#microsoft.graph.group")
                {
                    result.AddRange(ParseGroup(dirObject.Id));
                }
                else
                {
                    result.Add(dirObject);
                }
            }
            return result;
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
