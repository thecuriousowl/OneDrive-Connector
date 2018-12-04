using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace OneDrive_Connector.Controllers
{
    

    class UpdatePermissions
    {
        public static List<String> exclude = new List<String>()
        {
            "ae6b7114-cddc-4a13-b6ff-ea499e276859",
            "bdc66744-96cf-4da1-a412-94d263209420",
            "a467b40f-f36c-4f95-a341-1706eb7cb3c8",
            "e3596a16-3217-4a3f-8a34-f69dc387ed75",
            "f7d84154-ebb2-48b1-a6ef-50382c4206ee",
            "098cc422-53d2-4caf-a0bd-0e88b2e0858c",
            "9a1693a9-8205-4bdf-bde7-54f273d3c69c",
            "19301ad0-f64c-403f-b489-c0969bda83e2",
            "d6823e28-1d5e-4d77-85a1-b3fb371c208a",
            "bdb9891c-686a-400c-aacc-1916a98811eb",
            "b6e036bb-8e47-4b75-a005-53ce6131e802",
            "568c3118-5845-4fae-8bad-e4a9294730ec",
            "b6dd8420-5aa3-4331-8918-be622cf840f2",
            "a86f6ea0-be45-4728-8a24-025929e36819",
            "fca72927-04df-444e-b3b3-85baeb0f503a",
            "94b9fc42-3a90-4a1b-baba-e68a3701d30c",
            "d8ea23c4-1905-4364-bdc4-5dc207cfe8be",
            "eca3d73d-308e-4eb0-8d7f-2aa1cdbb6992",
            "9a261093-4c1b-4b46-8a16-4d5af43bf6e6",
            "d82e3608-4c3a-4a4d-b527-ea4c9c443705",
            "89925003-b1ca-414b-a374-575313f0ae22",
            "ff5ba7a8-e7d7-40f6-9d96-f317e176989d",
            "55912c25-8061-40b8-801b-924f00b8e2aa",
            "8ea77faa-427b-45e0-bb40-861f3de7e71c",
            "90870f26-1c62-4f78-99f7-3420084cc125",
            "be9522ef-5a65-4cc4-9d24-ce9e7640dd33",
            "85bbb84e-99e3-48a9-b82b-7830ea277446",
            "27ecb94f-1355-4858-8e7e-4cdd83e5d463",
            "3de2d4ac-8931-483c-abb6-b2d5f06bde2d",
            "4a98fc1f-33bc-40e1-bb81-4c768248f806",
            "4333578f-ae33-4a56-8d72-7e00d911dd45",
            "1b4c2287-365d-4eca-b305-78946ef9fa63",
            "a0858f18-5bf0-4710-b02b-d1d073bfb724",
            "6942e464-84cb-4db1-ac17-e2eb3e997b5f",
            "cce26b10-ad9f-4d4c-856e-86202d75972b"
        };

        public static void RunUpdate(String groupID)
        {
            var graphClient = Authentication.GetAuthenticatedClient();
            var usersToUpdate = ParseGroup(groupID);

            foreach(var user in usersToUpdate)
            {
                if (!(exclude.Contains(user.Id)))
                {
                    Console.WriteLine("Working on " + (graphClient.Users[user.Id].Request().GetAsync().Result).DisplayName);
                    var id = user.Id;
                    DriveItem usersDrive = null;
                    try
                    {
                        usersDrive = graphClient.Users[id].Drive.Root.Request().GetAsync().Result;
                    }
                    catch { Console.WriteLine("Error Retrieving User"); }

                    if (usersDrive != null)
                    {
                        List<String> workedOnThis = new List<String>() { id };
                        // Parse Drive and recreate permissions as we go
                        System.IO.File.AppendAllLines("C:/temp/permissionUpdateProgress.txt", workedOnThis);
                        parseFolders(id, usersDrive, graphClient);
                    }
                    else
                    {
                        Console.Write("User OneDrive Does not Exist");
                    }
                }
            }

        }

        public static void parseFolders(String user, DriveItem thisNode, GraphServiceClient graphClient)
        {
            var children = graphClient.Users[user].Drive.Items[thisNode.Id].Children.Request().GetAsync().Result;
            foreach(var child in children)
            {
                if(child.Folder != null) // if this is a folder
                {
                    Console.WriteLine("\tChecking folder " + child.Name + " owned by: " + user);
                    // Check if shared
                    if(child.Shared != null)
                    {
                        Console.WriteLine("\t\tfolder is shared. . ");
                        // update permissions
                        var pemList = graphClient.Users[user].Drive.Items[child.Id].Permissions.Request().GetAsync().Result;
                        Console.WriteLine("\t\t\t" + pemList.Count + " permissions found.");
                        foreach(var pem in pemList)
                        {
                            if (pem.GrantedTo != null)
                            {
                                if (pem.GrantedTo.User != null)
                                {
                                    if (pem.GrantedTo.User.Id != null)
                                    {
                                        Console.WriteLine("\t\t\t\tUpdating a Permission.");
                                        try
                                        {
                                            updatePermissions(user, child.Id, pem, graphClient);
                                        }
                                        catch
                                        {
                                            Console.WriteLine("Permission Update error");
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        // Search deeper
                        parseFolders(user, child, graphClient);
                    }
                }
            }
        }

        public static void updatePermissions(String userid, String folderid, Permission pem, GraphServiceClient graphClient)
        {
            var permissionid = pem.Id;
            var grantedto = pem.GrantedTo.User.Id;
            // upn has been updated, recreate permissions
            // graphClient.Users[userid].Drive.Items[folderid].Permissions[permissionid].Request().DeleteAsync().
            var currentPermission = graphClient.Users[userid].Drive.Items[folderid].Permissions[permissionid].Request().GetAsync().Result;

            // Run Delete and then await success
            Console.WriteLine("Deleting Permission. . ." + permissionid);
            var deleteTask = graphClient.Users[userid].Drive.Items[folderid].Permissions[permissionid].Request().DeleteAsync();

            // timeout error variables
            int errorCount = 0;
            bool apiError = false;
            while (deleteTask.IsCompleted != true && apiError == false)
            {
                if (errorCount == 600) { apiError = true; }
                Thread.Sleep(100);
                errorCount++;
            }
            Console.WriteLine("Deleted");

            if (deleteTask.IsFaulted || apiError == true)
            {
                Console.WriteLine("Failed to Delete this Permission");
            }
            else
            {
                List<DriveRecipient> invitees = new List<DriveRecipient>()
                        {
                            new DriveRecipient()
                            {
                                Email = (graphClient.Users[grantedto].Request().GetAsync().Result).Mail
                            }
                        };
                Console.WriteLine("Creating new permission. . .");
                var createTask = graphClient.Users[userid].Drive.Items[folderid].Invite(invitees, true, new List<String>() { "write" }, true, "Teneo rebrand permission re-established").Request().PostAsync();
                while (createTask.IsCompleted != true)
                {
                    if (createTask.IsFaulted) { Console.WriteLine("Error in creating invite"); }
                }
                if ((createTask.Result).Count > 0) { Console.WriteLine("Invite Sent for " + folderid); }
            }
        }

        public static List<DirectoryObject> ParseGroup(String groupID)
        {
            Console.WriteLine("Searching this Group");
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
                    Console.WriteLine(dirObject.Id);
                    result.Add(dirObject);
                }
            }
            return result;
        }
    }
}
