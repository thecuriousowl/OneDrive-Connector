using Microsoft.Graph;
using OneDrive_Connector.Controllers;
using OneDrive_Connector.OneDriveParser;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OneDrive_Connector
{
    class Program
    {
        // Environment Variables
        public static ManualResetEvent resetEvent = new ManualResetEvent(false); // Set Manual await event for asynchronous calls
        public static String path = "C:/temp/OneDrive_Sharing_Report.txt";
        public static EventWaitHandle ewh;




        static void Main(string[] args)
        {
            ewh = new EventWaitHandle(false, EventResetMode.ManualReset);

            GraphServiceClient graphClient = Authentication.GetAuthenticatedClient();

            // Get all shared items
            // DriveExplorer explore = new DriveExplorer(graphClient);

            /*
            List<User> teneoAll = new List<User>();
            // Get full licensed user list 
            int count = 1;
            System.Console.WriteLine("Requesting Page " + count);
            var users = graphClient.Users.Request().Top(200).Select("DisplayName,Id,AssignedLicenses").GetAsync().Result;
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
                if (licensing.Count > 0)
                {
                    enabledUsers.Add(user);
                }
            }
            */

            List<User> test = new List<User>() { graphClient.Users["ac6592a5-7466-4baf-ba90-7ea50042c8e7"].Request().Select("DisplayName,Id,AssignedLicenses").GetAsync().Result };

            //var sharedStuff = explore.reportSharedFolders(test);

            //updatePermissions("C:/temp/OneDrive_Sharing_Report.txt", graphClient);

            Recipient aPerson = new Recipient();
            EmailAddress address = new EmailAddress();
            address.Address = "Christian.Mariano@teneoholdings.com";
            aPerson.EmailAddress = address;

            Message mail = new Message()
            {
                Subject = "Test Send",
                Body = new ItemBody()
                {
                    ContentType = BodyType.Text,
                    Content = "Hello World"
                },
                ToRecipients = new List<Recipient>()
                {
                    aPerson
                }
            };


            var testEmail = graphClient.Users["ac6592a5-7466-4baf-ba90-7ea50042c8e7"].SendMail(mail, false).Request().PostAsync();
            while(testEmail.IsCompleted != true)
            {

            }


            // get all shared items


            /*
            var testCheck = new List<String>();
            testCheck.Add("ac6592a5-7466-4baf-ba90-7ea50042c8e7");

            List<Microsoft.Graph.User> users = new List<User>();
            foreach(var id in testCheck)
            {
                var temp = graphClient.Users[id].Request().GetAsync().Result;
                users.Add(temp);
            }

            List<User> teneoDirectory = new List<User>();


            var shared = explore.reportSharedFolders(users);

            
            var host = shared[0];
            var toShare = host.Shared[0];

            var recipients = new List<DriveRecipient>()
            {
                new DriveRecipient()
                {
                    Email = "christian.mariano@teneoglobal.com"
                }
            };
            var roles = new List<String>()
            {
                "write"
            };

            foreach(var folder in host.Shared)
            {
                // POST Example
                // var something = graphClient.Users[host.ID].Drive.Items[folder.FolderID].Invite(recipients, true, roles, true, null).Request().PostAsync();
            }
            */
        }

        public static void updatePermissions(String input, GraphServiceClient graphClient)
        {
            // Gather all shareditem meta data in a string array variable
            // input is file path
            var list = System.IO.File.ReadAllLines(input);
            foreach(var line in list)
            {
                Console.WriteLine(line);
            }

            Queue<String> work = new Queue<string>(list);
            List<String> upnUpdated = new List<String>();
            
            // Loop through queue until it is empty
            while(work.Count > 0)
            {
                Console.WriteLine("Total Permissions to Recreate is " + work.Count);
                // take one line of input from queue
                var temp = work.Dequeue();

                var split = temp.Split(';');
                var username = split[0];
                var userid = split[1];
                var folderid = split[2];
                var permissionid= split[3];
                var grantedto = split[4];

                bool upnChanged = false;

                // check user onedrive url for global
                if(upnUpdated.Contains(userid))
                {
                    upnChanged = true;
                }
                else
                {
                    var rootCheck = graphClient.Users[userid].Drive.Root.Request().GetAsync().Result;
                    if (rootCheck.WebUrl.Contains("teneoglobal"))
                    {
                        upnChanged = true;
                        upnUpdated.Add(userid);
                    }
                    else { upnChanged = false; }
                }

                if (true)
                {
                    var awaitDelete = new EventWaitHandle(false, EventResetMode.ManualReset);

                    // upn has been updated, recreate permissions
                    // graphClient.Users[userid].Drive.Items[folderid].Permissions[permissionid].Request().DeleteAsync().
                    var currentPermission = graphClient.Users[userid].Drive.Items[folderid].Permissions[permissionid].Request().GetAsync().Result;

                    // Run Delete and then await success
                    Console.WriteLine("Deleting Permission. . ." + permissionid);
                    var deleteTask = graphClient.Users[userid].Drive.Items[folderid].Permissions[permissionid].Request().DeleteAsync();
                    while(deleteTask.IsCompleted != true)
                    {
                        Thread.Sleep(100);
                    }
                    Console.WriteLine("Deleted");

                    if(deleteTask.IsFaulted)
                    {
                        Console.WriteLine("Failed to Delete this Permission");
                        work.Enqueue(temp);
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
                        var createTask = graphClient.Users[userid].Drive.Items[folderid].Invite(invitees,true,new List<String>() { "write" }, true, null).Request().PostAsync();
                        while(createTask.IsCompleted != true)
                        {
                            Console.WriteLine("waiting for success");
                            if (createTask.IsFaulted) { Console.WriteLine("Error in creating invite"); }
                        }
                        if((createTask.Result).Count > 0) { Console.WriteLine("Invite Sent!"); }
                    }
                }
                else { work.Enqueue(temp); }
                Thread.Sleep(50);
            }
        }
        
    }
}
