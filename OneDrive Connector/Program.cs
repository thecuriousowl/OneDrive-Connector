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

        public static List<String> exclusion = new List<String>()
        {
            
        };

        static void Main(string[] args)
        {
            GraphServiceClient graphClient = Authentication.GetAuthenticatedClient();

            string html = System.IO.File.ReadAllText("C:/Users/thecu/Desktop/System Notification.htm");
            Message test = MailHelper.ComposeMail("HTML TEST", html, new List<String>() { "philip.pan@teneoholdings.com" });
            var response = graphClient.Me.SendMail(test, null).Request().PostAsync();
            while (response.IsCompleted != true) { }
            if (response.IsFaulted) { Console.WriteLine(response.Exception.InnerException.Message); }
            else { Console.WriteLine("Mail was sent"); }

            // Get all shared items
            // DriveExplorer explore = new DriveExplorer(graphClient);
            // List <User> enabledUsers = GraphHelpers.GetAllEnabledUsers(exclusion)

            //var sharedStuff = explore.reportSharedFolders(test);

            //updatePermissions("C:/temp/OneDrive_Sharing_Report.txt", graphClient);

            
        }

        public static void UpdatePermissions(String filePath, GraphServiceClient graphClient)
        {
            // Gather all shareditem meta data in a string array variable
            // input is file path for current chunk
            var list = System.IO.File.ReadAllLines(filePath);

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

                if (upnChanged)
                {
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
                        var createTask = graphClient.Users[userid].Drive.Items[folderid].Invite(invitees,true,new List<String>() { "write" }, true, "Teneo Global rebrand permission re-established").Request().PostAsync();
                        while(createTask.IsCompleted != true)
                        {
                            if (createTask.IsFaulted) { Console.WriteLine("Error in creating invite"); }
                        }
                        if((createTask.Result).Count > 0) { Console.WriteLine("Invite Sent for " + folderid); }
                    }
                }
                else { work.Enqueue(temp); }
                Thread.Sleep(50);
            }
        }
        
        public static void CheckForNameChange(List<User> users, GraphServiceClient graphClient)
        {
            Queue<User> work = new Queue<User>();
            foreach (var user in users)
            {
                work.Enqueue(user);
            }

            while (work.Count > 0)
            {
                Thread.Sleep(25);

                var temp = work.Dequeue();
                Drive root;
                try
                {
                    root = graphClient.Users[temp.Id].Drive.Request().GetAsync().Result;
                }
                catch { root = null; }
                if (root != null)
                {
                    if (root.WebUrl.Contains("teneoglobal"))
                    {
                        Console.WriteLine(temp.DisplayName + " has been updated, sending mail.");
                        // Send mail of MySite reflecting UPN change
                        Message notification = MailHelper.ComposeMail("UPN Updated",
                            ("MySite URL is updated for " + temp.DisplayName + " at " + root.WebUrl),
                            new List<string>() { "Christian.Mariano@TeneoHoldings.com" });
                        graphClient.Me.SendMail(notification, false).Request().PostAsync();
                    }
                    else
                    {
                        Console.WriteLine("\t" + temp.DisplayName + " has not changed yet. . . ");
                        // update user object
                        temp = graphClient.Users[temp.Id].Request().GetAsync().Result;
                        Console.WriteLine("\tupdated user object and re-queueing.");
                        work.Enqueue(temp);
                    }
                }
            }
        }
    }
}
