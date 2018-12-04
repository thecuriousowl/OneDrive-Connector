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

        public static List<String> exclusion = new List<String>()
        {
            
        };

        static void Main(string[] args)
        {
            GraphServiceClient graphClient = Authentication.GetAuthenticatedClient();



            // Get all shared items
            // DriveExplorer explore = new DriveExplorer(graphClient);
            //List<String> exclusion = new List<String>();
            //List<User> enabledUsers = GraphHelpers.GetAllEnabledUsers(exclusion);

            //CheckForNameChange(enabledUsers, graphClient);

            //var sharedStuff = explore.reportSharedFolders(test);

            UpdatePermissions.RunUpdate("691f100e-8565-481b-a99b-1fcd4e85eee4");
            //UpdatePermissions.RunUpdate("477b1109-17b1-40d3-8993-0ba66134fa20");


        }

        public static String UpdatePermissions2(String filePath, GraphServiceClient graphClient)
        {
            // Gather all shareditem meta data in a string array variable
            // input is file path for current chunk
            var list = System.IO.File.ReadAllLines(filePath);

            Queue<String> work = new Queue<string>(list);
            List<String> updated = new List<String>();

            Console.WriteLine("Total Permissions to Recreate is " + work.Count);
            // Loop through queue until it is empty
            while(work.Count > 0)
            {
                bool upnChanged = false;
                bool DoesNotExist = false;

                // take one line of input from queue
                var temp = work.Dequeue();

                var split = temp.Split(';');
                var username = split[0];
                var userid = split[1];
                var folderid = split[2];
                var permissionid = split[3];
                var grantedto = split[4];

                Console.WriteLine(username);
                Console.WriteLine(userid);
                Console.WriteLine(folderid);
                Console.WriteLine(permissionid);
                Console.WriteLine(grantedto);

                if (updated.Contains(userid))
                {
                    upnChanged = true;
                }
                else
                {
                    var rootCheck = graphClient.Users[userid].Drive.Root.Request().GetAsync().Result;
                    if (rootCheck is null) { DoesNotExist = true; }
                    else
                    {
                        if (rootCheck.WebUrl.Contains("teneo_com"))
                        {
                            upnChanged = true;
                            updated.Add(userid);
                        }
                        else
                        {
                            upnChanged = false;
                        }
                    }
                }

                if (upnChanged)
                {
                    try
                    {
                        // upn has been updated, recreate permissions
                        // graphClient.Users[userid].Drive.Items[folderid].Permissions[permissionid].Request().DeleteAsync().
                        var currentPermission = graphClient.Users[userid].Drive.Items[folderid].Permissions[permissionid].Request().GetAsync().Result;

                        // Run Delete and then await success
                        Console.WriteLine("Deleting Permission. . ." + permissionid);
                        var deleteTask = graphClient.Users[userid].Drive.Items[folderid].Permissions[permissionid].Request().DeleteAsync();
                        while (deleteTask.IsCompleted != true)
                        {
                            Thread.Sleep(100);
                        }
                        Console.WriteLine("Deleted");

                        if (deleteTask.IsFaulted)
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
                            var createTask = graphClient.Users[userid].Drive.Items[folderid].Invite(invitees, true, new List<String>() { "write" }, true, "Teneo rebrand permission re-established").Request().PostAsync();
                            while (createTask.IsCompleted != true)
                            {
                                if (createTask.IsFaulted) { Console.WriteLine("Error in creating invite"); }
                            }
                            if ((createTask.Result).Count > 0) { Console.WriteLine("Invite Sent for " + folderid); }
                        }
                    }
                    catch
                    {
                        Console.WriteLine("This Permission Errored Out.");
                        DoesNotExist = false;
                    }
                }
                else if(DoesNotExist) { }
                else { work.Enqueue(temp); }
                Thread.Sleep(50);
            }
            return null;
        }
        
        public static void CheckForNameChange(List<User> users, GraphServiceClient graphClient)
        {
            Queue<User> work = new Queue<User>();
        
            foreach (var user in users)
            {
                work.Enqueue(user);
            }

            int wait = 0;

            while (work.Count > 0)
            {
                wait++;
                Thread.Sleep(25);

                var temp = work.Dequeue();
                /*
                User teneoUser;
                try
                {
                    teneoUser = graphClient.Users[temp.Id].Request().GetAsync().Result;
                }
                catch { teneoUser = null; }*/
                if (temp.Mail != null)
                {
                    if (temp.Mail.Contains("teneo.com"))
                    {
                        Console.WriteLine(temp.DisplayName + " has been updated, sending mail.");
                        String html = System.IO.File.ReadAllText("C:/Rebrand/Misc/Notification.html");
                        html = html.Replace("TENEOAZUREFIRSTNAME",temp.GivenName);
                        // Send mail of MySite reflecting UPN change
                        Message notification = MailHelper.ComposeMail("Your Teneo Email Has Been Updated",
                            html,
                            new List<string>() { temp.Mail });
                        graphClient.Me.SendMail(notification, false).Request().PostAsync();

                        // Changing OneDrive Permissions
                        //var result = UpdatePermissions("C:/Rebrand/Teneo_PEM.txt", graphClient, temp.Id);

                        //if (result != null) { OneDrive.Enqueue(result); }

                        // Send asynchronous call to permission rewrite for this user.
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

               /* List<String> moreToDo = new List<String>();
                if (wait == 25)
                {
                    for(int x = 0; x < OneDrive.Count;x++)
                    {
                        var moreWork = OneDrive.Dequeue();
                        var stillMore = UpdatePermissions("C:/Rebrand/Teneo_PEM.txt", graphClient, moreWork);

                        if (stillMore != null) { moreToDo.Add(stillMore); }
                    }
                }

                OneDrive = new Queue<string>(moreToDo);*/
            }
        }
    }
}
