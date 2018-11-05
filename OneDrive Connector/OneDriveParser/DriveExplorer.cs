using Microsoft.Graph;
using OneDrive_Connector.Controllers;
using OneDrive_Connector.Definitions;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace OneDrive_Connector.OneDriveParser
{
    class DriveExplorer
    {
        public static String path = "C:/temp/OneDrive_Sharing_Report.txt";
        public static System.IO.StreamWriter file;
        private GraphServiceClient graphClient;

        // User exclusion group
        private List<String> checkList = new List<String>(){
               "99b7d0b1-f6a0-460d-bf4e-f654da45ca5d",
               "740021a3-dbf0-4145-954f-878b8cf24ff8",
               "649fd178-dfd7-4c5d-9ef1-168aa5fcf291",
               "350e0523-1195-464e-ba36-cdffadb21cc6",
               "83d46a59-95ff-485b-bf67-6d7844e0929c",
               "3ca8debc-2445-40c6-98a0-a7dc658ca704",
               "2ea9c413-555e-40a6-ba26-7fdaaee1f2f0",
               "0662e0c8-b788-4e5d-bac2-c24ccb21b8f7",
               "002cc7fb-1970-4457-8bbd-0c49dc914cfb",
               "145e0420-d562-42ce-aad6-6696c8567e42",
               "d1223ac0-ff7e-45b0-9bb5-8c1f6446a962",
               "2d110c9a-5a61-4317-93ad-69a6853ccad1",
               ""
        };

        // Constructor Method
        public DriveExplorer(GraphServiceClient client)
        {
            graphClient = client; // retrieve OAuth 2.0 token and instantiate a graph services client
        }

        // Exposed Method for requesting a list of shared folders given a list of users
        public List<UsersSharedItems> reportSharedFolders(List<Microsoft.Graph.User> usersToCheck)
        {
            List<UsersSharedItems> reportResult = new List<UsersSharedItems>();
            file = new System.IO.StreamWriter(path);
            file.AutoFlush = true;
            foreach (var user in usersToCheck)
            {
                Console.WriteLine("Working on: " + user.DisplayName + " . . . ");
                //
                // Check to see if users are enabled (to limit unnecessary API calls)
                //

                // locate OneDrive Root Folder, if it doesn't exist, skip
                try // Root found path
                {
                    DriveItem root = graphClient.Users[user.Id].Drive.Root.Request().GetAsync().Result;
                    Console.WriteLine(root.WebUrl);

                    UsersSharedItems currentUser = new UsersSharedItems(user.DisplayName, user.Id); // Create UserSharedItems object to store a list of SharedFolder items

                    // Recurse from root and return list of folders that are shared 
                    // Task of async result is store in a temp variable to be opened and then appended to the result object
                    var temp = recurseFolders(user , root);
                    if(temp.Count > 0)
                    {
                        currentUser.addSharedItems(temp);
                        reportResult.Add(currentUser);
                    }
                }
                catch // exception means the root was unreachable/does not exist
                {
                    Console.WriteLine("This User's OneDrive does not exist.");
                }
            }

            // List of Users with shared items should be filled after prior loop. Return the result
            file.Close();
            return reportResult;
        }

        private List<SharedFolder> recurseFolders(Microsoft.Graph.User user , DriveItem folder)
        {
            List<SharedFolder> partialResult = new List<SharedFolder>(); // return case variable 
            var childItems = graphClient.Users[user.Id].Drive.Items[folder.Id].Children.Request().GetAsync().Result; // list child items of inbound folder
            System.Threading.Thread.Sleep(50); // spaces out API calls

            foreach (var child in childItems)
            {
                if (child.Folder != null) // Check if item is a folder
                {
                    Console.WriteLine(" . . . . Folder Debug " + child.Name); 
                    if(child.Shared != null) // Check if folder is shared, if it is shared, check permissions and stop recursing
                    {
                        Console.Write(" (SHARED) ");
                        // Create folder object to store permissions
                        SharedFolder temp = new SharedFolder(child.Id, child.WebUrl, user.Id);

                        // Request Permissions on folder
                        var permissions = graphClient.Users[user.Id].Drive.Items[child.Id].Permissions.Request().GetAsync().Result;

                        // Loop Through all Permissions
                        int count = 0;
                        foreach(var permission in permissions)
                        {
                            String grantedTo = null;
                            if(permission.GrantedTo != null) // GrantedTo refers to an in tenant object that this folder is shared with
                            {
                                grantedTo = permission.GrantedTo.User.Id;
                                if(grantedTo != user.Id && exclusionCheck(grantedTo) && grantedTo != null)
                                {
                                    // permission is found
                                    temp.SharedWith.Add(permission);

                                    // output
                                    file.WriteLine(user.DisplayName + ";" + user.Id + ";" + child.Id + ";" + permission.Id + ";" + grantedTo);
                                    System.Console.WriteLine(user.DisplayName + ";" + user.Id + ";" + child.Id + ";" + permission.Id + ";" + grantedTo);
                                    file.Flush();
                                    count++;
                                }
                            }
                            else // exception case indicates an external link share
                            {
                                temp.SharedWith.Add(permission);
                            }
                        }

                        if (count > 0) { partialResult.Add(temp); }

                    }
                    else // if folder is not shared, go down one level
                    {
                        partialResult.AddRange(recurseFolders(user, child));
                    }
                }
            }
            return partialResult;
        }















        


        // V2.0
        // Method Entry Point
        public List<UsersSharedItems> FindSharedItems(List<User> userList)
        {
            // Handle Primary Loop
            foreach(var user in userList)
            {
                
            }
            return null;
        }

        private List<SharedFolder> traverseFolders(DriveItem root)
        {
            List<SharedFolder> tempResult = new List<SharedFolder>();
            var children = root.Children;
            
            foreach(var child in children)
            {
                if(child.Folder != null)
                {

                }
            }
            return null;
        }


        //Helper Functions
        private SharedFolder isShared(User thisUser, DriveItem thisFolder)
        {
            if(thisFolder.Shared != null)
            {
                var temp = new SharedFolder(thisFolder.Id, thisFolder.WebUrl, thisUser.Id);
                var permissions = graphClient.Users[thisUser.Id].Drive.Items[thisFolder.Id].Permissions.Request().GetAsync().Result;

                foreach(var permission in permissions)
                {
                    // link test
                    if (permission.GrantedTo != null)
                    {
                        String assignedTo = permission.GrantedTo.User.Id;
                        if (assignedTo != thisUser.Id && exclusionCheck(assignedTo) && assignedTo != null)
                        {
                            temp.AddPermission(permission);
                        }
                    }
                    else
                    {
                        // is link
                    }
                }

                return temp;
            }
            else { return null; }
        }

        private Boolean exclusionCheck(String grantedTo)
        {
            if (checkList.Contains(grantedTo)) return false;
            else return true;
        }





















    }
}
