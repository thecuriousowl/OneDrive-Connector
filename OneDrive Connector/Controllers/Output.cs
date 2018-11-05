using OneDrive_Connector.Definitions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OneDrive_Connector.Controllers
{
    class Output
    {
        public static void report(List<UsersSharedItems> input)
        {
            // Loop through users with shared items

            foreach(var user in input)
            {
                // Report User Details

                // Loop through user folders
                foreach(var folder in user.Shared)
                {
                    // Report Folder Details

                    // Loop through permission objects on shared folder
                    foreach(var permission in folder.SharedWith)
                    {
                        // List Permission Details
                        // Permissions deatsildwasdcdqawda
                    }
                }
            }

        }
    }

    
}
