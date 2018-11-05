using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OneDrive_Connector.Controllers
{
    class SyncCheck
    {
        // check sync status of user UPN in Azure as well as MySite URL change
        public static Boolean checkUPNsync(String inputID, GraphServiceClient graphClient)
        {
            var result = graphClient.Users[inputID].Drive.Root.Request().GetAsync().Result;
            var url = result.WebUrl;

            if(url.Contains("teneoglobal"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        
    }
}
