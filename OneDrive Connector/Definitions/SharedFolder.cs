using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OneDrive_Connector.Definitions
{
    class SharedFolder
    {
        public String FolderID { get; }
        public String FolderLink { get; }
        public String Owner { get; }
        public List<Microsoft.Graph.Permission> SharedWith { get; }

        public SharedFolder(String folder, String link, String from)
        {
            FolderID = folder;
            FolderLink = link;
            Owner = from;
            SharedWith = new List<Microsoft.Graph.Permission>();
        }

        public void AddPermission(Microsoft.Graph.Permission input)
        {
            SharedWith.Add(input);
        }
    }
}
