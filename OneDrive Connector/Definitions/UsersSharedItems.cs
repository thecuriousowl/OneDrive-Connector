using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OneDrive_Connector.Definitions
{
    class UsersSharedItems
    {
        public String Name { get; }
        public String ID { get; }
        public List<SharedFolder> Shared { get;set;}

        public UsersSharedItems(String DisplayName, String userID)
        {
            Name = DisplayName;
            ID = userID;
            Shared = new List<SharedFolder>();
        }

        public void addSharedItems (List<SharedFolder> input)
        {
            Shared.AddRange(input);
        }
    }
}
