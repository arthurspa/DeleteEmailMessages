using MailKit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DeleteEmailMessages
{
    public class FolderListItem
    {
        public bool IsSelected { get; set; }

        public string FolderName { get; set; }

        public int TotalFilteredMessages { get; set; }

        public IList<UniqueId> MessageUniqueIds { get; set; }
    }
}
