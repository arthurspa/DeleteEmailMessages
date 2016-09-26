using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DeleteEmailMessages
{
    public class AppConfiguration
    {
        public readonly string HostKeyName = "Host";
        public readonly string PortKeyName = "Port";
        public readonly string UseSslKeyName = "UseSsl";
        public readonly string UsernameKeyName = "Username";

        public string Host { get; set; }
        public int Port { get; set; }
        public bool UseSsl { get; set; }
        public string Username { get; set; }

    }
}
