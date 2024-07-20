using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FilterAddin
{
    public class Settings
    {
        public List<string> blacklist { get; set; }
        public List<string> whitelist { get; set; }
        public string blacklistWarning { get; set; }
        public string unkownSenderWarning { get; set; }

    }
}
