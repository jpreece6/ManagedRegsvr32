using CommandLine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ManagedRegsvr32
{
    internal class Options
    {
        [Option('s', "Silent", DefaultValue = false, HelpText = "Supresses Messagebox Output", MetaValue = "Silent", Required = false)]
        public bool Silent { get; set; }

        [Option('u', "UnReg", DefaultValue = false, HelpText = "Unregisters a COM dll from the system", MetaValue = "Unreg", Required = false)]
        public bool UnRegister { get; set; }

        [Option('p', "Path", HelpText = "Path to the COM dll", Required = true)]
        public string DllPath { get; set; }

        [Option('c', "Console", HelpText = "Show console window with command output.", MetaValue = "Console", Required = false)]
        public bool ShowConsole { get; set; }
    }
}
