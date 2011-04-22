using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VisioPowerTools
{
    public class PowerToolsSessionOptions : VisioAutomation.Scripting.SessionOptions
    {
        public PowerToolsSessionOptions()
        {

        }

        public override void WriteDebug(string s)
        {
            string msg = string.Format("DEBUG: {0}", s);
            this.DefaultWriteString(msg);
        }

        public override void WriteError(string s)
        {
            string msg = string.Format("ERROR: {0}", s);
            this.DefaultWriteString(msg);
        }

        public override void WriteUser(string s)
        {
            this.DefaultWriteString(s);
        }

        public override void WriteVerbose(string s)
        {
            System.Console.WriteLine(s);
        }

        public override void DefaultWriteString(string s)
        {
            System.Console.WriteLine(s);
        }
    }

}
