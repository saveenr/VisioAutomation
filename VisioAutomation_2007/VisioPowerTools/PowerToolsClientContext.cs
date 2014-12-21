using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VisioPowerTools
{
    public class PowerToolsClientContext : VisioAutomation.Scripting.Context
    {
        public delegate void WriteString(string s);

        public event WriteString OnWriteString;

        public PowerToolsClientContext()
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

        public override void WriteWarning(string s)
        {
            string msg = string.Format("Warning: {0}", s);
            this.DefaultWriteString(msg);
        }


        public override void WriteUser(string s)
        {
            this.DefaultWriteString(s);
        }

        public override void WriteVerbose(string s)
        {
           this.DefaultWriteString(s);
        }

        public void DefaultWriteString(string s)
        {
            if (this.OnWriteString != null)
            {
                this.OnWriteString(s);
            }
        }
    }
}
