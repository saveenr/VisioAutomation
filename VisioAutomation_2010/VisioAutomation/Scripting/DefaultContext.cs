using System;

namespace VisioAutomation.Scripting
{
    public class DefaultContext : Context
    {
        public override void WriteDebug(string s)
        {
            string msg = String.Format("DEBUG: {0}", s);
            this.DefaultWriteString(msg);
        }

        public override void WriteUser(string s)
        {
            this.DefaultWriteString(s);
        }

        public override void WriteError(string s)
        {
            string msg = String.Format("ERROR: {0}", s);
            this.DefaultWriteString(msg);
        }

        public override void WriteVerbose(string s)
        {
            string msg = String.Format("VERBOSE: {0}", s);
            this.DefaultWriteString(msg);
        }

        public override void WriteWarning(string s)
        {
            string msg = String.Format("WARNING: {0}", s);
            this.DefaultWriteString(msg);
        }

        public virtual void DefaultWriteString(string s)
        {
            System.Console.WriteLine(s);
        }
    }
}