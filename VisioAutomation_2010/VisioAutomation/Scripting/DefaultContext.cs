namespace VisioAutomation.Scripting
{
    public class DefaultContext : Context
    {
        public override void WriteDebug(string s)
        {
            string msg = $"DEBUG: {s}";
            this.DefaultWriteString(msg);
        }

        public override void WriteUser(string s)
        {
            this.DefaultWriteString(s);
        }

        public override void WriteError(string s)
        {
            string msg = $"ERROR: {s}";
            this.DefaultWriteString(msg);
        }

        public override void WriteVerbose(string s)
        {
            string msg = $"VERBOSE: {s}";
            this.DefaultWriteString(msg);
        }

        public override void WriteWarning(string s)
        {
            string msg = $"WARNING: {s}";
            this.DefaultWriteString(msg);
        }

        public virtual void DefaultWriteString(string s)
        {
            System.Console.WriteLine(s);
        }
    }
}