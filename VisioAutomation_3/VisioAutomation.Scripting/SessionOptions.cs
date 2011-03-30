namespace VisioAutomation.Scripting
{
    public class SessionOptions
    {
        public virtual void WriteDebug(string s)
        {
            string msg = string.Format("DEBUG: {0}", s);
            this.DefaultWriteString(msg);
        }

        public virtual void WriteUser(string s)
        {
            this.DefaultWriteString(s);
        }

        public virtual void WriteError(string s)
        {
            string msg = string.Format("ERROR: {0}", s);
            this.DefaultWriteString(msg);
        }

        public virtual void WriteVerbose(string s)
        {
            this.DefaultWriteString(s);
        }

        public virtual void DefaultWriteString(string s)
        {
            System.Console.WriteLine(s);
        }        
    }
}