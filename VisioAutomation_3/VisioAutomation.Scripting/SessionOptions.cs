namespace VisioAutomation.Scripting
{
    public class SessionOptions
    {
        public virtual void WriteDebug(string s)
        {
            this.DefaultWriteString(s);
        }

        public virtual void WriteUser(string s)
        {
            this.DefaultWriteString(s);
        }

        public virtual void WriteError(string s)
        {
            this.DefaultWriteString(s);
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