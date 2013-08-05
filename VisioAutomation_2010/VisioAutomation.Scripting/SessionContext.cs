namespace VisioAutomation.Scripting
{
    public class SessionContext
    {
        // this class is for storing additional data
        // about a session and handling I/O
        // for example, if you want to use a Scripting Session
        // and you want to handle all I/O (to log it or to send it
        // to special outputs) then derive from this class and
        // set the Session.Context property

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

        public virtual void WriteWarning(string s)
        {
            this.DefaultWriteString(s);
        }


        public virtual void DefaultWriteString(string s)
        {
            System.Console.WriteLine(s);
        }        
    }
}