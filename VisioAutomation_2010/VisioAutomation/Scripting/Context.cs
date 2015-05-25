namespace VisioAutomation.Scripting
{
    public abstract class Context
    {
        // this class is for storing additional data
        // about a session and handling I/O
        //
        // for example, if you want to use a Scripting Session
        // and you want to handle all I/O (to log it or to send it
        // to special outputs) then derive from this class and
        // set the Session.Context property

        public abstract void WriteDebug(string s);
        public abstract void WriteUser(string s);
        public abstract void WriteError(string s);
        public abstract void WriteVerbose(string s);
        public abstract void WriteWarning(string s);
    }
}