namespace VisioAutomation.Scripting.Commands
{
    public class OutputCommands : CommandSet
    {
        internal OutputCommands(Client client) :
            base(client)
        {

        }

        public void WriteUser(string s)
        {
            this.Client.WriteUser(s);
        }

        public void WriteDebug(string s)
        {
            this.Client.WriteDebug(s);
        }

        public void WriteError(string s)
        {
            this.Client.WriteError(s);
        }

        public void WriteVerbose(string s)
        {
            this.Client.WriteVerbose(s);
        }

        public void WriteUser(string fmt, params object[] items)
        {
            this.Client.WriteUser(fmt, items);
        }

        public void WriteDebug(string fmt, params object[] items)
        {
            this.Client.WriteDebug( fmt, items);           
        }

        public void WriteError(string fmt, params object[] items)
        {
            this.Client.WriteError( fmt, items);
        }

        public void WriteVerbose(string fmt, params object[] items)
        {
            this.Client.WriteVerbose( fmt, items);
        }
    }
}