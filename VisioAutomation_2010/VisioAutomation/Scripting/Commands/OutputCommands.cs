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
            this._client.WriteUser(s);
        }

        public void WriteDebug(string s)
        {
            this._client.WriteDebug(s);
        }

        public void WriteError(string s)
        {
            this._client.WriteError(s);
        }

        public void WriteVerbose(string s)
        {
            this._client.WriteVerbose(s);
        }

        public void WriteUser(string fmt, params object[] items)
        {
            this._client.WriteUser(fmt, items);
        }

        public void WriteDebug(string fmt, params object[] items)
        {
            this._client.WriteDebug( fmt, items);           
        }

        public void WriteError(string fmt, params object[] items)
        {
            this._client.WriteError( fmt, items);
        }

        public void WriteVerbose(string fmt, params object[] items)
        {
            this._client.WriteVerbose( fmt, items);
        }
    }
}