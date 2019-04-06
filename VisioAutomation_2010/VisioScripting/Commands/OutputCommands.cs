namespace VisioScripting.Commands
{
    public class OutputCommands : CommandSet
    {
        internal OutputCommands(Client client) :
            base(client)
        {

        }

        public void WriteUser(string fmt, params object[] items)
        {
            string s = string.Format(fmt, items);
            this._client.ClientContext.WriteUser(s);
        }

        public void WriteDebug(string fmt, params object[] items)
        {
            string s = string.Format(fmt, items);
            this._client.ClientContext.WriteDebug(s);
        }

        public void WriteVerbose(string fmt, params object[] items)
        {
            string s = string.Format(fmt, items);
            this._client.ClientContext.WriteVerbose(s);
        }

        public void WriteWarning(string fmt, params object[] items)
        {
            string s = string.Format(fmt, items);
            this._client.ClientContext.WriteWarning(s);
        }

        public void WriteError(string fmt, params object[] items)
        {
            string s = string.Format(fmt, items);
            this._client.ClientContext.WriteError(s);
        }
    }
}