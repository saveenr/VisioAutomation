using SMA = System.Management.Automation;

namespace VisioPowerShell
{
    public class VisioPsClientContext : VisioScripting.ClientContext
    {
        private readonly SMA.Cmdlet _cmdlet;
        
        public VisioPsClientContext(SMA.Cmdlet cmdlet)
        {
            this._cmdlet = cmdlet;
        }

        public override void WriteDebug(string s)
        {
            this._cmdlet.WriteDebug(s);
        }

        public override void WriteError(string s)
        {
            this._cmdlet.WriteObject(s);
        }

        public override void WriteUser(string s)
        {
            this._cmdlet.WriteObject(s);
        }

        public override void WriteVerbose(string s)
        {
            this._cmdlet.WriteVerbose(s);
        }

        public override void WriteWarning(string s)
        {
            this._cmdlet.WriteWarning(s);
        }

    }
}