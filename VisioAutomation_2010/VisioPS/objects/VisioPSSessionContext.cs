using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS
{
    public class VisioPSSessionContext : VisioAutomation.Scripting.SessionContext
    {
        private readonly SMA.Cmdlet cmdlet;
        
        public VisioPSSessionContext(SMA.Cmdlet cmdlet)
        {
            this.cmdlet = cmdlet;
        }

        public override void WriteDebug(string s)
        {
            this.cmdlet.WriteDebug(s);
        }

        public override void WriteError(string s)
        {
            this.cmdlet.WriteObject(s);
        }

        public override void WriteUser(string s)
        {
            this.cmdlet.WriteObject(s);
        }

        public override void WriteVerbose(string s)
        {
            this.cmdlet.WriteVerbose(s);
        }
    }
}