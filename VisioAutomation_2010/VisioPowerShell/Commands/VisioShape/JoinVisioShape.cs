using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Join, Nouns.VisioShape)]
    public class JoinVisioShape : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var group = this.Client.Grouping.GroupSelectedShapes();
            this.WriteObject(group);
        }
    }
}