using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsCommon.Join, Nouns.VisioShape)]
    public class JoinVisioShape : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var group = this.Client.Grouping.Group(VisioScripting.TargetSelection.Active);
            this.WriteObject(group);
        }
    }
}