using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioGroup)]
    public class NewVisioGroup : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var group = this.Client.Grouping.Group();
            this.WriteObject(group);
        }
    }
}