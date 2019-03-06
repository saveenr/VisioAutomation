using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, Nouns.VisioGroup)]
    public class NewVisioGroup : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var group = this.Client.Grouping.GroupSelectedShapes();
            this.WriteObject(group);
        }
    }
}