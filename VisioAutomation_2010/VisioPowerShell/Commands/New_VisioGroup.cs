using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioGroup)]
    public class New_VisioGroup : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var group = this.Client.Grouping.Group();
            this.WriteObject(group);
        }
    }
}