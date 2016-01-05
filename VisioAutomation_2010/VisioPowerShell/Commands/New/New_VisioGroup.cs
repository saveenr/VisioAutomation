using System.Management.Automation;

namespace VisioPowerShell.Commands.New
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Nouns.VisioGroup)]
    public class New_VisioGroup : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var group = this.Client.Arrange.Group();
            this.WriteObject(group);
        }
    }
}