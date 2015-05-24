using System.Management.Automation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.New
{
    [Cmdlet(SMA.VerbsCommon.New, "VisioGroup")]
    public class New_VisioGroup : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var group = this.client.Arrange.Group();
            this.WriteObject(group);
        }
    }
}