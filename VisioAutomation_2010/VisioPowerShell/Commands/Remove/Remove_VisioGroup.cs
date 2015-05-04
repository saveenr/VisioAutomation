using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Remove, "VisioGroup")]
    public class Remove_VisioGroup : VisioCmdlet
    {
        [SMA.ParameterAttribute(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            this.client.Arrange.Ungroup(this.Shapes);
        }
    }
}