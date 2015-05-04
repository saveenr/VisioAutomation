using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Open, "VisioMaster")]
    public class Open_VisioMaster : VisioCmdlet
    {
        [SMA.ParameterAttribute(Position = 0, Mandatory = true)]
        [SMA.ValidateNotNullAttribute]
        public IVisio.Master Master;

        protected override void ProcessRecord()
        {
            this.client.Master.OpenForEdit(this.Master);
        }
    }
}