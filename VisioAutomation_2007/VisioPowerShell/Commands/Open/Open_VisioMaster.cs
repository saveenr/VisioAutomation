using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Open, "VisioMaster")]
    public class Open_VisioMaster : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        [SMA.ValidateNotNull]
        public IVisio.Master Master;

        protected override void ProcessRecord()
        {
            this.client.Master.OpenForEdit(this.Master);
        }
    }
}