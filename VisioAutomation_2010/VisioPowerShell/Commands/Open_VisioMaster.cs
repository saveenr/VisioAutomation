using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Open, VisioPowerShell.Nouns.VisioMaster)]
    public class Open_VisioMaster : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        [ValidateNotNull]
        public IVisio.Master Master;

        protected override void ProcessRecord()
        {
            this.Client.Master.OpenForEdit(this.Master);
        }
    }
}