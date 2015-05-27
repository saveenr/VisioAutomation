using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.New
{
    [Cmdlet(VerbsCommon.New, "VisioMaster")]
    public class New_VisioMaster : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public string Name;

        [Parameter(Mandatory = false)]
        public IVisio.Document Document;

        protected override void ProcessRecord()
        {
            var master = this.client.Master.New(this.Document, this.Name);
            this.WriteObject(master);
        }
    }
}