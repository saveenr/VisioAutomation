using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.New, "VisioMaster")]
    public class New_VisioMaster : VisioCmdlet
    {
        [SMA.ParameterAttribute(Mandatory = false)]
        public string Name;

        [SMA.ParameterAttribute(Mandatory = false)]
        public IVisio.Document Document;

        protected override void ProcessRecord()
        {
            var master = this.client.Master.New(this.Document, this.Name);
            this.WriteObject(master);
        }
    }
}