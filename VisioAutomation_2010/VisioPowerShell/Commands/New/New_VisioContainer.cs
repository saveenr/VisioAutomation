using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioContainer")]
    public class New_VisioContainer : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public IVisio.Master Masters { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public double[] Points { get; set; }

        protected override void ProcessRecord()
        {
            var shape = this.client.Master.DropContainer(Masters);
            this.WriteObject(shape);
        }
    }
}
