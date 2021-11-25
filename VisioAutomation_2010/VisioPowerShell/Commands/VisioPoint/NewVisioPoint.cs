using System.Linq;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioRectangle
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, Nouns.VisioPoint)]
    public class NewVisioPoint : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public float X { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public float Y { get; set; }

        protected override void ProcessRecord()
        {
            var point = new VisioAutomation.Core.Point(this.X, this.Y);
            this.WriteObject(point);
        }
    }

}