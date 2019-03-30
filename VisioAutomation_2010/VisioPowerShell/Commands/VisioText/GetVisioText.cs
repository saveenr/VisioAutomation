using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioText
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioText)]
    public class GetVisioText : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var targets = new VisioScripting.Models.TargetShapes(this.Shapes);
            var t = this.Client.Text.GetShapeText(targets);
            this.WriteObject(t);
        }
    }
}