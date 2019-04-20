using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioText
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioText)]
    public class GetVisioText : VisioCmdlet
    {
        // CONTEXT:SHAPES 

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;
        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.TargetShapes(this.Shapes);
            var listof_string = this.Client.Text.GetShapeText(targetshapes);
            this.WriteObject(listof_string);
        }
    }
}