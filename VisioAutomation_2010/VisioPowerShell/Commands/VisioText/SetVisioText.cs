using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioText
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, Nouns.VisioText)]
    public class SetVisioText : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string[] Text { get; set; }

        // CONTEXT:SHAPE 
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.TargetShapes(this.Shapes);
            this.Client.Text.SetShapeText(targetshapes, this.Text);
        }
    }
}
