using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioHyperlink
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, Nouns.VisioHyperlink)]
    public class RemoveVisioHyperlink : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public int Index { get; set; }

        // CONTEXT:SHAPE
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.TargetShapes(this.Shapes);
            this.Client.Hyperlink.DeleteHyperlinkAtIndex(targetshapes,this.Index);
        }
    }
}