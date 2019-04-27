using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsCommon.Copy, Nouns.VisioShape)]
    public class CopyVisioShape : VisioCmdlet
    {
        // PSEUDOCONTEXT:SHAPES
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shape;

        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.TargetShapes(this.Shape);
            targetshapes.ResolveToSelection(this.Client);

            this.Client.Selection.DuplicateShapes(VisioScripting.TargetSelection.Auto);
        }
    }
}