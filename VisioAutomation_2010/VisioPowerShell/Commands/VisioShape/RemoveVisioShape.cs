using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, Nouns.VisioShape)]
    public class RemoveVisioShape : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            // TODO: Investigate why this doesn't use the Targets method of identifying shapes
            if (this.Shapes == null)
            {
                var selection = new VisioScripting.TargetSelection();

                this.Client.Selection.DeleteShapes(selection);                
            }
            else
            {
                foreach (var shape in this.Shapes)
                {
                    shape.Delete();
                }
            }
        }
    }
}