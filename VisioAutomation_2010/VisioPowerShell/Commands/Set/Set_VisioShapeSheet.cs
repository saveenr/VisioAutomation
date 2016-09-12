using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.Set
{
    [Cmdlet(VerbsCommon.Set, VisioPowerShell.Nouns.VisioShapeSheet)]
    public class Set_VisioShapeSheet : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        public VisioAutomation.ShapeSheet.SRC[] Cell { get; set; }

        [Parameter(Position = 1, Mandatory = true)]
        public string[] Value { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter BlastGuards;

        [Parameter(Mandatory = false)]
        public SwitchParameter TestCircular;

        [Parameter(Mandatory = false)]
        public SwitchParameter SetResults;

        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var targets = new VisioAutomation.Scripting.TargetShapes(this.Shapes);

            short flags = 0;
            
            if (this.BlastGuards)
            {
                flags = (short)(flags | (short)IVisio.VisGetSetArgs.visSetBlastGuards);
            }
            
            if (this.TestCircular)
            {
                flags = (short)(flags | (short)IVisio.VisGetSetArgs.visSetTestCircular);
            }

            if (!this.SetResults)
            {
                this.Client.ShapeSheet.SetFormula(targets, this.Cell, this.Value, (IVisio.VisGetSetArgs)flags);               
            }
            else
            {
                this.Client.ShapeSheet.SetResult<string>(targets, this.Cell, this.Value, (IVisio.VisGetSetArgs)flags);                               
            }
        }
    }
}