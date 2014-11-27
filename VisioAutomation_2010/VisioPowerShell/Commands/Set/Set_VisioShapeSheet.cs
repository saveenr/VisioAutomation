using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioShapeSheet")]
    public class Set_VisioShapeSheet : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public VisioAutomation.ShapeSheet.SRC[] Cell { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public string[] Value { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter SetResults;

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
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
                this.client.ShapeSheet.SetFormula(this.Shapes, Cell, Value, (IVisio.VisGetSetArgs)flags);               
            }
            else
            {
                this.client.ShapeSheet.SetResult(this.Shapes, Cell, Value, (IVisio.VisGetSetArgs)flags);                               
            }
        }
    }
}