using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Set, "VisioShapeSheet")]
    public class Set_VisioShapeSheet : VisioCmdlet
    {
        [SMA.ParameterAttribute(Position = 0, Mandatory = true)]
        public VisioAutomation.ShapeSheet.SRC[] Cell { get; set; }

        [SMA.ParameterAttribute(Position = 1, Mandatory = true)]
        public string[] Value { get; set; }

        [SMA.ParameterAttribute(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards;

        [SMA.ParameterAttribute(Mandatory = false)]
        public SMA.SwitchParameter TestCircular;

        [SMA.ParameterAttribute(Mandatory = false)]
        public SMA.SwitchParameter SetResults;

        [SMA.ParameterAttribute(Mandatory = false)]
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
                this.client.ShapeSheet.SetFormula(this.Shapes, this.Cell, this.Value, (IVisio.VisGetSetArgs)flags);               
            }
            else
            {
                this.client.ShapeSheet.SetResult(this.Shapes, this.Cell, this.Value, (IVisio.VisGetSetArgs)flags);                               
            }
        }
    }
}