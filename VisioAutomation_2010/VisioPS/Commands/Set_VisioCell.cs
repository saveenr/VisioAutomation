using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioCell")]
    public class Set_VisioCell : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public VisioAutomation.ShapeSheet.SRC Cell { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public string Result { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular;

        [SMA.Parameter(Mandatory = false)]
        public IList<IVisio.Shape> Shapes;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter SetResults;

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

            var scriptingsession = this.ScriptingSession;

            if (!this.SetResults)
            {
                scriptingsession.ShapeSheet.SetFormula(this.Shapes, new[] { Cell }, new[] { Result }, (IVisio.VisGetSetArgs)flags);               
            }
            else
            {
                var d = double.Parse(Result);
                scriptingsession.ShapeSheet.SetResult(this.Shapes, new[] { Cell }, new[] { d }, (IVisio.VisGetSetArgs)flags);                               
            }
        }
    }
}