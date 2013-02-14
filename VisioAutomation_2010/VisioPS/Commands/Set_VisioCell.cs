using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioFormula")]
    public class Set_VisioCell : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public VisioAutomation.ShapeSheet.SRC Cell { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public string Formula { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular;

        [SMA.Parameter(Mandatory = false)]
        public IList<IVisio.Shape> Shapes;

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
            scriptingsession.ShapeSheet.SetFormula(this.Shapes,new [] {Cell}, new [] {Formula}, (IVisio.VisGetSetArgs)flags);
        }
    }
}