using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using VA=VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsData.Update, "VisioShapeSheet")]
    public class Update_VisioShapeSheet : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public VA.Scripting.ShapeSheetUpdate Update { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            scriptingsession.ShapeSheet.Update(this.Update, this.BlastGuards, this.TestCircular);
        }
    }
}