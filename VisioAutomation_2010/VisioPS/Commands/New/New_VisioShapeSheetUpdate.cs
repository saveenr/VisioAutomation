using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using VA=VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioShapeSheetUpdate")]
    public class New_VisioShapeSheetUpdate : VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)] 
        public SMA.SwitchParameter BlastGuards;
        
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular;

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Page Page;

        protected override void ProcessRecord()
        {
            var session = this.ScriptingSession;
            
            if (!this.ScriptingSession.HasActiveDocument)
            {
                return;
            }

            if (this.Page == null)
            {
                this.Page = session.VisioApplication.ActivePage;                
            }

            var update = new VA.Scripting.ShapeSheetUpdate(this.ScriptingSession,this.Page);
            update.BlastGuards = this.BlastGuards;
            update.TestCircular = this.TestCircular;
            this.WriteObject(update);
        }
    }
}