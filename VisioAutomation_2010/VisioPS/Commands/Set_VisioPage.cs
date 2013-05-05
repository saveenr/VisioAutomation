using VAS=VisioAutomation.Scripting;
using VA = VisioAutomation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioPage")]
    public class Set_VisioPage : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true, ParameterSetName = "Name")]
        public string Name { get; set; }

        [SMA.Parameter(Position = 0, Mandatory = true, ParameterSetName = "Page")]
        public IVisio.Page Page  { get; set; }

        [SMA.Parameter(Position = 0, Mandatory = true, ParameterSetName = "Flags")]
        public VA.Scripting.PageNavigation Flag { get; set; }

        protected override void ProcessRecord()
        {
            if (this.Name != null)
            {
                this.ScriptingSession.Page.Set(this.Name);
            }
            else if (this.Page != null)
            {
                this.ScriptingSession.Page.Set(this.Page);
            }
            else
            {
                var scriptingsession = this.ScriptingSession;
                scriptingsession.Page.GoTo(this.Flag);                
            }
        }
    }
}