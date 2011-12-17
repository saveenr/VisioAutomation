using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Export", "SelectionAsSVGXHTML")]
    public class Export_SelectionAsSVGXHTML : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string Filename;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.Export.ExportSelectionAsSVGXHTML(this.Filename);
        }
    }
}