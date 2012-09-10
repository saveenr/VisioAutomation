using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Export", "SelectionAsXAML")]
    public class Export_SelectionAsXAML : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string Filename;

        protected override void ProcessRecord()
        {
            var ss = this.ScriptingSession;
            ss.Export.ExportSelectionToXAML(this.Filename);
        }
    }
}