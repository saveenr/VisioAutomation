using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsData.Export, "VisioSelectionAsXHTML")]
    public class Export_VisioSelectionAsXHTML : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string Filename;

        protected override void ProcessRecord()
        {
            if (!this.CheckFileExists(Filename))
            {
                return;
            }

            var scriptingsession = this.ScriptingSession;
            scriptingsession.Export.SelectionToSVGXHTML(this.Filename);
        }
    }
}