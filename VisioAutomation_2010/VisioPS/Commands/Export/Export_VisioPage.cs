using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsData.Export, "VisioPage")]
    public class Export_VisioPage : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)] 
        public string Filename;

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public SMA.SwitchParameter AllPages;

        protected override void ProcessRecord()
        {
            if (!this.AllPages)
            {
                // this means use the current page 
                var scriptingsession = this.ScriptingSession;
                scriptingsession.Export.ExportPageToFile(this.Filename);
            }
            else
            {
                // is -AllPages is set then export them all
                var scriptingsession = this.ScriptingSession;
                scriptingsession.Export.ExportPagesToFiles(this.Filename);
            }
        }
    }
}