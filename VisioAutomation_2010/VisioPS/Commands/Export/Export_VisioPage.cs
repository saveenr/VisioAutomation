using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsData.Export, "VisioPage")]
    public class Export_VisioPage : VisioCmdlet
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
                scriptingsession.Export.PageToFile(this.Filename);
            }
            else
            {
                // is -AllPages is set then export them all
                var scriptingsession = this.ScriptingSession;
                scriptingsession.Export.PagesToFiles(this.Filename);
            }
        }
    }
}