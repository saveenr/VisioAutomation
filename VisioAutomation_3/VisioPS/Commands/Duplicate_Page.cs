using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Duplicate", "Page")]
    public class Duplicate_Page : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        private SMA.SwitchParameter NewDoc;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            if (!NewDoc.ToBool())
            {
                scriptingsession.Page.DuplicatePage();
            }
            else
            {
                scriptingsession.Page.DuplicatePageToNewDocument();
            }
        }
    }
}