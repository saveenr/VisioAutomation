using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Duplicate", "VisioPage")]
    public class Duplicate_VisioPage : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        private SMA.SwitchParameter NewDoc;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            if (!NewDoc.ToBool())
            {
                scriptingsession.Page.Duplicate();
            }
            else
            {
                scriptingsession.Page.DuplicateToNewDocument();
            }
        }
    }
}