using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsData.Save, "VisioDocument")]
    public class Save_VisioDocument : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        public string Filename;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            if (Filename!=null)
            {
                scriptingsession.Document.SaveAs(this.Filename);
            }
            else
            {
                scriptingsession.Document.Save();
            }
        }
    }
}