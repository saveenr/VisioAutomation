using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Save", "VisioDocument")]
    public class Save_VisioDocument : VisioPS.VisioPSCmdlet
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