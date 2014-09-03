using VisioAutomation;

namespace VisioPowerShell.Commands
{
    [System.Management.Automation.Cmdlet(System.Management.Automation.VerbsCommon.Close, "VisioMaster")]
    public class Close_VisioMaster : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var window = this.ScriptingSession.VisioApplication.ActiveWindow;

            var st = window.SubType;
            if (st != 64)
            {
                throw new AutomationException("The active window is not a master window");
            }


            var master = (Microsoft.Office.Interop.Visio.Master)window.Master;
            master.Close();
        }
    }
}