using VisioAutomation;
using VAS=VisioAutomation.Scripting;
using IVisio=Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioMaster")]
    public class New_VisioMaster : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public string Name;

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Document Document;

        protected override void ProcessRecord()
        {
            var master = this.ScriptingSession.Master.New(this.Document, this.Name);
            this.WriteObject(master);
        }
    }

    [SMA.Cmdlet(SMA.VerbsCommon.Open, "VisioMaster")]
    public class Open_VisioMaster : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public IVisio.Master Master;

        protected override void ProcessRecord()
        {
            // Edit the master by adding a shape
            var mdraw_window = this.Master.OpenDrawWindow();
            mdraw_window.Activate();
        }
    }

    [SMA.Cmdlet(SMA.VerbsCommon.Close, "VisioMaster")]
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


            var master = (IVisio.Master)window.Master;
            master.Close();
        }
    }

}