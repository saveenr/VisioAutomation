using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Close, "VisioApplication")]
    public class Close_VisioApplication : VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            if (Globals.Application == null)
            {
                this.WriteWarning("There is no Visio Application to stop");
                return;
            }

            Globals.Application.Quit();
        }
    }
}