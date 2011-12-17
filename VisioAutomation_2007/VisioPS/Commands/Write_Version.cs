using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommunications.Write, "Version")]
    public class Write_Version : SMA.Cmdlet
    {
        protected override void ProcessRecord()
        {
            this.WriteObject("Version 1.0");
        }
    }
}