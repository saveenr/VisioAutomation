using VAS=VisioAutomation.Scripting;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioZoom")]
    public class Set_VisioZoom : VisioPSCmdlet
    {
        [SMA.Parameter(ParameterSetName = "level", Position = 0, Mandatory = true)] 
        public VA.Scripting.Zoom Level = VA.Scripting.Zoom.In;

        [SMA.Parameter(ParameterSetName = "percent", Position = 0, Mandatory = true)] 
        public double Percent = 0;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            if (this.Percent > 0)
            {
                scriptingsession.View.ZoomToPercentage(this.Percent);
            }
            else
            {
                scriptingsession.View.Zoom(this.Level);       
            }
        }
    }
}