using VisioPS.Extensions;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Draw", "OrgChart")]
    public class Draw_OrgChart : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = true)]
        public string Filename { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var oc = VA.Scripting.OrgChart.OrgChartBuilder.LoadFromXML(scriptingsession, this.Filename);
            VA.Scripting.OrgChart.OrgChartBuilder.RenderDiagrams(scriptingsession, oc);
        }
    }
}