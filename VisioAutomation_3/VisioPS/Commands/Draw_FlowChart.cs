using VisioPS.Extensions;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Draw", "FlowChart")]
    public class Draw_FlowChart : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = true)]
        public string Filename { get; set; }

        protected override void ProcessRecord()
        {
            var abs_filename = System.IO.Path.GetFullPath(this.Filename);
           
            if (!System.IO.File.Exists(abs_filename))
            {
                this.WriteVerbose("ERROR: File not found {0}",abs_filename);
                return;
            }

            var scriptingsession = this.ScriptingSession;

            if (scriptingsession.VisioApplication == null)
            {

                this.WriteVerbose("ERROR: No Visio Application is attached");
                return;
            }

            var flowchart_model = VA.Scripting.FlowChart.FlowChartBuilder.LoadFromXML(scriptingsession, abs_filename);
            VA.Scripting.FlowChart.FlowChartBuilder.RenderDiagrams(scriptingsession, flowchart_model);
        }
    }
}