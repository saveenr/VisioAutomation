using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsData.Import, "VisioOrgChart")]
    public class Import_VisioOrgChart : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = true,Position= 0)]
        public string Filename { get; set; }

        protected override void ProcessRecord()
        {
            if (!this.CheckFileExists(this.Filename))
            {
                return;
            }

            var scriptingsession = this.ScriptingSession;
            var oc = VA.Scripting.OrgChart.OrgChartBuilder.LoadFromXML(scriptingsession, this.Filename);
            this.WriteObject(oc);
        }
    }
}