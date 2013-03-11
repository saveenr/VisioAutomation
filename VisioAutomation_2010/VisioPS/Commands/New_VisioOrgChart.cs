using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioOrgChart")]
    public class Bew_VisioOrgChart : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = true, Position=0, ParameterSetName = "xml")]
        public string Xml { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var xml = System.Xml.Linq.XDocument.Parse(this.Xml);
            var orgchart = VA.Scripting.OrgChart.OrgChartBuilder.LoadFromXML(ScriptingSession, xml);
            this.WriteObject(orgchart);
        }
    }
}