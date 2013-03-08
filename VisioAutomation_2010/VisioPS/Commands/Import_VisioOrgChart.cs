using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsData.Import, "VisioOrgChart")]
    public class Import_VisioOrgChart : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = true,ParameterSetName = "filename")]
        public string Filename { get; set; }

        [SMA.Parameter(Mandatory = true, ParameterSetName = "xml")]
        public string Xml { get; set; }

        protected override void ProcessRecord()
        {
            if (!this.CheckFileExists(this.Filename))
            {
                return;
            }

            var scriptingsession = this.ScriptingSession;
            if (this.Filename != null)
            {
                var oc = VA.Scripting.OrgChart.OrgChartBuilder.LoadFromXML(scriptingsession, this.Filename);
                this.WriteObject(oc);
            }
            else if (this.Xml!= null)
            {
                var x = System.Xml.Linq.XDocument.Parse(this.Xml);
                var oc = VA.Scripting.OrgChart.OrgChartBuilder.LoadFromXML(ScriptingSession, x);
                this.WriteObject(oc);
            }
        }
    }
}