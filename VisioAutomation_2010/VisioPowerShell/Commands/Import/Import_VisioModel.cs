using VA = VisioAutomation;
using SMA = System.Management.Automation;
using SXL= System.Xml.Linq;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsData.Import, "VisioModel")]
    public class Import_VisioModel : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = true, Position = 0)]
        [SMA.ValidateNotNullOrEmpty]
        public string Filename { get; set; }

        protected override void ProcessRecord()
        {
            if (!this.CheckFileExists(this.Filename))
            {
                return;
            }

            this.WriteVerbose("Loading {0} as xml", this.Filename);
            var xmldoc = SXL.XDocument.Load(this.Filename);

            var root = xmldoc.Root;
            this.WriteVerbose("Root element name ={0}", root.Name);
            if (root.Name == "directedgraph")
            {
                this.WriteVerbose("Loading as a Directed Graph");
                var dg_model = VA.Scripting.DirectedGraph.DirectedGraphBuilder.LoadFromXML(
                    this.client,
                    xmldoc);
                this.WriteObject(dg_model);               
            }
            else if (root.Name == "orgchart")
            {
                this.WriteVerbose("Loading as an Org Chart");
                var oc = VA.Scripting.OrgChart.OrgChartBuilder.LoadFromXML(this.client, xmldoc);
                this.WriteObject(oc);
            }
            else
            {
                var exc = new System.ArgumentException("Unknown root element for XML");
                throw exc;
            }
        }
    }
}