using System;
using System.Management.Automation;
using VisioScripting.Builders;
using SXL = System.Xml.Linq;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsData.Import, VisioPowerShell.Commands.Nouns.VisioModel)]
    public class Import_VisioModel : VisioCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        [ValidateNotNullOrEmpty]
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
                var dg_model = DirectedGraphBuilder.LoadFromXML(
                    this.Client,
                    xmldoc);
                this.WriteObject(dg_model);               
            }
            else if (root.Name == "orgchart")
            {
                this.WriteVerbose("Loading as an Org Chart");
                var oc = OrgChartBuilder.LoadFromXml(this.Client, xmldoc);
                this.WriteObject(oc);
            }
            else
            {
                var exc = new ArgumentException("Unknown root element for XML");
                throw exc;
            }
        }
    }
}