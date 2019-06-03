using System;
using SXL = System.Xml.Linq;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioModel
{
    [SMA.Cmdlet(SMA.VerbsData.Import, Nouns.VisioModel)]
    public class ImportVisioModel : VisioCmdlet
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
            this.WriteVerbose("Root element name = \"{0}\"", root.Name);
            if (root.Name == "directedgraph")
            {
                this.WriteVerbose("Loading Directed Graph");
                var list_dglayout = VisioScripting.Builders.DirectedGraphDocumentLoader.LoadFromXml(this.Client, xmldoc);
                this.WriteObject(list_dglayout);               
            }
            else if (root.Name == "orgchart")
            {
                this.WriteVerbose("Loading as Org Chart");
                var orgchart_docmodel = VisioScripting.Builders.OrgChartBuilder.LoadFromXml(this.Client, xmldoc);
                this.WriteObject(orgchart_docmodel);
            }
            else
            {
                var exc = new ArgumentOutOfRangeException("Unknown root element for XML");
                throw exc;
            }
        }

        protected bool CheckFileExists(string file)
        {
            if (System.IO.File.Exists(file)) return true;

            this.WriteVerbose("Filename: {0}", file);
            this.WriteVerbose("Abs Filename: {0}", System.IO.Path.GetFullPath(file));
            var exc = new System.IO.FileNotFoundException(file);
            var er = new SMA.ErrorRecord(exc, "FILE_NOT_FOUND", SMA.ErrorCategory.ResourceUnavailable, null);
            this.WriteError(er);
            return false;
        }
    }
}