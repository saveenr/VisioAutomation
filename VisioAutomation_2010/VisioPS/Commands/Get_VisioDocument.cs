using System.Linq;
using VisioAutomation.Extensions;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioDocument")]
    public class Get_VisioDocument : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(ParameterSetName="named",Position = 0, Mandatory = false)]
        public string Name = null;

        [SMA.Parameter(ParameterSetName = "active", Position = 0, Mandatory = false)]
        public SMA.SwitchParameter ActiveDocument;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var application = scriptingsession.VisioApplication;

            if (this.ActiveDocument)
            {
                // return the active document
                var active_doc = application.ActiveDocument;
                this.WriteObject(active_doc);
            }
            else if (this.Name=="*" || this.Name == null)
            {
                // return all pages
                var documents = application.Documents;
                var docs = documents.AsEnumerable().ToList();
                this.WriteObject(docs);                
            }
            else 
            {
                // get the named document
                var docs = application.Documents;
                var doc = docs[ Name ];
                this.WriteObject(doc);
            }
        }
    }
}