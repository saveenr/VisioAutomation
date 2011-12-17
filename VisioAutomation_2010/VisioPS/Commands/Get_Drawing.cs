using System.Linq;
using VisioAutomation.Extensions;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "Drawing")]
    public class Get_Drawing : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        public string Name = null;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var application = scriptingsession.VisioApplication;

            if (this.Name=="*")
            {
                var documents = application.Documents;
                var docs = documents.AsEnumerable().ToList();
                this.WriteObject(docs);                
            }
            else if (this.Name !=null)
            {
                var docs = application.Documents;
                var doc = docs[ Name ];
                this.WriteObject(doc);
                
            }
            else if (this.Name == null)
            {
                var active_doc = application.ActiveDocument;
                this.WriteObject(active_doc);
            }
        }
    }
}