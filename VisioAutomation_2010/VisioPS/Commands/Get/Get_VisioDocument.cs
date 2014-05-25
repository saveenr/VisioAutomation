using System.Linq;
using VisioAutomation.Extensions;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioDocument")]
    public class Get_VisioDocument : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        [SMA.ValidateNotNullOrEmpty]
        public string Name = null;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ActiveDocument;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var application = scriptingsession.VisioApplication;

            if (this.ActiveDocument)
            {
                var active_doc = application.ActiveDocument;
                this.WriteObject(active_doc);
                return;
            }

            var docs = scriptingsession.Document.GetDocumentsByName(this.Name);
            this.WriteObject(docs, true);
        }
    }
}