using System.Collections.Generic;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Close, "VisioDocument")]
    public class Close_VisioDocument : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Document[] Documents;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Force;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            if (this.Documents== null)
            {
                var app = scriptingsession.VisioApplication;
                var doc = app.ActiveDocument;
                if (doc != null)
                {
                    VA.Documents.DocumentHelper.Close(doc,this.Force);
                }
            }
            else
            {
                foreach (var doc in this.Documents)
                {
                    VA.Documents.DocumentHelper.Close(doc, this.Force);
                }
            }
        }
    }
}