using System.Linq;
using VisioAutomation.Extensions;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioDocument")]
    public class Get_VisioDocument : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
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

            if (this.Name == null || this.Name == "*")
            {
                // return all pages
                var documents = application.Documents;
                var docs = documents.AsEnumerable().ToList();
                this.WriteObject(docs,false);                
            }
            else 
            {
                // get the named document
                var documents = application.Documents;

                this.Name = this.Name.Trim();

                var regex = VisioAutomation.TextUtil.GetRegexForWildcardPattern(this.Name, true);
                var docs2 = documents.AsEnumerable().Where(d => regex.IsMatch(d.Name)).ToList();
                this.WriteObject(docs2,true);
            }
        }
    }
}