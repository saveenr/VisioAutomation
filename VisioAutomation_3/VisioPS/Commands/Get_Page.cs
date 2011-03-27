using System.Linq;
using VisioAutomation.Extensions;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "Page")]
    public class Get_Page : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position=0, Mandatory = false)]
        public string Name=null;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var application = scriptingsession.VisioApplication;
            if (Name=="*")
            {
                var active_document = application.ActiveDocument;
                var pages = active_document.Pages.AsEnumerable().ToList();
                this.WriteObject(pages);               
            }
            else if (Name != null)
            {
                var active_document = application.ActiveDocument;
                var pages = active_document.Pages;
                var page = pages[Name];
                this.WriteObject(page);
            }
            else if (Name==null)
            {
                var active_page = application.ActivePage;
                this.WriteObject(active_page);
            }
        }
    }
}