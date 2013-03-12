using System.Linq;
using VisioAutomation.Extensions;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioPage")]
    public class Get_VisioPage : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(ParameterSetName="named",Position=0, Mandatory = false)]
        public string Name=null;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var application = scriptingsession.VisioApplication;

            if (Name == null)
            {
                // return the active page
                this.WriteObject(scriptingsession.Page.Get());
            }
            else if (Name=="*" )
            {
                // return all pages
                var active_document = application.ActiveDocument;
                var pages = active_document.Pages.AsEnumerable().ToList();
                this.WriteObject(pages);
            }
            else
            {
                // return the named page
                var active_document = application.ActiveDocument;
                var pages = active_document.Pages;
                var page = pages[Name];
                this.WriteObject(page);
            }
        }
    }
}