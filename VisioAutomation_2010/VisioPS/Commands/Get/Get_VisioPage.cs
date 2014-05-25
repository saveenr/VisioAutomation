using System.Linq;
using VisioAutomation.Extensions;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioPage")]
    public class Get_VisioPage : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position=0, Mandatory = false)]
        public string Name=null;

        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter ActivePage;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var application = scriptingsession.VisioApplication;

            if (this.ActivePage)
            {
                var page = scriptingsession.Page.Get();
                this.WriteObject(page);
                return;
            }

            if (Name == null || Name=="*" )
            {
                // return all pages
                var active_document = application.ActiveDocument;
                var pages = active_document.Pages.AsEnumerable().ToList();
                this.WriteObject(pages,false);
            }
            else
            {
                // return the named page
                var active_document = application.ActiveDocument;
                var pages = active_document.Pages;

                this.Name = this.Name.Trim();

                var regex = VisioAutomation.TextUtil.GetRegexForWildcardPattern(this.Name, true);
                var pages2 = pages.AsEnumerable().Where(d => regex.IsMatch(d.Name)).ToList();
                this.WriteObject(pages2, true);
            }
        }
    }
}