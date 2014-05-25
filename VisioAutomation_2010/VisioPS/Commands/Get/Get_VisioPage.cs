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

            var pages = scriptingsession.Page.GetPagesByName(this.Name);
            this.WriteObject(pages, true);
        }
    }
}