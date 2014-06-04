using System.Linq;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioPage")]
    public class Get_VisioPage : VisioCmdlet
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