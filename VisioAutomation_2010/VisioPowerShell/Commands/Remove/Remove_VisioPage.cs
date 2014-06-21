using System.Collections.Generic;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, "VisioPage")]
    public class Remove_VisioPage : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false, Position=0, ValueFromPipeline = true)]
        public IVisio.Page[] Pages;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Renumber;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            if (this.Pages == null)
            {
                this.WriteVerboseEx("No Page objects ");
                this.WriteVerboseEx("Removing the Active Page");
                var page = scriptingsession.VisioApplication.ActivePage;
                scriptingsession.Page.Delete(new [] { page }, this.Renumber);
                return;
            }

            if (this.Pages != null)
            {
                this.WriteVerboseEx("Removing the Page Objects");
                scriptingsession.Page.Delete(this.Pages, this.Renumber);                
            }
        }
    }
}