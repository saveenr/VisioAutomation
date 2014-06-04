using System.Collections.Generic;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, "VisioPage")]
    public class Remove_VisioPage : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Page[] Pages;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Renumber;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.Page.Delete(this.Pages, this.Renumber);
        }
    }
}