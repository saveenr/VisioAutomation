using System.Collections.Generic;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, "VisioPage")]
    public class Remove_VisioPage : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Page[] Pages;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Renumber;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            if (this.Pages== null)
            {
                var page = scriptingsession.Page.Get();
                page.Delete(this.Renumber ? (short)1 : (short)0);
            }
            else
            {
                foreach (var page in this.Pages)
                {
                    page.Delete(this.Renumber ? (short)1 : (short)0);
                }
            }
        }
    }
}