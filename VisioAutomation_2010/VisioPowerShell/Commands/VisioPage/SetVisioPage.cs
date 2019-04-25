using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioPage
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, Nouns.VisioPage)]
    public class SetVisioPage : VisioCmdlet
    {
        // NONCONTEXT:PAGE

        [SMA.Parameter(Position = 0, Mandatory = true)]
        [SMA.ValidateNotNull]
        public IVisio.Page Page  { get; set; }
        
        protected override void ProcessRecord()
        {
            this.Client.Page.SetActivePage(this.Page);
        }
    }
}