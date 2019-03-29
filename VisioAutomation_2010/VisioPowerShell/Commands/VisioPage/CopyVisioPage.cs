using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Copy, Nouns.VisioPage)]
    public class CopyVisioPage : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Document ToDocument=null;

        protected override void ProcessRecord()
        {
            var target_page = new VisioScripting.Models.TargetPage();
            var page = target_page.Resolve(this.Client);

            IVisio.Page newpage;
            if (this.ToDocument == null)
            {
                newpage = this.Client.Page.DuplicateActivePage();
            }
            else
            {
                newpage = this.Client.Page.DuplicatePageToDocument(target_page, this.ToDocument);
            }

            this.WriteObject(newpage);            
        }
    }
}