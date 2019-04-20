using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioPage
{
    [SMA.Cmdlet(SMA.VerbsCommon.Copy, Nouns.VisioPage)]
    public class CopyVisioPage : VisioCmdlet
    {
        // NONCONTEXT:DOCUMENT

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Document ToDocument=null;

        protected override void ProcessRecord()
        {
            IVisio.Page newpage;
            if (this.ToDocument == null)
            {
                newpage = this.Client.Page.DuplicatePage(VisioScripting.TargetPage.Auto);
            }
            else
            {
                newpage = this.Client.Page.DuplicatePageToDocument(VisioScripting.TargetPage.Auto, this.ToDocument);
            }

            this.WriteObject(newpage);            
        }
    }
}