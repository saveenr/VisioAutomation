using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Copy, VisioPowerShell.Commands.Nouns.VisioPage)]
    public class CopyVisioPage : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Document ToDocument=null;

        protected override void ProcessRecord()
        {
            IVisio.Page newpage;
            if (this.ToDocument == null)
            {
                newpage = this.Client.Page.DuplicateActivePage();
            }
            else
            {
                newpage = this.Client.Page.DuplicateActivePageToDocument(this.ToDocument);
            }

            this.WriteObject(newpage);            
        }
    }
}