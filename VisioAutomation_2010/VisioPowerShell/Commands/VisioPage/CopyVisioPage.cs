using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioPage
{
    [SMA.Cmdlet(SMA.VerbsCommon.Copy, Nouns.VisioPage)]
    public class CopyVisioPage : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Document ToDocument=null;

        protected override void ProcessRecord()
        {
            var targetpage = new VisioScripting.TargetPage();

            IVisio.Page newpage;
            if (this.ToDocument == null)
            {
                newpage = this.Client.Page.Duplicate(targetpage);
            }
            else
            {
                newpage = this.Client.Page.DuplicateToDocument(targetpage, this.ToDocument);
            }

            this.WriteObject(newpage);            
        }
    }
}