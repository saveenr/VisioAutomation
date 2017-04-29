using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Copy, VisioPowerShell.Commands.Nouns.VisioPage)]
    public class CopyVisioPage : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public IVisio.Document ToDocument=null;

        protected override void ProcessRecord()
        {
            IVisio.Page newpage;
            if (this.ToDocument == null)
            {
                newpage = this.Client.Page.Duplicate();
            }
            else
            {
                newpage = this.Client.Page.Duplicate(this.ToDocument);
            }

            this.WriteObject(newpage);            
        }
    }
}