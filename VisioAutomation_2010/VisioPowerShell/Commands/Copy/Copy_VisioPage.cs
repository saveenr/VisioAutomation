using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.Copy
{
    [Cmdlet(VerbsCommon.Copy, "VisioPage")]
    public class Copy_VisioPage : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public IVisio.Document ToDocument=null;

        protected override void ProcessRecord()
        {
            IVisio.Page newpage;
            if (this.ToDocument == null)
            {
                newpage = this.client.Page.Duplicate();
            }
            else
            {
                newpage = this.client.Page.Duplicate(this.ToDocument);
            }

            this.WriteObject(newpage);            
        }
    }
}