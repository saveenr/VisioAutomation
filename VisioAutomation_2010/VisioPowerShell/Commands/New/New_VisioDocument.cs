using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioDocument")]
    public class New_VisioDocument : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            if (!this.client.HasApplication)
            {
                this.client.Application.New();
            }
            else
            {
                if (!this.client.Application.Validate())
                {
                    this.client.Application.New();
                }
            }
            var doc = this.client.Document.New();
            this.WriteObject(doc);
        }
    }
}