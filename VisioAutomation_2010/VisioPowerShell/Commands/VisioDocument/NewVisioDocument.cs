using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioDocument)]
    public class NewVisioDocument : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public string Stencil { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string Template { get; set; }

        protected override void ProcessRecord()
        {
            if (!this.Client.Application.HasActiveApplication)
            {
                this.Client.Application.NewActiveApplication();
            }
            else
            {
                if (!this.Client.Application.ValidateActiveApplication())
                {
                    this.Client.Application.NewActiveApplication();
                }
            }

            var doc = this.Client.Document.NewDocumentFromTemplate(this.Template);

            if (!string.IsNullOrEmpty(this.Stencil))
            {
                var stencildoc = this.Client.Document.OpenStencilDocument(this.Stencil);
            }

            this.WriteObject(doc);
        }
    }
}