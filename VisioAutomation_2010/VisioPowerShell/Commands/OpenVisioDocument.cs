using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Open, VisioPowerShell.Commands.Nouns.VisioDocument)]
    public class OpenVisioDocument : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        [SMA.ValidateNotNullOrEmpty]
        public string Filename { get; set; }

        protected override void ProcessRecord()
        {
            if (this.Client.Application.HasApplication == false)
            {
                // no app - let's create one
                this.Client.Application.NewApplication();
            }

            if (this.filename_is_stencil(this.Filename))
            {
                var doc = this.Client.Document.OpenStencilDocument(this.Filename);
                this.WriteObject(doc);                
            }
            else
            {
                var doc = this.Client.Document.OpenDocument(this.Filename);
                this.WriteObject(doc);                
            }
        }

        public bool filename_is_stencil(string fname)
        {
            var ext = System.IO.Path.GetExtension(fname).ToLowerInvariant();
            return (ext == ".vss" || ext == ".vst" || ext == ".vssx" || ext == ".vstx");
        }
    }
}