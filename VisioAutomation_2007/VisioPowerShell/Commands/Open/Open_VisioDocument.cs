using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Open, "VisioDocument")]
    public class Open_VisioDocument : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        [SMA.ValidateNotNullOrEmpty]
        public string Filename { get; set; }

        protected override void ProcessRecord()
        {
            var ext = System.IO.Path.GetExtension(this.Filename).ToLowerInvariant();
            if (ext == ".vss" || ext == ".vst" || ext == ".vssx" || ext == ".vstx")
            {
                var doc = this.client.Document.OpenStencil(this.Filename);
                this.WriteObject(doc);                
            }
            else
            {
                var doc = this.client.Document.Open(this.Filename);
                this.WriteObject(doc);                
            }
        }
    }
}