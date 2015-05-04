using System.IO;
using SMA = System.Management.Automation;
using VA=VisioAutomation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Open, "VisioDocument")]
    public class Open_VisioDocument : VisioCmdlet
    {
        [SMA.ParameterAttribute(Position = 0, Mandatory = true)]
        [SMA.ValidateNotNullOrEmptyAttribute]
        public string Filename { get; set; }

        protected override void ProcessRecord()
        {
            this.client.Application.SafeNew();

            if (this.filename_is_stencil(this.Filename))
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

        public bool filename_is_stencil(string fname)
        {
            var ext = Path.GetExtension(fname).ToLowerInvariant();
            return (ext == ".vss" || ext == ".vst" || ext == ".vssx" || ext == ".vstx");
        }
    }
}