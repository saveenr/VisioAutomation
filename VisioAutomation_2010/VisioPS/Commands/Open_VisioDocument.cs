using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Open, "VisioDocument")]
    public class Open_VisioDocument : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string Filename { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            var ext = System.IO.Path.GetExtension(this.Filename).ToLowerInvariant();
            if (ext == ".vss" || ext == ".vst")
            {
                var doc = scriptingsession.Document.OpenStencil(this.Filename);
                this.WriteObject(doc);                
            }
            else
            {
                var doc = scriptingsession.Document.Open(this.Filename);
                this.WriteObject(doc);                
            }
        }
    }
}