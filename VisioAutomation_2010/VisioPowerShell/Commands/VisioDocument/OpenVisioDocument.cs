

namespace VisioPowerShell.Commands.VisioDocument;

[SMA.Cmdlet(SMA.VerbsCommon.Open, Nouns.VisioDocument)]
public class OpenVisioDocument : VisioCmdlet
{
    [SMA.Parameter(Position = 0, Mandatory = true)]
    [SMA.ValidateNotNullOrEmpty]
    public string Filename { get; set; }

    protected override void ProcessRecord()
    {
        this._new_app_if_needed();

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