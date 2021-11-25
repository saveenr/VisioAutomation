namespace VisioPowerShell.Commands.VisioDocument;

[SMA.Cmdlet(SMA.VerbsCommon.Select, Nouns.VisioDocument)]
public class SelectVisioDocument : VisioCmdlet
{

    // NONCONTEXT:DOCUMENT
    [SMA.Parameter(Position = 0, Mandatory = true)]
    [SMA.ValidateNotNull]
    public IVisio.Document Document  { get; set; }
        
    protected override void ProcessRecord()
    {
        this.Client.Document.ActivateDocument(this.Document);
    }
}