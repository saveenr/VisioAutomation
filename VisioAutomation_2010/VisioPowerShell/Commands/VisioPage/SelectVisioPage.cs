namespace VisioPowerShell.Commands.VisioPage;

[SMA.Cmdlet(SMA.VerbsCommon.Select, Nouns.VisioPage)]
public class SelectVisioPage : VisioCmdlet
{
    // NONCONTEXT:PAGE

    [SMA.Parameter(Position = 0, Mandatory = true)]
    [SMA.ValidateNotNull]
    public IVisio.Page Page  { get; set; }
        
    protected override void ProcessRecord()
    {
        this.Client.Page.SetActivePage(this.Page);
    }
}