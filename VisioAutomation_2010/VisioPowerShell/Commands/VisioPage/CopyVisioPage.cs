namespace VisioPowerShell.Commands.VisioPage;

[SMA.Cmdlet(SMA.VerbsCommon.Copy, Nouns.VisioPage)]
public class CopyVisioPage : VisioCmdlet
{
    [SMA.Parameter(Mandatory = false)]
    public IVisio.Document ToDocument=null;

    [SMA.Parameter(Mandatory = false)]
    public IVisio.Page Page;

    protected override void ProcessRecord()
    {
        var targetpage = new VisioScripting.TargetPage(this.Page);

        IVisio.Page newpage;
        if (this.ToDocument == null)
        {
            newpage = this.Client.Page.DuplicatePage(targetpage);
        }
        else
        {
            newpage = this.Client.Page.DuplicatePageToDocument(targetpage, this.ToDocument);
        }

        this.WriteObject(newpage);            
    }
}