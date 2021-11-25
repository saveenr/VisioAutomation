namespace VisioPowerShell.Commands.VisioPage;

[SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioPage)]
public class GetVisioPage : VisioCmdlet
{
    [SMA.Parameter(Mandatory = false, ParameterSetName = "active")]
    public SMA.SwitchParameter ActivePage;

    [SMA.Parameter(Position=0, Mandatory = false, ParameterSetName = "pagebyname")]
    public string[] Name;

    [SMA.Parameter(Mandatory = false, ParameterSetName = "pagebyid")]
    public int[] ID;

    [SMA.Parameter(Mandatory = false)]
    public IVisio.Document Document;
        
    protected override void ProcessRecord()
    {
        if (this.ActivePage)
        {
            var page_active = this.Client.Page.GetActivePage();
            this.WriteObject(page_active);
            return;
        }

        // If the active page  is not specified then work on all the pages in a document (user-specified or auto)

        var targetdoc = new VisioScripting.TargetDocument(this.Document);

        // First, the ID case
        if (this.ID != null)
        {
            var t = targetdoc.ResolveToDocument(this.Client);
            foreach (var id in this.ID)
            {
                var page = t.Document.Pages[id];
                this.WriteObject(page);
            }
            return;
        }

        // Then, handle the name case

        if (this.Name == null)
        {
            var pages_by_name = this.Client.Page.FindPagesInDocument(targetdoc, null);
            this.WriteObject(pages_by_name, true);
            return;
        }

        var list_page = new List<IVisio.Page>();
        foreach (var name in this.Name)
        {
            var pages_by_name = this.Client.Page.FindPagesInDocument(targetdoc, name);
            list_page.AddRange(pages_by_name);
        }
        this.WriteObject(list_page, true);
    }
}