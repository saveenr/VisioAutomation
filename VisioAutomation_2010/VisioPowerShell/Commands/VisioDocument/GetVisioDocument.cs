
namespace VisioPowerShell.Commands.VisioDocument;

[SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioDocument)]
public class GetVisioDocument : VisioCmdlet
{
    [SMA.Parameter(Mandatory = false, ParameterSetName = "active")]
    public SMA.SwitchParameter ActiveDocument;

    [SMA.Parameter(Mandatory = false, ParameterSetName = "docbyname")]
    public string[] Name = null;
        
        
    protected override void ProcessRecord()
    {
        if (this.ActiveDocument)
        {
            var application = this.Client.Application.GetApplication();
            var active_doc = application.ActiveDocument;
            this.WriteObject(active_doc);
            return;
        }

        // If the active document is not specified then work on all the pages in the application

        if (this.Name == null)
        {
            // Get all docs
            var docs = this.Client.Document.FindDocuments(null);
            this.WriteObject(docs, true);
            return;
        }

        var list_doc = new List<IVisio.Document>();
        foreach (var name in Name)
        {
            var docs = this.Client.Document.FindDocuments(name);
            list_doc.AddRange(docs);

        }
        this.WriteObject(list_doc, true);
    }
}