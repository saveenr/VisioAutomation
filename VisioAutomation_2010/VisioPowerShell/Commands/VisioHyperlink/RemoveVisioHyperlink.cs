namespace VisioPowerShell.Commands.VisioHyperlink;

[SMA.Cmdlet(SMA.VerbsCommon.Remove, Nouns.VisioHyperlink)]
public class RemoveVisioHyperlink : VisioCmdlet
{
    [SMA.Parameter(Position = 0, Mandatory = true)]
    public int Index { get; set; }

    // CONTEXT:SHAPES
    [SMA.Parameter(Mandatory = false)]
    public IVisio.Shape[] Shape;

    protected override void ProcessRecord()
    {
        var targetshapes = new VisioScripting.TargetShapes(this.Shape);
        this.Client.Hyperlink.DeleteHyperlinkAtIndex(targetshapes,this.Index);
    }
}