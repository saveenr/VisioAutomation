

namespace VisioPowerShell.Commands.VisioText;

[SMA.Cmdlet(SMA.VerbsCommon.Set, Nouns.VisioText)]
public class SetVisioText : VisioCmdlet
{
    [SMA.Parameter(Position = 0, Mandatory = true)]
    public string[] Text { get; set; }

    // CONTEXT:SHAPES 
    [SMA.Parameter(Mandatory = false)]
    public IVisio.Shape[] Shape;

    protected override void ProcessRecord()
    {
        var targetshapes = new VisioScripting.TargetShapes(this.Shape);
        this.Client.Text.SetShapeText(targetshapes, this.Text);
    }
}