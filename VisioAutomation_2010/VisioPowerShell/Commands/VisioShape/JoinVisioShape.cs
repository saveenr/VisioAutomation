
namespace VisioPowerShell.Commands.VisioShape;

[SMA.Cmdlet(SMA.VerbsCommon.Join, Nouns.VisioShape)]
public class JoinVisioShape : VisioCmdlet
{
    // CONTEXT:SHAPESSELECTION
    [SMA.Parameter(Mandatory = false)]
    public IVisio.Shape[] Shape;

    protected override void ProcessRecord()
    {
        var targetshapes = new VisioScripting.TargetShapes(this.Shape);
        targetshapes.ResolveToSelection(this.Client);

        var group = this.Client.Grouping.Group(VisioScripting.TargetSelection.Auto);
        this.WriteObject(group);
    }
}