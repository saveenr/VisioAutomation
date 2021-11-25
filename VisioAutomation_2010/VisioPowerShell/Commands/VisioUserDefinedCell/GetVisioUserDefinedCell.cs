

namespace VisioPowerShell.Commands.VisioUserDefinedCell;

[SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioUserDefinedCell)]
public class GetVisioUserDefinedCell : VisioCmdlet
{
    // CONTEXT:SHAPES 
    [SMA.Parameter(Mandatory = false)]
    public IVisio.Shape[] Shape;

    protected override void ProcessRecord()
    {
        var targetshapes = new VisioScripting.TargetShapes(this.Shape);
        var type = VASS.CellValueType.Formula;
        var dicof_shape_to_udcelldic = this.Client.UserDefinedCell.GetUserDefinedCellsAsShapeDictionary(targetshapes, type);

        this.WriteObject(dicof_shape_to_udcelldic);
    }
}