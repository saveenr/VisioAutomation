namespace VisioPowerShell.Commands.VisioShapeCells;

[SMA.Cmdlet(SMA.VerbsCommon.New, Nouns.VisioShapeCells)]
public class NewVisioShapeCells : VisioCmdlet
{
    [SMA.Parameter(Mandatory = false)]
    public int Count=-1;
    protected override void ProcessRecord()
    {
        if (Count < 0)
        {
            var cells = new VisioPowerShell.Models.ShapeCells();
            this.WriteObject(cells);
        }
        else
        {
            var list_cells = new List<VisioPowerShell.Models.ShapeCells>(this.Count);
            var indices = Enumerable.Range(0, this.Count);
            var enum_cells = indices.Select(i => new VisioPowerShell.Models.ShapeCells());
            list_cells.AddRange(enum_cells);
            this.WriteObject(list_cells,false);
        }
    }
}