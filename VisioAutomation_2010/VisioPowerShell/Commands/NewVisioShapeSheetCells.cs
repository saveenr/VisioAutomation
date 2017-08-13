using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioShapeSheetCells)]
    public class NewVisioShapeSheetCells : VisioCmdlet
    {
        [Parameter(Mandatory = true)]
        public CellsType Type { get; set; }

        protected override void ProcessRecord()
        {
            if (this.Type == CellsType.Page)
            {
                var pagecells = new VisioPowerShell.Models.PageCells();
                this.WriteObject(pagecells);

            }
            else if (this.Type == CellsType.ShapeFormat)
            {
                var shapecells = new VisioPowerShell.Models.ShapeCells();
                this.WriteObject(shapecells);

            }
            else if (this.Type == CellsType.TextFormat)
            {
                var textcells = new VisioPowerShell.Models.TextCells();
                this.WriteObject(textcells);
            }
            else
            {
                throw new System.ArgumentException();
            }
        }
    }
}