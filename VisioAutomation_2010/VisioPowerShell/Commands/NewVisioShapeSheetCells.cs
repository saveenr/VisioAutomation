using System.Management.Automation;
using VisioPowerShell.Models;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioShapeSheetCells)]
    public class NewVisioShapeSheetCells : VisioCmdlet
    {
        [Parameter(Mandatory = true)]
        public VisioPowerShell.Models.CellsType Type { get; set; }

        protected override void ProcessRecord()
        {
            var cells = BaseCells.CreateCells(this.Type);
            this.WriteObject(cells);
        }
    }
}