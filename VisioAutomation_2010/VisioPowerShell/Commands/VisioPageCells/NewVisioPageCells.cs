using System.Collections.Generic;
using System.Linq;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioPageCells
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, Nouns.VisioPageCells)]
    public class NewVisioPageCells : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public int Count = -1;
        protected override void ProcessRecord()
        {
            if (Count < 0)
            {
                var cells = new VisioPowerShell.Models.PageCells();
                this.WriteObject(cells);
            }
            else
            {
                var list_cells = new List<VisioPowerShell.Models.PageCells>(this.Count);
                var indices = Enumerable.Range(0, this.Count);
                var enum_cells = indices.Select(i => new VisioPowerShell.Models.PageCells());
                list_cells.AddRange(enum_cells);
                this.WriteObject(list_cells, false);
            }
        }
    }
}