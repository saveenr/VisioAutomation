using System.Management.Automation;
using VisioPowerShell.Models;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioCells)]
    public class NewVisioCells : VisioCmdlet
    {
        [Parameter(Mandatory = true)]
        public VisioPowerShell.Models.CellType Type { get; set; }

        protected override void ProcessRecord()
        {
            var cells = BaseCells.CreateCells(this.Type);
            this.WriteObject(cells);
        }
    }
}