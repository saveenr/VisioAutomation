using System.Management.Automation;
using VisioPowerShell.Models;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioShapeCells)]
    public class NewVisioShapeCells : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var cells = new VisioPowerShell.Models.ShapeCells();
            this.WriteObject(cells);
        }
    }

    [Cmdlet(VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioPageCells)]
    public class NewVisioPageCells : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var cells = new VisioPowerShell.Models.PageCells();
            this.WriteObject(cells);
        }
    }
}