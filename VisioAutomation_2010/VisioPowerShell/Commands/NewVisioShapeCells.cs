using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioShapeCells)]
    public class NewVisioShapeCells : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var cells = new VisioPowerShell.Models.ShapeCells();
            this.WriteObject(cells);
        }
    }
}