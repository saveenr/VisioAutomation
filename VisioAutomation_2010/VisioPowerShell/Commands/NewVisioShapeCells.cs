using System.Management.Automation;

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
}