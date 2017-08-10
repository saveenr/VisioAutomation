using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioShapeCells)]
    public class NewVisioShapeCekks : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var cells = new VisioAutomation.Models.Dom.ShapeCells();
            this.WriteObject(cells);
        }
    }
}