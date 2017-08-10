using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioPageCells)]
    public class NewVisioPageCells : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var cells = new VisioAutomation.Models.Dom.PageCells();
            this.WriteObject(cells);
        }
    }
}