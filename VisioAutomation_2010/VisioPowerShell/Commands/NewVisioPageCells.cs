using System.Management.Automation;

namespace VisioPowerShell.Commands
{
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