using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioTextCells)]
    public class NewVisioTextCells : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var cells = new VisioPowerShell.Models.TextCells();
            this.WriteObject(cells);
        }
    }
}