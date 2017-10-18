using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioPageCells)]
    public class NewVisioPageCells : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var cells = new VisioPowerShell.Models.PageCells();
            this.WriteObject(cells);
        }
    }
}