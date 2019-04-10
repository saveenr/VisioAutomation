using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsCommon.Copy, Nouns.VisioShape)]
    public class CopyVisioShape : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var targetselection = new VisioScripting.TargetSelection();
            this.Client.Selection.DuplicateShapes(targetselection);
        }
    }
}