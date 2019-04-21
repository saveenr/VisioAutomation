using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsCommon.Copy, Nouns.VisioShape)]
    public class CopyVisioShape : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.Client.Selection.DuplicateShapes(VisioScripting.TargetSelection.Auto);
        }
    }
}