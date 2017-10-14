using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, VisioPowerShell.Commands.Nouns.VisioGroup)]
    public class RemoveVisioGroup : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var targets = new VisioScripting.Models.TargetShapes(this.Shapes);
            this.Client.Grouping.UngroupSelectedShapes(targets);
        }
    }
}