using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, VisioPowerShell.Commands.Nouns.VisioText)]
    [System.Management.Automation.Alias("Set-VisioShapeText")]
    public class SetVisioText : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string[] Text { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var targets = new VisioScripting.Models.TargetShapes(this.Shapes);
            this.Client.Text.Set(targets, this.Text);
        }
    }
}
