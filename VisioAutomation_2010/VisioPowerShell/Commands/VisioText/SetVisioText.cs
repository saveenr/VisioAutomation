using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, Nouns.VisioText)]
    public class SetVisioText : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string[] Text { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var targets = new VisioScripting.Models.TargetShapes(this.Shapes);
            this.Client.Text.SetShapeText(targets, this.Text);
        }
    }
}
