using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, VisioPowerShell.Commands.Nouns.VisioHyperlink)]
    public class RemoveVisioHyperlink : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public int Index { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var targets = new VisioScripting.Models.TargetShapes(this.Shapes);
            this.Client.Hyperlink.DeleteHyperlinkAtIndex(targets,this.Index);
        }
    }
}