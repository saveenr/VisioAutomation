using System.Management.Automation;
using VisioAutomation.Scripting.Models;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Remove, VisioPowerShell.Nouns.VisioHyperlink)]
    public class Remove_VisioHyperlink : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        public int Index { get; set; }

        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var targets = new TargetShapes(this.Shapes);
            this.Client.Hyperlink.Delete(targets,this.Index);
        }
    }
}