using System.Management.Automation;
using VisioScripting.Models;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Copy, VisioPowerShell.Nouns.VisioShape)]
    public class Copy_VisioShape : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var targets = new TargetShapes(this.Shapes);
            this.Client.Selection.Duplicate(targets);
        }
    }
}