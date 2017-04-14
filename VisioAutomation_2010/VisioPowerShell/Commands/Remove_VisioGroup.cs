using System.Management.Automation;
using VisioAutomation.Scripting.Models;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Remove, VisioPowerShell.Nouns.VisioGroup)]
    public class Remove_VisioGroup : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var targets = new TargetShapes(this.Shapes);
            this.Client.Grouping.Ungroup(targets);
        }
    }
}