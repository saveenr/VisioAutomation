using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Format, VisioPowerShell.Commands.Nouns.VisioText)]
    public class FormatVisioShapeText : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        [Parameter(Mandatory = false)]
        [ValidateNotNullOrEmpty]
        public string  Font { get; set; }

        protected override void ProcessRecord()
        {
            var targets = new VisioScripting.Models.TargetShapes(this.Shapes);
            if (this.Font != null)
            {
                this.Client.Text.SetFont(targets, this.Font);                
            }
        }
    }
}