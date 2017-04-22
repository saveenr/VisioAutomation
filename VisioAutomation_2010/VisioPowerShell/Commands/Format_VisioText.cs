using System.Management.Automation;
using VisioScripting.Models;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Format, VisioPowerShell.Nouns.VisioShapeText)]
    public class Format_VisioText : VisioCmdlet
    {

        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        [Parameter(Mandatory = false)]
        [ValidateNotNullOrEmpty]
        public string  Font { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter Togglecase { get; set; }

        protected override void ProcessRecord()
        {
            var targets = new TargetShapes(this.Shapes);
            if (this.Font != null)
            {
                this.Client.Text.SetFont(targets, this.Font);                
            }

            if (this.Togglecase)
            {
                this.Client.Text.ToogleCase(targets);
            }
        }
    }
}