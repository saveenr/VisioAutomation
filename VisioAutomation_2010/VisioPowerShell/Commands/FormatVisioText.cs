using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Format, VisioPowerShell.Commands.Nouns.VisioText)]
    public class FormatVisioText : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        [SMA.Parameter(Mandatory = false)]
        [SMA.ValidateNotNullOrEmpty]
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