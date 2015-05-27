using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.Format
{
    [Cmdlet(VerbsCommon.Format, "VisioText")]
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
            if (this.Font != null)
            {
                this.client.Text.SetFont(this.Shapes, this.Font);                
            }

            if (this.Togglecase)
            {
                this.client.Text.ToogleCase(this.Shapes);
            }
        }
    }
}