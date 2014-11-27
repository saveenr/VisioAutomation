using VA = VisioAutomation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Format, "VisioText")]
    public class Format_VisioText : VisioCmdlet
    {

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        [SMA.Parameter(Mandatory = false)]
        [SMA.ValidateNotNullOrEmpty]
        public string  Font { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Togglecase { get; set; }

        protected override void ProcessRecord()
        {
            if (this.Font != null)
            {
                this.client.Text.SetFont(this.Shapes, Font);                
            }

            if (this.Togglecase)
            {
                this.client.Text.ToogleCase(this.Shapes);
            }
        }
    }
}